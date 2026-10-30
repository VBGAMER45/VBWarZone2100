Attribute VB_Name = "textdraw_gdi"
' lib/ivis_opengl/textdraw.cpp, textdraw.h: the text, on GDI instead of
' FreeType, HarfBuzz and FriBiDi (docs/PLAN.md, "Renderer"). The structure is
' the C++'s: the fonts and their sizes (iV_TextInit, the CJK fonts on demand),
' the TextShaper (text split into runs by script and bidi level, each run
' shaped into glyphs at pen positions, the glyphs' bitmaps composed into one
' RGBA image whose RGB is the coverage and A its luma, the image offset from
' the pen's origin on the baseline), the point / pixel conversions, WzText's
' texture cache (WzText.cls), iV_FormatText's line breaking, the ellipsis.
' What GDI does instead:
' * FTFace (an FT_Face with its char size): a GDI font (CreateFontW) of the
'   font file, loaded through PhysFS into AddFontMemResourceEx (private to the
'   process). Its size metrics (ascender, descender) are GetTextMetricsW's.
'   A font weight (the CJK variable font's 'wght' axis) is the font's weight.
' * FT_Glyph_To_Bitmap with FT_RENDER_MODE_LCD: GetGlyphOutlineW's
'   GGO_GRAY8_BITMAP (hinted greyscale, 65 levels), laid out as the LCD
'   bitmap is (3 bytes a pixel), each subpixel the pixel's coverage, so the
'   composition below is the C++'s. Glyphs are cached per face by glyph index
'   (C++: an LRU cache of the FT_Glyphs).
' * hb_shape: GetCharacterPlacementW (glyph indices, ligatures, Arabic
'   shapes; per-character glyph indices if it fails); a glyph's advance is
'   its GDI advance (gmCellIncX), without kerning; pen positions are whole
'   pixels (C++: 26.6 fixed point, so no subpixel offsets).
' * hb_unicode_script: the scripts that matter here (the CJK ones pick the
'   CJK fonts; Arabic and Hebrew are right-to-left), the rest one "Latin".
' * FriBiDi's embedding levels (WZ_FRIBIDI_ENABLED, as the official build):
'   a simplified bidi algorithm (strong types, numbers, W1-W3, W7, N1, N2,
'   I1, I2; no explicit embeddings); the base direction is getBaseDirection's.
' Offsets and lengths of text are in code points (WzString's characters).
' The fonts' files are where an installed game has them (fonts/DejaVuSans.ttf),
' or (VB) where the source tree that data/ is has them (fonts/DejaVu/...).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As Long) As Long
Private Declare Function GetTextMetricsW Lib "gdi32" (ByVal hdc As Long, lptm As TEXTMETRICW) As Long
Private Declare Function GetTextFaceW Lib "gdi32" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFaceName As Long) As Long
Private Declare Function GetGlyphOutlineW Lib "gdi32" (ByVal hdc As Long, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As WIN_GLYPHMETRICS, ByVal cbBuffer As Long, ByVal lpvBuffer As Long, lpmat2 As MAT2) As Long
Private Declare Function GetCharacterPlacementW Lib "gdi32" (ByVal hdc As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal nMexExtent As Long, lpResults As GCP_RESULTSW, ByVal dwFlags As Long) As Long
Private Declare Function GetGlyphIndicesW Lib "gdi32" (ByVal hdc As Long, ByVal lpstr As Long, ByVal c As Long, ByVal pgi As Long, ByVal fl As Long) As Long
Private Declare Function AddFontMemResourceEx Lib "gdi32" (ByVal pbFont As Long, ByVal cbFont As Long, ByVal pdv As Long, pcFonts As Long) As Long
Private Declare Function RemoveFontMemResourceEx Lib "gdi32" (ByVal fh As Long) As Long
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long

Private Const GDI_ERROR As Long = &HFFFFFFFF
Private Const GGO_GRAY8_BITMAP As Long = 6
Private Const GGO_GLYPH_INDEX As Long = &H80
Private Const GCP_LIGATE As Long = &H20
Private Const GCP_GLYPHSHAPE As Long = &H10
Private Const FW_NORMAL As Long = 400
Private Const FW_BOLD As Long = 700
Private Const DEFAULT_CHARSET As Long = 1
Private Const OUT_TT_PRECIS As Long = 4
Private Const CLIP_DEFAULT_PRECIS As Long = 0
Private Const ANTIALIASED_QUALITY As Long = 4
Private Const DEFAULT_PITCH As Long = 0
Private Const SYSTEM_FONT As Long = 13

Private Type WIN_GLYPHMETRICS
    gmBlackBoxX As Long
    gmBlackBoxY As Long
    gmptGlyphOriginX As Long
    gmptGlyphOriginY As Long
    gmCellIncX As Integer
    gmCellIncY As Integer
End Type

Private Type FIXED_
    fract As Integer
    value As Integer
End Type

Private Type MAT2
    eM11 As FIXED_
    eM12 As FIXED_
    eM21 As FIXED_
    eM22 As FIXED_
End Type

Private Type TEXTMETRICW
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Integer
    tmLastChar As Integer
    tmDefaultChar As Integer
    tmBreakChar As Integer
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type GCP_RESULTSW
    lStructSize As Long
    lpOutString As Long
    lpOrder As Long
    lpDx As Long
    lpCaretPos As Long
    lpClass As Long
    lpGlyphs As Long
    nGlyphs As Long
    nMaxFit As Long
End Type

Private Const ASCII_SPACE As Long = 32
Private Const ASCII_NEWLINE As Long = 64        ' ('@')
Private Const ASCII_COLOURMODE As Long = 35     ' ('#')

' (VB) hb_script_t: the scripts this port tells apart (hb_unicode_script)
Public Const HB_SCRIPT_UNKNOWN As Long = 0
Public Const HB_SCRIPT_COMMON As Long = 1
Public Const HB_SCRIPT_INHERITED As Long = 2
Public Const HB_SCRIPT_LATIN As Long = 3        ' (VB) every left-to-right script but the CJK ones
Public Const HB_SCRIPT_ARABIC As Long = 4
Public Const HB_SCRIPT_HEBREW As Long = 5
Public Const HB_SCRIPT_HAN As Long = 6
Public Const HB_SCRIPT_BOPOMOFO As Long = 7
Public Const HB_SCRIPT_HANGUL As Long = 8
Public Const HB_SCRIPT_HIRAGANA As Long = 9
Public Const HB_SCRIPT_KATAKANA As Long = 10

' hb_direction_t
Private Const HB_DIRECTION_INVALID As Long = 0
Private Const HB_DIRECTION_LTR As Long = 4
Private Const HB_DIRECTION_RTL As Long = 5

' (VB) the bidi types the simplified algorithm uses (FriBidiCharType)
Private Const BIDI_L As Long = 0
Private Const BIDI_R As Long = 1
Private Const BIDI_AL As Long = 2
Private Const BIDI_EN As Long = 3
Private Const BIDI_AN As Long = 4
Private Const BIDI_NSM As Long = 5
Private Const BIDI_ON As Long = 6

' Contains the font color in the following order: red, green, blue, alpha
Private font_colour(0 To 3) As Single           ' = {1.f, 1.f, 1.f, 1.f}

Private horizScaleFactor_ As Single             ' float _horizScaleFactor = 1.0f;
Private vertScaleFactor_ As Single              ' float _vertScaleFactor = 1.0f;
Private textdrawGlobalsReady As Boolean         ' (VB) the initialisers have run

Private Type RasterizedGlyph
    buffer() As Byte
    pitch As Long
    width As Long
    height As Long
    bearing_x As Long
    bearing_y As Long
End Type

Private Type GlyphMetrics
    width As Long
    height As Long
    bearing_x As Long
    bearing_y As Long
End Type

Private bLoadedTextSystem As Boolean
' m_loadedFontDataCache: the font files loaded (AddFontMemResourceEx's handles; GDI keeps its own copy of the data)
Private loadedFontFileNames() As String
Private loadedFontHandles() As Long
Private loadedFontCount As Long

' FTFace: the GDI font of a font file at a char size
Private Type FTFace
    hFont As Long                       ' (VB) m_face, with FT_Set_Char_Size's size
    ascender As Long                    ' face->size->metrics.ascender >> 6 (pixels)
    descender As Long                   ' face->size->metrics.descender >> 6 (pixels, negative)
    glyphSlots() As Long                ' (VB) the glyph cache: each glyph index's slot in glyphStore (0 = not cached)
    hasGlyphSlots As Boolean
    loaded As Boolean                   ' (VB) the WZFontCollection's unique_ptr is set
End Type

' WZFontCollection's faces: baseFonts are 0 to 6, cjkFonts 7 to 13
Private Const FACE_REGULAR As Long = 0
Private Const FACE_REGULARBOLD As Long = 1
Private Const FACE_BOLD As Long = 2
Private Const FACE_MEDIUM As Long = 3
Private Const FACE_MEDIUMBOLD As Long = 4
Private Const FACE_SMALL As Long = 5
Private Const FACE_SMALLBOLD As Long = 6
Private Const CJK_FACES As Long = 7
Private faces(0 To 13) As FTFace
Private baseFonts As Boolean                    ' (VB) baseFonts != nullptr
Private cjkFonts As Boolean                     ' (VB) cjkFonts != nullptr
Private failedToLoadCJKFonts As Boolean

' FTCache: the glyphs, rasterised (glyphStore(0) is the empty glyph, {})
Private Type CachedGlyph
    glyph As RasterizedGlyph
    advance As Long                     ' (VB) the glyph's x_advance in pixels (HarfBuzz's)
End Type
Private glyphStore() As CachedGlyph
Private glyphStoreCount As Long
Private glyphCache As Boolean                   ' (VB) glyphCache != nullptr

Private textDC As Long                          ' (VB) the GDI context the faces are measured in
Private selectedFont As Long                    ' (VB) the font selected into textDC

Private Type HarfbuzzPosition
    codepoint As Long                   ' the glyph index
    cluster As Long
    penPosition As Vector2i             ' (VB) in pixels (C++: 26.6)
    face As Long                        ' FTFace& face: its index in faces
End Type

Private Type ShapingResult
    glyphes() As HarfbuzzPosition
    glyphes_size As Long
    x_advance As Long                   ' (VB) pixels (C++: 26.6)
    y_advance As Long
End Type

Private Type TextRun
    startOffset As Long
    endOffset As Long
    script As Long
    direction As Long
    fontFace As Long                    ' FTFace*: its index in faces
    ' the shaped glyphs (hb_buffer's glyph infos and positions)
    glyphCount As Long
    glyphInfos_codepoint() As Long
    glyphInfos_cluster() As Long        ' in code points from startOffset
    glyphPositions_x_advance() As Long  ' pixels
End Type

Private Type SplitTextRunsResult
    textRuns() As TextRun
    textRuns_size As Long
    baseDirection As Long               ' FriBidiParType: HB_DIRECTION_LTR or HB_DIRECTION_RTL (see getBaseDirection)
End Type

' (VB) text.toUtf32(), with where each code point is in the VB String
Private Type CodePoints
    s As String
    cp() As Long
    utf16Start() As Long                ' 0-based; utf16Start(size) = Len(s)
    size As Long
End Type

Public Type TextLayoutMetrics
    width As Long
    height As Long
End Type

Public Type RenderedText
    bitmap As iV_Image                  ' Nothing = nullptr
    offset_x As Long
    offset_y As Long
End Type

Public Type DrawTextResult
    text As RenderedText
    layoutMetrics As TextLayoutMetrics
End Type

Public Type WzTextRun
    startOffset As Long
    endOffset As Long
    rightToLeft As Boolean
End Type

Public Type TextLine
    text As String
    dimensions As Vector2i
    offset As Vector2i
End Type

' The base DPI used internally.
' Do not change this, or various layout in the game interface & menus will break.
Private Const DEFAULT_DPI As Single = 72!

' fontToEllipsisMap
Private ellipsisTexts(0 To 7) As WzText

Private Const CJK_FONT_PATH As String = "fonts/NotoSansCJK-VF.otf.ttc"

Private textureID As gfx_api_texture

' (VB) the C++ initialisers
Private Sub InitTextdrawGlobals()
    If textdrawGlobalsReady Then Exit Sub
    textdrawGlobalsReady = True
    font_colour(0) = 1!
    font_colour(1) = 1!
    font_colour(2) = 1!
    font_colour(3) = 1!
    horizScaleFactor_ = 1!
    vertScaleFactor_ = 1!
End Sub

'***************************************************************************
'
' Internal classes
'
'***************************************************************************

' (VB) where the font file is: the installed game's place, else the source tree's (data/fonts/DejaVu, data/fonts/Noto)
Private Function fontLoadPath(ByVal fileName As String) As String
    Dim base As String
    fontLoadPath = fileName
    If PHYSFS_exists(fileName) <> 0 Then Exit Function
    base = Mid$(fileName, InStrRev(fileName, "/") + 1)
    If Left$(base, 6) = "DejaVu" Then
        fontLoadPath = "fonts/DejaVu/" & base
    ElseIf Left$(base, 4) = "Noto" Then
        fontLoadPath = "fonts/Noto/" & base
    End If
End Function

' (VB) the family name GDI knows a font file's font by (FreeType opens the file itself)
Private Function fontFamilyName(ByVal fileName As String) As String
    If InStr(fileName, "DejaVuSans") > 0 Then
        fontFamilyName = "DejaVu Sans"
    ElseIf InStr(fileName, "NotoSansCJK") > 0 Then
        fontFamilyName = "Noto Sans CJK JP"     ' the collection's first font (FreeType's face index 0)
    End If
End Function

' returns the font data's index (1-based), 0 = nullptr
Private Function loadFontData(ByVal fileName As String) As Long
    Dim i As Long, data() As Byte, size As Long, h As Long, nFonts As Long
    For i = 1 To loadedFontCount
        If loadedFontFileNames(i) = fileName Then
            loadFontData = i
            Exit Function
        End If
    Next
    debug_ 130, LOG_WZ, "loadFontData", "Loading font data: %s", fileName
    If Not loadFileToBufferVector(fontLoadPath(fileName), data, size, True, False) Then
        debug_ 133, LOG_WZ, "loadFontData", "Failed to load font file: %s", fileName
        Exit Function
    End If
    ' (VB) the font, private to the process
    If size > 0 Then h = AddFontMemResourceEx(VarPtr(data(0)), size, 0, nFonts)
    If h = 0 Then
        debug_ 0, LOG_WZ, "loadFontData", "(VB) AddFontMemResourceEx failed: %s", fileName
        Exit Function
    End If
    loadedFontCount = loadedFontCount + 1
    ReDim Preserve loadedFontFileNames(1 To loadedFontCount)
    ReDim Preserve loadedFontHandles(1 To loadedFontCount)
    loadedFontFileNames(loadedFontCount) = fileName
    loadedFontHandles(loadedFontCount) = h
    loadFontData = loadedFontCount
End Function

Private Sub clearFontDataCache()
    Dim i As Long
    For i = 1 To loadedFontCount
        RemoveFontMemResourceEx loadedFontHandles(i)
    Next
    loadedFontCount = 0
    Erase loadedFontFileNames
    Erase loadedFontHandles
End Sub

' (VB) selects the face's font into textDC
Private Sub selectFace(ByVal face As Long)
    If textDC = 0 Then textDC = CreateCompatibleDC(0)
    If selectedFont <> faces(face).hFont Then
        SelectObject textDC, faces(face).hFont
        selectedFont = faces(face).hFont
    End If
End Sub

' FTFace(FT_Library &lib, const std::string &fileName, int32_t charSize, uint32_t horizDPI, uint32_t vertDPI, optional<uint16_t> fontWeight = nullopt)
' (VB) returns "" or the std::runtime_error's what() (C++ throws)
Private Function FTFace_construct(ByVal face As Long, ByVal fileName As String, ByVal charSize As Long, ByVal horizDPI As Long, ByVal vertDPI As Long, Optional ByVal has_fontWeight As Boolean = False, Optional ByVal fontWeight As Long = 0) As String
    Dim pFileData As Long, weight As Long, family As String, tm As TEXTMETRICW, faceName As String, n As Long
    FTFace_destroy face                 ' (VB) the unique_ptr's old face, if any
    pFileData = loadFontData(fileName)
    If pFileData = 0 Then
        FTFace_construct = va("Failed to load font file: %s", fileName)
        Exit Function
    End If
    ' FT_Set_Char_Size(m_face, 0, charSize, horizDPI, vertDPI): an em of charSize / 64 points at vertDPI
    ' (GDI has one resolution: the vertical one, the horizontal being the same in the game)
    If has_fontWeight Then
        ' set the desired font weight (GDI picks the variable font's instance of that weight)
        weight = fontWeight
    ElseIf InStr(fileName, "-Bold") > 0 Then
        weight = FW_BOLD
    Else
        weight = FW_NORMAL
    End If
    family = fontFamilyName(fileName)
    faces(face).hFont = CreateFontW(-Int(charSize / 64# * vertDPI / 72# + 0.5), 0, 0, 0, weight, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, DEFAULT_PITCH, StrPtr(family))
    If faces(face).hFont = 0 Then
        FTFace_construct = "Could not set character size"
        Exit Function
    End If
    selectFace face
    faceName = String$(64, 0)
    n = GetTextFaceW(textDC, 64, StrPtr(faceName))
    If n > 0 Then faceName = Left$(faceName, n - 1)
    If faceName <> family Then
        debug_ 0, LOG_WZ, "FTFace", "(VB) GDI gave %s for %s", faceName, fileName
    End If
    GetTextMetricsW textDC, tm
    faces(face).ascender = tm.tmAscent
    faces(face).descender = -tm.tmDescent
    faces(face).hasGlyphSlots = False
    faces(face).loaded = True
End Function

' ~FTFace
Private Sub FTFace_destroy(ByVal face As Long)
    If Not faces(face).loaded Then Exit Sub
    If textDC <> 0 And selectedFont = faces(face).hFont Then
        SelectObject textDC, GetStockObject(SYSTEM_FONT)   ' (VB) a selected font is not deleted
        selectedFont = 0
    End If
    DeleteObject faces(face).hFont
    faces(face).hFont = 0
    Erase faces(face).glyphSlots
    faces(face).hasGlyphSlots = False
    faces(face).loaded = False
End Sub

' (VB) the identity transformation GetGlyphOutlineW wants
Private Function identityMat2() As MAT2
    identityMat2.eM11.value = 1
    identityMat2.eM22.value = 1
End Function

' (VB) FT_Load_Glyph + FT_Get_Glyph + FT_Glyph_To_Bitmap of glyph index glyphIdx: the glyph's slot in
' glyphStore, rasterised once (-1 if GDI has no such glyph)
Private Function rasterizeGlyph(ByVal face As Long, ByVal glyphIdx As Long) As Long
    Dim gm As WIN_GLYPHMETRICS, m As MAT2, size As Long, gray() As Byte, grayPitch As Long
    Dim slot As Long, i As Long, j As Long, v As Long, k As Long
    If glyphStoreCount = 0 Then
        ReDim glyphStore(0 To 63)
        glyphStoreCount = 1             ' glyphStore(0): the empty glyph, {}
    End If
    If glyphIdx < 0 Or glyphIdx > &HFFFF& Then
        rasterizeGlyph = -1
        Exit Function
    End If
    If Not faces(face).hasGlyphSlots Then
        ReDim faces(face).glyphSlots(0 To &HFFFF&)
        faces(face).hasGlyphSlots = True
    End If
    slot = faces(face).glyphSlots(glyphIdx)
    If slot <> 0 Then
        rasterizeGlyph = slot
        Exit Function
    End If
    selectFace face
    m = identityMat2()
    size = GetGlyphOutlineW(textDC, glyphIdx, GGO_GRAY8_BITMAP Or GGO_GLYPH_INDEX, gm, 0, 0, m)
    If size = GDI_ERROR Then
        rasterizeGlyph = -1
        Exit Function
    End If
    If glyphStoreCount > UBound(glyphStore) Then ReDim Preserve glyphStore(0 To 2 * glyphStoreCount - 1)
    slot = glyphStoreCount
    glyphStoreCount = glyphStoreCount + 1
    With glyphStore(slot)
        .advance = gm.gmCellIncX
        If size = 0 Then
            ' a glyph without pixels (a space): FreeType's empty bitmap
            .glyph.width = 0
            .glyph.height = 0
            .glyph.pitch = 0
            .glyph.bearing_x = 0
            .glyph.bearing_y = 0
        Else
            ReDim gray(0 To size - 1)
            GetGlyphOutlineW textDC, glyphIdx, GGO_GRAY8_BITMAP Or GGO_GLYPH_INDEX, gm, size, VarPtr(gray(0)), m
            grayPitch = (gm.gmBlackBoxX + 3) And Not 3
            .glyph.width = gm.gmBlackBoxX
            .glyph.height = gm.gmBlackBoxY
            .glyph.pitch = 3 * gm.gmBlackBoxX
            .glyph.bearing_x = gm.gmptGlyphOriginX
            .glyph.bearing_y = gm.gmptGlyphOriginY
            ReDim .glyph.buffer(0 To .glyph.pitch * .glyph.height)
            For i = 0 To .glyph.height - 1
                For j = 0 To .glyph.width - 1
                    k = i * grayPitch + j
                    If k < size Then v = gray(k) Else v = 0
                    v = (v * 255 + 32) \ 64     ' 0..64 -> 0..255
                    If v > 255 Then v = 255
                    .glyph.buffer(i * .glyph.pitch + 3 * j) = v
                    .glyph.buffer(i * .glyph.pitch + 3 * j + 1) = v
                    .glyph.buffer(i * .glyph.pitch + 3 * j + 2) = v
                Next
            Next
        End If
    End With
    faces(face).glyphSlots(glyphIdx) = slot
    rasterizeGlyph = slot
End Function

' FTFace::getGlyphWidth (the glyph_index in the font file is the codePoint, as the C++ has it)
' (VB) returns pixels: metrics.width >> 6
Private Function FTFace_getGlyphWidth(ByVal face As Long, ByVal codePoint As Long) As Long
    Dim slot As Long
    slot = rasterizeGlyph(face, codePoint)
    If Not (slot >= 0) Then
        ASSERT_FAILURE "error == FT_Err_Ok", "lib/ivis_opengl/textdraw.cpp:287", "getGlyphWidth", "Unable to load glyph for %u", codePoint
        Exit Function
    End If
    FTFace_getGlyphWidth = glyphStore(slot).glyph.width
End Function

' FTFace::getGlyph: (VB) the slot, 0 = nullptr
Private Function FTFace_getGlyph(ByVal face As Long, ByVal codePoint As Long) As Long
    Dim slot As Long
    slot = rasterizeGlyph(face, codePoint)
    If Not (slot >= 0) Then ASSERT_FAILURE "error == FT_Err_Ok", "lib/ivis_opengl/textdraw.cpp:297", "getGlyph", "Unable to load glyph %u", codePoint: Exit Function
    FTFace_getGlyph = slot
End Function

' FTCache::get: (VB) the slot of the rasterised glyph (0 = {})
Private Function FTCache_get(ByVal face As Long, ByVal codePoint As Long) As Long
    Dim glyph As Long
    glyph = FTCache_getGlyph(face, codePoint)
    If Not (glyph <> 0) Then ASSERT_FAILURE "glyph != nullptr", "lib/ivis_opengl/textdraw.cpp:387", "get", "Failed to get glyph: %u", codePoint: Exit Function
    FTCache_get = glyph
End Function

Private Function FTCache_getGlyphMetrics(ByVal face As Long, ByVal codePoint As Long) As GlyphMetrics
    Dim glyph As Long
    glyph = FTCache_getGlyph(face, codePoint)
    If Not (glyph <> 0) Then ASSERT_FAILURE "glyph != nullptr", "lib/ivis_opengl/textdraw.cpp:423", "getGlyphMetrics", "Failed to get glyph: %u", codePoint: Exit Function
    With glyphStore(glyph).glyph
        FTCache_getGlyphMetrics.width = .width
        FTCache_getGlyphMetrics.height = .height
        FTCache_getGlyphMetrics.bearing_x = .bearing_x
        FTCache_getGlyphMetrics.bearing_y = .bearing_y
    End With
End Function

Private Sub FTCache_clear()
    Dim f As Long
    Erase glyphStore
    glyphStoreCount = 0
    For f = 0 To 13
        Erase faces(f).glyphSlots
        faces(f).hasGlyphSlots = False
    Next
End Sub

' The glyph is owned by the cache
Private Function FTCache_getGlyph(ByVal face As Long, ByVal codepoint As Long) As Long
    ' (VB) cached or loaded fresh: rasterizeGlyph keeps it
    FTCache_getGlyph = FTFace_getGlyph(face, codepoint)
End Function

'---------------------------------------------------------------------------
' (VB) hb_unicode_script: Unicode's Scripts.txt for the scripts told apart here
Public Function hb_unicode_script(ByVal cp As Long) As Long
    Select Case cp
    Case &H41& To &H5A&, &H61& To &H7A&, &HAA&, &HBA&, &HC0& To &HD6&, &HD8& To &HF6&, &HF8& To &H2B8&, &H2E0& To &H2E4&
        hb_unicode_script = HB_SCRIPT_LATIN
    Case &H0& To &H40&, &H5B& To &H60&, &H7B& To &HA9&, &HAB& To &HB9&, &HBB& To &HBF&, &HD7&, &HF7&, &H2B9& To &H2DF&, &H2E5& To &H2E9&, &H2EC& To &H2FF&
        hb_unicode_script = HB_SCRIPT_COMMON
    Case &H2EA&, &H2EB&
        hb_unicode_script = HB_SCRIPT_BOPOMOFO
    Case &H300& To &H36F&, &H485&, &H486&, &H64B& To &H655&, &H670&, &H1AB0& To &H1AFF&, &H1DC0& To &H1DFF&, &H200C&, &H200D&, &H20D0& To &H20FF&, &H302A& To &H302D&, &H3099&, &H309A&, &HFE00& To &HFE0F&, &HFE20& To &HFE2F&
        hb_unicode_script = HB_SCRIPT_INHERITED
    Case &H591& To &H5FF&, &HFB1D& To &HFB4F&
        hb_unicode_script = HB_SCRIPT_HEBREW
    Case &H605&, &H60C&, &H61B&, &H61F&, &H640&, &H6DD&
        hb_unicode_script = HB_SCRIPT_COMMON
    Case &H600& To &H6FF&, &H750& To &H77F&, &H8A0& To &H8FF&, &HFB50& To &HFDFF&, &HFE70& To &HFEFC&
        hb_unicode_script = HB_SCRIPT_ARABIC
    Case &H1100& To &H11FF&, &H302E&, &H302F&, &H3131& To &H318E&, &H3200& To &H321E&, &H3260& To &H327E&, &HA960& To &HA97F&, &HAC00& To &HD7FF&, &HFFA0& To &HFFDC&
        hb_unicode_script = HB_SCRIPT_HANGUL
    Case &H2000& To &H200B&, &H200E& To &H2064&, &H2066& To &H2070&, &H2074& To &H207E&, &H2080& To &H208E&, &H20A0& To &H20CF&, &H2100& To &H2125&, &H2127& To &H2129&, &H212C& To &H2131&, &H2133& To &H214D&, &H214F& To &H215F&, &H2189& To &H2BFF&, &H2E00& To &H2E7F&
        hb_unicode_script = HB_SCRIPT_COMMON
    Case &H2E80& To &H2FDF&, &H3005&, &H3007&, &H3021& To &H3029&, &H3038& To &H303B&, &H3400& To &H4DBF&, &H4E00& To &H9FFF&, &HF900& To &HFAFF&, &H20000 To &H3FFFF
        hb_unicode_script = HB_SCRIPT_HAN
    Case &H3000& To &H3004&, &H3006&, &H3008& To &H3020&, &H3030& To &H3037&, &H303C& To &H303F&, &H309B&, &H309C&, &H30A0&, &H30FB&, &H30FC&, &H3190& To &H319F&, &H31C0& To &H31E3&, &H3220& To &H325F&, &H327F& To &H32CF&, &H32FF&, &H3358& To &H33FF&, &HFE10& To &HFE1F&, &HFE30& To &HFE6F&, &HFEFF&, &HFF01& To &HFF20&, &HFF3B& To &HFF40&, &HFF5B& To &HFF65&, &HFF70&, &HFF9E&, &HFF9F&, &HFFE0& To &HFFFD&, &H1F000 To &H1FAFF
        hb_unicode_script = HB_SCRIPT_COMMON
    Case &H3041& To &H3096&, &H309D& To &H309F&, &H1B001 To &H1B11F
        hb_unicode_script = HB_SCRIPT_HIRAGANA
    Case &H30A1& To &H30FA&, &H30FD& To &H30FF&, &H31F0& To &H31FF&, &H32D0& To &H32FE&, &H3300& To &H3357&, &HFF66& To &HFF6F&, &HFF71& To &HFF9D&
        hb_unicode_script = HB_SCRIPT_KATAKANA
    Case &H3105& To &H312F&, &H31A0& To &H31BF&
        hb_unicode_script = HB_SCRIPT_BOPOMOFO
    Case Else
        hb_unicode_script = HB_SCRIPT_LATIN
    End Select
End Function

' (VB) hb_script_get_horizontal_direction
Private Function hb_script_get_horizontal_direction(ByVal script As Long) As Long
    Select Case script
    Case HB_SCRIPT_ARABIC, HB_SCRIPT_HEBREW
        hb_script_get_horizontal_direction = HB_DIRECTION_RTL
    Case HB_SCRIPT_COMMON, HB_SCRIPT_INHERITED, HB_SCRIPT_UNKNOWN
        hb_script_get_horizontal_direction = HB_DIRECTION_INVALID
    Case Else
        hb_script_get_horizontal_direction = HB_DIRECTION_LTR
    End Select
End Function

' (VB) fribidi_get_bidi_types, of the types the simplified algorithm uses
Private Function fribidi_get_bidi_type(ByVal cp As Long) As Long
    Select Case cp
    Case &H30& To &H39&, &H6F0& To &H6F9&, &HFF10& To &HFF19&
        fribidi_get_bidi_type = BIDI_EN
    Case &H660& To &H669&, &H66B&, &H66C&
        fribidi_get_bidi_type = BIDI_AN
    Case &H200E&
        fribidi_get_bidi_type = BIDI_L
    Case &H200F&
        fribidi_get_bidi_type = BIDI_R
    Case &H61C&
        fribidi_get_bidi_type = BIDI_AL
    Case Else
        Select Case hb_unicode_script(cp)
        Case HB_SCRIPT_ARABIC
            fribidi_get_bidi_type = BIDI_AL
        Case HB_SCRIPT_HEBREW
            fribidi_get_bidi_type = BIDI_R
        Case HB_SCRIPT_INHERITED
            fribidi_get_bidi_type = BIDI_NSM
        Case HB_SCRIPT_COMMON
            fribidi_get_bidi_type = BIDI_ON
        Case Else
            fribidi_get_bidi_type = BIDI_L
        End Select
    End Select
End Function

' (VB) fribidi_get_par_embedding_levels_ex, simplified: the embedding level of each code point in a
' paragraph of base direction baseDirection (no explicit embeddings, isolates or separators)
Private Sub fribidi_get_par_embedding_levels(t As CodePoints, ByVal baseDirection As Long, levels() As Long)
    Dim types() As Long, i As Long, j As Long, k As Long, baseLevel As Long, sos As Long, prev As Long
    Dim lastStrong As Long, before As Long, after As Long, e As Long
    If baseDirection = HB_DIRECTION_RTL Then baseLevel = 1 Else baseLevel = 0
    If baseLevel = 1 Then sos = BIDI_R Else sos = BIDI_L
    ReDim types(0 To t.size - 1)
    For i = 0 To t.size - 1
        types(i) = fribidi_get_bidi_type(t.cp(i))
    Next
    ' W1: a non-spacing mark takes the type of the character before it
    prev = sos
    For i = 0 To t.size - 1
        If types(i) = BIDI_NSM Then types(i) = prev Else prev = types(i)
    Next
    ' W2: a European number after an Arabic letter (back to the last strong type) is an Arabic number; W3: AL is R
    lastStrong = sos
    For i = 0 To t.size - 1
        Select Case types(i)
        Case BIDI_L, BIDI_R
            lastStrong = types(i)
        Case BIDI_AL
            lastStrong = BIDI_AL
            types(i) = BIDI_R
        Case BIDI_EN
            If lastStrong = BIDI_AL Then types(i) = BIDI_AN
        End Select
    Next
    ' W7: a European number after L (back to the last strong type) is L
    lastStrong = sos
    For i = 0 To t.size - 1
        Select Case types(i)
        Case BIDI_L, BIDI_R
            lastStrong = types(i)
        Case BIDI_EN
            If lastStrong = BIDI_L Then types(i) = BIDI_L
        End Select
    Next
    ' N1, N2: neutrals between strong types of one direction take it (numbers count as R), others the embedding direction
    i = 0
    Do While i < t.size
        If types(i) = BIDI_ON Then
            j = i
            Do While j < t.size
                If types(j) <> BIDI_ON Then Exit Do
                j = j + 1
            Loop
            If i = 0 Then before = sos Else before = types(i - 1)
            If before <> BIDI_L Then before = BIDI_R
            If j = t.size Then after = sos Else after = types(j)
            If after <> BIDI_L Then after = BIDI_R
            If before = after Then e = before Else e = sos
            For k = i To j - 1
                types(k) = e
            Next
            i = j
        Else
            i = i + 1
        End If
    Loop
    ' I1, I2
    ReDim levels(0 To t.size - 1)
    For i = 0 To t.size - 1
        levels(i) = baseLevel
        If baseLevel = 0 Then
            Select Case types(i)
            Case BIDI_R
                levels(i) = baseLevel + 1
            Case BIDI_AN, BIDI_EN
                levels(i) = baseLevel + 2
            End Select
        Else
            Select Case types(i)
            Case BIDI_L, BIDI_EN, BIDI_AN
                levels(i) = baseLevel + 1
            End Select
        End If
    Next
End Sub

' (VB) WzString::toUtf32(), keeping each code point's place in the String
Private Sub toUtf32(ByVal s As String, t As CodePoints)
    Dim i As Long, n As Long, c As Long, c2 As Long
    t.s = s
    n = Len(s)
    t.size = 0
    ReDim t.cp(0 To n)
    ReDim t.utf16Start(0 To n)
    i = 1
    Do While i <= n
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        t.utf16Start(t.size) = i - 1
        If c >= &HD800& And c <= &HDBFF& And i < n Then
            c2 = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If c2 >= &HDC00& And c2 <= &HDFFF& Then
                c = &H10000 + (c - &HD800&) * &H400& + (c2 - &HDC00&)
                i = i + 1
            End If
        End If
        t.cp(t.size) = c
        t.size = t.size + 1
        i = i + 1
    Loop
    t.utf16Start(t.size) = n
End Sub

'---------------------------------------------------------------------------
' TextShaper

' Note:
' Technically glyph antialiasing is dependent of text rotation.
' Rotated text needs to set transform inside freetype2.
' However there is few rotated text in wz2100 and it's likely to make
' only minimal visual difference.

' Returns the maximum text run length (in WzString characters) that fits within a max width (supplied *IN PIXELS*)
Private Function TextShaper_getTextMaxLenForWidth(ByVal text As String, ByVal fontID As iV_fonts, ByVal maxWidthInPixels As Long, ByVal rightToLeft As Boolean) As Long
    Dim t As CodePoints, shapingResult_ As ShapingResult, glyph As GlyphMetrics
    Dim min_x As Long, max_x As Long, min_y As Long, max_y As Long, x0 As Long, y0 As Long
    Dim glyphIdx As Long, glyphIdxEnd As Long, texture_width As Long, truncatedCluster As Long
    toUtf32 text, t
    shapingResult_ = TextShaper_shapeText(t, fontID)

    If shapingResult_.glyphes_size = 0 Then
        TextShaper_getTextMaxLenForWidth = 0
        Exit Function
    End If

    min_x = &H7FFFFFFF
    max_x = &H80000000
    min_y = &H7FFFFFFF
    max_y = &H80000000

    If Not rightToLeft Then glyphIdx = 0 Else glyphIdx = shapingResult_.glyphes_size
    If Not rightToLeft Then glyphIdxEnd = shapingResult_.glyphes_size Else glyphIdxEnd = 0
    Do While glyphIdx <> glyphIdxEnd
        If rightToLeft Then glyphIdx = glyphIdx - 1

        With shapingResult_.glyphes(glyphIdx)
            glyph = FTCache_getGlyphMetrics(.face, .codepoint)
            x0 = .penPosition.x + glyph.bearing_x
            y0 = .penPosition.y - glyph.bearing_y
        End With

        If x0 < min_x Then min_x = x0
        If x0 + glyph.width > max_x Then max_x = x0 + glyph.width
        If y0 < min_y Then min_y = y0
        If y0 + glyph.height > max_y Then max_y = y0 + glyph.height

        texture_width = max_x - min_x + 1
        If ULT(maxWidthInPixels, texture_width) Then
            Exit Do
        End If

        If Not rightToLeft Then glyphIdx = glyphIdx + 1
    Loop

    If glyphIdx < shapingResult_.glyphes_size Then
        If glyphIdx = 0 Then
            If Not rightToLeft Then
                texture_width = max_x - min_x + 1
                If ULT(maxWidthInPixels, texture_width) Then ' first glyph doesn't even fit
                    TextShaper_getTextMaxLenForWidth = 0
                    Exit Function
                End If
            Else
                ' whole thing fits
                TextShaper_getTextMaxLenForWidth = t.size
                Exit Function
            End If
        End If

        ' must truncate to fit
        ' we have a cluster index which does not fit, which is an index into the codepoints
        truncatedCluster = shapingResult_.glyphes(glyphIdx).cluster
        TextShaper_getTextMaxLenForWidth = truncatedCluster
        Exit Function
    End If

    TextShaper_getTextMaxLenForWidth = t.size
End Function

' Returns the text width and height *IN PIXELS*
Private Function TextShaper_getTextMetrics(ByVal text As String, ByVal fontID As iV_fonts) As TextLayoutMetrics
    Dim t As CodePoints, shapingResult_ As ShapingResult, glyph As GlyphMetrics, i As Long
    Dim min_x As Long, max_x As Long, min_y As Long, max_y As Long, x0 As Long, y0 As Long
    Dim texture_width As Long, texture_height As Long, x_advance As Long, y_advance As Long
    toUtf32 text, t
    shapingResult_ = TextShaper_shapeText(t, fontID)

    If shapingResult_.glyphes_size = 0 Then
        TextShaper_getTextMetrics.width = shapingResult_.x_advance
        TextShaper_getTextMetrics.height = shapingResult_.y_advance
        Exit Function
    End If

    ' std::accumulate from (1000, -1000, 1000, -1000)
    min_x = 1000
    max_x = -1000
    min_y = 1000
    max_y = -1000
    For i = 0 To shapingResult_.glyphes_size - 1
        With shapingResult_.glyphes(i)
            glyph = FTCache_getGlyphMetrics(.face, .codepoint)
            x0 = .penPosition.x + glyph.bearing_x
            y0 = .penPosition.y - glyph.bearing_y
        End With
        If x0 < min_x Then min_x = x0
        If x0 + glyph.width > max_x Then max_x = x0 + glyph.width
        If y0 < min_y Then min_y = y0
        If y0 + glyph.height > max_y Then max_y = y0 + glyph.height
    Next

    texture_width = max_x - min_x + 1
    texture_height = max_y - min_y + 1
    x_advance = shapingResult_.x_advance
    y_advance = shapingResult_.y_advance

    ' return the maximum of the x_advance / y_advance (converted from harfbuzz units) and the texture dimensions
    If ULT(texture_width, x_advance) Then TextShaper_getTextMetrics.width = x_advance Else TextShaper_getTextMetrics.width = texture_width
    If ULT(texture_height, y_advance) Then TextShaper_getTextMetrics.height = y_advance Else TextShaper_getTextMetrics.height = texture_height
End Function

' FriBidiParType getBaseDirection()
Private Function TextShaper_getBaseDirection() As Long
    Dim language As String
    language = getLanguage()

    If language = "ar_SA" Then
        TextShaper_getBaseDirection = HB_DIRECTION_RTL
    Else
        TextShaper_getBaseDirection = HB_DIRECTION_LTR
    End If
End Function

' Draws the text and returns the text buffer, width and height, etc *IN PIXELS*
Public Function TextShaper_drawText(ByVal text As String, ByVal fontID As iV_fonts) As DrawTextResult
    Dim t As CodePoints, shapingResult_ As ShapingResult, i As Long, j As Long, n As Long
    Dim min_x As Long, max_x As Long, min_y As Long, max_y As Long, x0 As Long, y0 As Long
    Dim texture_width As Long, texture_height As Long, x_advance As Long, y_advance As Long
    Dim glyphSlot() As Long, glyphPixelX() As Long, glyphPixelY() As Long ' the glyphRaster structs
    Dim stringBitmap As iV_Image, stringTexture() As Byte, stringTextureSize As Long
    Dim glyphNum As Long, glyphBufferSize As Long, i0 As Long, j0 As Long, srcBufferPos As Long, stringTexturePos As Long
    Dim s0 As Long, s1 As Long, s2 As Long, d As Long
    toUtf32 text, t
    shapingResult_ = TextShaper_shapeText(t, fontID)

    If shapingResult_.glyphes_size = 0 Then
        TextShaper_drawText.layoutMetrics.width = shapingResult_.x_advance
        TextShaper_drawText.layoutMetrics.height = shapingResult_.y_advance
        Exit Function
    End If

    min_x = 1000
    max_x = -1000
    min_y = 1000
    max_y = -1000

    ' build glyphes
    n = shapingResult_.glyphes_size
    ReDim glyphSlot(0 To n - 1)
    ReDim glyphPixelX(0 To n - 1)
    ReDim glyphPixelY(0 To n - 1)
    For i = 0 To n - 1
        With shapingResult_.glyphes(i)
            glyphSlot(i) = FTCache_get(.face, .codepoint)
            x0 = .penPosition.x + glyphStore(glyphSlot(i)).glyph.bearing_x
            y0 = .penPosition.y - glyphStore(glyphSlot(i)).glyph.bearing_y
        End With
        If x0 < min_x Then min_x = x0
        If x0 + glyphStore(glyphSlot(i)).glyph.width > max_x Then max_x = x0 + glyphStore(glyphSlot(i)).glyph.width
        If y0 < min_y Then min_y = y0
        If y0 + glyphStore(glyphSlot(i)).glyph.height > max_y Then max_y = y0 + glyphStore(glyphSlot(i)).glyph.height
        glyphPixelX(i) = x0
        glyphPixelY(i) = y0
    Next

    texture_width = max_x - min_x + 1
    texture_height = max_y - min_y + 1
    x_advance = shapingResult_.x_advance
    y_advance = shapingResult_.y_advance

    If texture_width = 0 Or texture_height = 0 Then
        TextShaper_drawText.layoutMetrics.width = x_advance
        TextShaper_drawText.layoutMetrics.height = y_advance
        Exit Function
    End If

    Set stringBitmap = New iV_Image
    stringBitmap.allocate texture_width, texture_height, 4, True
    stringTextureSize = stringBitmap.data_size()
    ' (VB) composed here, then copied to the image's bmp_w()
    ReDim stringTexture(0 To stringTextureSize - 1)

    ' TODO: Someone should document this piece.
    glyphNum = 0
    For d = 0 To n - 1
        With glyphStore(glyphSlot(d)).glyph
            glyphBufferSize = .pitch * .height
            For i = 0 To .height - 1
                i0 = glyphPixelY(d) - min_y
                For j = 0 To .width - 1
                    j0 = glyphPixelX(d) - min_x
                    srcBufferPos = i * .pitch + 3 * j
                    If Not (srcBufferPos + 2 < glyphBufferSize) Then ASSERT_FAILURE "srcBufferPos + 2 < glyphBufferSize", "lib/ivis_opengl/textdraw.cpp:843", "drawText", "Invalid source (%u / %u) reading glyph %zu for string ""%s""; (%d, %d, %d, %d, %u, %d, %d, %d, %u, %u)", srcBufferPos, glyphBufferSize, glyphNum, text, i, .height, glyphPixelY(d), min_y, i0, j, glyphPixelX(d), min_x, j0, .pitch
                    s0 = .buffer(srcBufferPos)
                    s1 = .buffer(srcBufferPos + 1)
                    s2 = .buffer(srcBufferPos + 2)
                    stringTexturePos = 4 * ((i0 + i) * texture_width + j + j0)
                    If Not (stringTexturePos + 3 < stringTextureSize) Then ASSERT_FAILURE "stringTexturePos + 3 < stringTextureSize", "lib/ivis_opengl/textdraw.cpp:846", "drawText", "Invalid destination (%u / %zu) writing glyph %zu for string ""%s""; (%d, %d, %d, %d, %u, %d, %d, %d, %u, %u)", stringTexturePos, stringTextureSize, glyphNum, text, i, .height, glyphPixelY(d), min_y, i0, j, glyphPixelX(d), min_x, j0, texture_width
                    stringTexture(stringTexturePos) = min_(stringTexture(stringTexturePos) + s0, 255)
                    stringTexture(stringTexturePos + 1) = min_(stringTexture(stringTexturePos + 1) + s1, 255)
                    stringTexture(stringTexturePos + 2) = min_(stringTexture(stringTexturePos + 2) + s2, 255)
                    stringTexture(stringTexturePos + 3) = min_(stringTexture(stringTexturePos + 3) + ((s0 * 77 + s1 * 150 + s2 * 29) \ 256), 255)
                Next
            Next
        End With
        glyphNum = glyphNum + 1
    Next
    CopyMemory ByVal stringBitmap.bmp_w(), stringTexture(0), stringTextureSize

    Set TextShaper_drawText.text.bitmap = stringBitmap
    TextShaper_drawText.text.offset_x = min_x
    TextShaper_drawText.text.offset_y = min_y
    If ULT(texture_width, x_advance) Then TextShaper_drawText.layoutMetrics.width = x_advance Else TextShaper_drawText.layoutMetrics.width = texture_width
    If ULT(texture_height, y_advance) Then TextShaper_drawText.layoutMetrics.height = y_advance Else TextShaper_drawText.layoutMetrics.height = texture_height
End Function

' (VB) std::min of two ints
Private Function min_(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then min_ = a Else min_ = b
End Function

Private Function TextShaper_splitTextRuns(t As CodePoints, ByVal fontID As iV_fonts) As SplitTextRunsResult
    Dim codePoints_size As Long, scripts() As Long, levels() As Long, baseDirection As Long, i As Long, j As Long
    Dim lastScriptValue As Long, lastScriptIndex As Long, lastSetIndex As Long
    Dim result As SplitTextRunsResult, lastScript As Long, lastLevel As Long, lastRunStart As Long
    Dim startOffset As Long, endOffset As Long, script As Long, direction As Long, face As Long, newRun As Boolean
    If t.size = 0 Then
        Exit Function
    End If
    codePoints_size = t.size

    ReDim scripts(0 To codePoints_size - 1)

    ' Step 1: Initialize fribidi variables.

    baseDirection = TextShaper_getBaseDirection()

    ' Step 2: Run fribidi.

    ' Get the bidi type of each character in the string, and the embedding levels
    fribidi_get_par_embedding_levels t, baseDirection, levels

    ' Fill the array of scripts with scripts of each character
    For i = 0 To codePoints_size - 1
        scripts(i) = hb_unicode_script(t.cp(i))
    Next

    ' Step 3: Resolve common or inherited scripts.

    lastScriptValue = HB_SCRIPT_UNKNOWN
    lastScriptIndex = -1
    lastSetIndex = -1

    For i = 0 To codePoints_size - 1
        If scripts(i) = HB_SCRIPT_COMMON Or scripts(i) = HB_SCRIPT_INHERITED Then
            If lastScriptIndex <> -1 Then
                scripts(i) = lastScriptValue
                lastSetIndex = i
            End If
        Else
            For j = lastSetIndex + 1 To i - 1
                scripts(j) = scripts(i)
            Next
            lastScriptValue = scripts(i)
            lastScriptIndex = i
            lastSetIndex = i
        End If
    Next

    ' Step 4: Create the different runs

    lastScript = scripts(0)
    lastLevel = levels(0)
    lastRunStart = 0 ' where the last run started

    ' i == size means that we've reached the end of the string,
    ' and that the last run should be created.
    For i = 0 To codePoints_size
        ' If the script or level is of the current point is the same as the previous one,
        ' then this means that the we have not reached the end of the current run.
        ' If there's change, create a new run.
        If i = codePoints_size Then
            newRun = True
        ElseIf scripts(i) <> lastScript Or levels(i) <> lastLevel Then
            newRun = True
        Else
            newRun = False
        End If
        If newRun Then
            startOffset = lastRunStart
            endOffset = i
            script = lastScript

            ' "lastLevel & 1" yields either 1 or 0, depending on the least significant bit of lastLevel.
            If (lastLevel And 1) <> 0 Then direction = HB_DIRECTION_RTL Else direction = HB_DIRECTION_LTR

            face = getFTFace(fontID, script)

            If result.textRuns_size = 0 Then ReDim result.textRuns(0 To 7)
            If result.textRuns_size > UBound(result.textRuns) Then ReDim Preserve result.textRuns(0 To 2 * result.textRuns_size - 1)
            With result.textRuns(result.textRuns_size)
                .startOffset = startOffset
                .endOffset = endOffset
                .script = script
                .direction = direction
                .fontFace = face
                .glyphCount = 0
            End With
            result.textRuns_size = result.textRuns_size + 1

            If i < codePoints_size Then
                lastScript = scripts(i)
                lastLevel = levels(i)
                lastRunStart = i
            Else
                Exit For
            End If
        End If
    Next

    result.baseDirection = baseDirection
    TextShaper_splitTextRuns = result
End Function

Private Function TextShaper_shapeText(t As CodePoints, ByVal fontID As iV_fonts) As ShapingResult
    Dim textRunResult As SplitTextRunsResult, baseDirection As Long, shapingResult_ As ShapingResult
    Dim i As Long, x As Long, y As Long
    If t.size = 0 Then
        Exit Function
    End If

    textRunResult = TextShaper_splitTextRuns(t, fontID)
    baseDirection = textRunResult.baseDirection

    ' Step 6: Shape each run using harfbuzz.

    For i = 0 To textRunResult.textRuns_size - 1
        TextShaper_shapeHarfbuzz textRunResult.textRuns(i), t
    Next

    x = 0
    y = 0

    ' The direction of the loop must change depending on the base direction
    If Not (baseDirection = HB_DIRECTION_RTL) Then
        For i = 0 To textRunResult.textRuns_size - 1
            processTextRunGlyphs textRunResult.textRuns(i), shapingResult_, x, y
        Next
    Else
        For i = textRunResult.textRuns_size - 1 To 0 Step -1
            processTextRunGlyphs textRunResult.textRuns(i), shapingResult_, x, y
        Next
    End If

    shapingResult_.x_advance = shapingResult_.x_advance + x
    shapingResult_.y_advance = shapingResult_.y_advance + y

    ' Step 7: Finalize.

    TextShaper_shapeText = shapingResult_
End Function

' shapeText's lambda processTextRunGlyphs
Private Sub processTextRunGlyphs(run As TextRun, shapingResult_ As ShapingResult, x As Long, y As Long)
    Dim glyphIndex As Long, clusterIdx As Long
    For glyphIndex = 0 To run.glyphCount - 1
        clusterIdx = run.glyphInfos_cluster(glyphIndex) + run.startOffset
        If shapingResult_.glyphes_size = 0 Then ReDim shapingResult_.glyphes(0 To 15)
        If shapingResult_.glyphes_size > UBound(shapingResult_.glyphes) Then ReDim Preserve shapingResult_.glyphes(0 To 2 * shapingResult_.glyphes_size - 1)
        With shapingResult_.glyphes(shapingResult_.glyphes_size)
            .codepoint = run.glyphInfos_codepoint(glyphIndex)
            .cluster = clusterIdx
            .penPosition.x = x              ' + x_offset (GDI's are 0)
            .penPosition.y = y
            .face = run.fontFace
        End With
        shapingResult_.glyphes_size = shapingResult_.glyphes_size + 1

        x = x + run.glyphPositions_x_advance(glyphIndex)
        ' y += y_advance (0 for horizontal text)
    Next
End Sub

' shapeHarfbuzz: (VB) GetCharacterPlacementW shapes the run in its font
Private Sub TextShaper_shapeHarfbuzz(run As TextRun, t As CodePoints)
    Dim units As Long, startUnit As Long, sub_ As String, gcp As GCP_RESULTSW, r As Long
    Dim glyphs() As Integer, order() As Long, cl() As Long, cpOfUnit() As Long
    Dim g As Long, k As Long, n As Long, slot As Long, tmp As Long
    startUnit = t.utf16Start(run.startOffset)
    units = t.utf16Start(run.endOffset) - startUnit
    If units <= 0 Then Exit Sub
    sub_ = Mid$(t.s, startUnit + 1, units)
    ' the code point (from startOffset) of each UTF-16 unit of the run
    ReDim cpOfUnit(0 To units - 1)
    For k = run.startOffset To run.endOffset - 1
        For g = t.utf16Start(k) To t.utf16Start(k + 1) - 1
            cpOfUnit(g - startUnit) = k - run.startOffset
        Next
    Next
    selectFace run.fontFace
    ReDim glyphs(0 To 2 * units + 7)
    ReDim order(0 To units - 1)
    For k = 0 To units - 1
        order(k) = -1
    Next
    gcp.lStructSize = LenB(gcp)
    gcp.lpOrder = VarPtr(order(0))
    gcp.lpGlyphs = VarPtr(glyphs(0))
    gcp.nGlyphs = UBound(glyphs) + 1
    r = GetCharacterPlacementW(textDC, StrPtr(sub_), units, 0, gcp, GCP_LIGATE Or GCP_GLYPHSHAPE)
    If r <> 0 And gcp.nGlyphs > 0 Then
        n = gcp.nGlyphs
        ReDim run.glyphInfos_codepoint(0 To n - 1)
        ReDim run.glyphInfos_cluster(0 To n - 1)
        ReDim cl(0 To n - 1)
        For g = 0 To n - 1
            cl(g) = -1
        Next
        ' a glyph's cluster: the first character that maps to it
        For k = 0 To units - 1
            g = order(k)
            If g >= 0 And g < n Then
                If cl(g) = -1 Then cl(g) = cpOfUnit(k)
            End If
        Next
        For g = 0 To n - 1
            run.glyphInfos_codepoint(g) = glyphs(g) And &HFFFF&
            If cl(g) = -1 Then
                ' no lpOrder: one glyph a character
                If g < units Then cl(g) = cpOfUnit(g) Else cl(g) = cpOfUnit(units - 1)
            End If
            run.glyphInfos_cluster(g) = cl(g)
        Next
    Else
        ' per character (BMP; another plane's characters are the missing glyph)
        n = run.endOffset - run.startOffset
        ReDim run.glyphInfos_codepoint(0 To n - 1)
        ReDim run.glyphInfos_cluster(0 To n - 1)
        For k = 0 To n - 1
            If t.cp(run.startOffset + k) <= &HFFFF& Then
                GetGlyphIndicesW textDC, StrPtr(sub_) + 2 * (t.utf16Start(run.startOffset + k) - startUnit), 1, VarPtr(glyphs(0)), 0
                run.glyphInfos_codepoint(k) = glyphs(0) And &HFFFF&
            Else
                run.glyphInfos_codepoint(k) = 0
            End If
            run.glyphInfos_cluster(k) = k
        Next
    End If
    ReDim run.glyphPositions_x_advance(0 To n - 1)
    For g = 0 To n - 1
        slot = rasterizeGlyph(run.fontFace, run.glyphInfos_codepoint(g))
        If slot > 0 Then run.glyphPositions_x_advance(g) = glyphStore(slot).advance
    Next
    ' a right-to-left run's glyphs in visual order, as hb_shape gives them
    If run.direction = HB_DIRECTION_RTL Then
        For g = 0 To n \ 2 - 1
            tmp = run.glyphInfos_codepoint(g): run.glyphInfos_codepoint(g) = run.glyphInfos_codepoint(n - 1 - g): run.glyphInfos_codepoint(n - 1 - g) = tmp
            tmp = run.glyphInfos_cluster(g): run.glyphInfos_cluster(g) = run.glyphInfos_cluster(n - 1 - g): run.glyphInfos_cluster(n - 1 - g) = tmp
            tmp = run.glyphPositions_x_advance(g): run.glyphPositions_x_advance(g) = run.glyphPositions_x_advance(n - 1 - g): run.glyphPositions_x_advance(n - 1 - g) = tmp
        Next
    End If
    run.glyphCount = n
End Sub

'***************************************************************************/
'
' Main source
'
'***************************************************************************/

Public Sub iV_font(ByVal fontName As String, ByVal fontFace As String, ByVal fontFaceBold As String)
End Sub

Public Function iV_GetHorizScaleFactor() As Single
    InitTextdrawGlobals
    iV_GetHorizScaleFactor = horizScaleFactor_
End Function

Public Function iV_GetVertScaleFactor() As Single
    InitTextdrawGlobals
    iV_GetVertScaleFactor = vertScaleFactor_
End Function

Private Function initializeCJKFontsIfNeeded() As Boolean
    Dim horizDPI As Long, vertDPI As Long, e As String, f As Long
    If cjkFonts Then initializeCJKFontsIfNeeded = True: Exit Function
    If failedToLoadCJKFonts Then initializeCJKFontsIfNeeded = False: Exit Function
    cjkFonts = True
    horizDPI = Fix(DEFAULT_DPI * horizScaleFactor_)
    vertDPI = Fix(DEFAULT_DPI * vertScaleFactor_)
    e = FTFace_construct(CJK_FACES + FACE_REGULAR, CJK_FONT_PATH, 12 * 64, horizDPI, vertDPI, True, 400)
    If e = "" Then e = FTFace_construct(CJK_FACES + FACE_REGULARBOLD, CJK_FONT_PATH, 12 * 64, horizDPI, vertDPI, True, 700)
    If e = "" Then e = FTFace_construct(CJK_FACES + FACE_BOLD, CJK_FONT_PATH, 21 * 64, horizDPI, vertDPI, True, 400)
    If e = "" Then e = FTFace_construct(CJK_FACES + FACE_MEDIUM, CJK_FONT_PATH, 16 * 64, horizDPI, vertDPI, True, 400)
    If e = "" Then e = FTFace_construct(CJK_FACES + FACE_MEDIUMBOLD, CJK_FONT_PATH, 16 * 64, horizDPI, vertDPI, True, 700)
    If e = "" Then e = FTFace_construct(CJK_FACES + FACE_SMALL, CJK_FONT_PATH, 9 * 64, horizDPI, vertDPI, True, 400)
    If e = "" Then e = FTFace_construct(CJK_FACES + FACE_SMALLBOLD, CJK_FONT_PATH, 9 * 64, horizDPI, vertDPI, True, 700)
    If e <> "" Then
        ' catch (const std::exception &e)
        debug_ 1194, LOG_ERROR, "initializeCJKFontsIfNeeded", "Failed to load font:\n%s", e
        For f = CJK_FACES To CJK_FACES + FACE_SMALLBOLD
            FTFace_destroy f
        Next
        cjkFonts = False
        failedToLoadCJKFonts = True
        initializeCJKFontsIfNeeded = False
        Exit Function
    End If

    initializeCJKFontsIfNeeded = True
End Function

' static FTFace &getFTFace(iV_fonts FontID, hb_script_t script): (VB) the face's index
Public Function getFTFace(ByVal FontID As iV_fonts, ByVal script As Long) As Long
    Select Case script
    Case HB_SCRIPT_HAN, HB_SCRIPT_BOPOMOFO, HB_SCRIPT_HANGUL, HB_SCRIPT_HIRAGANA, HB_SCRIPT_KATAKANA
        If initializeCJKFontsIfNeeded() Then
            Select Case FontID
            Case font_regular_bold
                getFTFace = CJK_FACES + FACE_REGULARBOLD
            Case font_large
                getFTFace = CJK_FACES + FACE_BOLD
            Case font_medium
                getFTFace = CJK_FACES + FACE_MEDIUM
            Case font_medium_bold
                getFTFace = CJK_FACES + FACE_MEDIUMBOLD
            Case font_small
                getFTFace = CJK_FACES + FACE_SMALL
            Case font_bar
                getFTFace = CJK_FACES + FACE_SMALLBOLD
            Case Else ' font_regular
                getFTFace = CJK_FACES + FACE_REGULAR
            End Select
            Exit Function
        End If
    End Select
    Select Case FontID
    Case font_regular_bold
        getFTFace = FACE_REGULARBOLD
    Case font_large
        getFTFace = FACE_BOLD
    Case font_medium
        getFTFace = FACE_MEDIUM
    Case font_medium_bold
        getFTFace = FACE_MEDIUMBOLD
    Case font_small
        getFTFace = FACE_SMALL
    Case font_bar
        getFTFace = FACE_SMALLBOLD
    Case Else ' font_regular
        getFTFace = FACE_REGULAR
    End Select
End Function

' (VB) face->size->metrics.ascender >> 6 and descender >> 6, in pixels
Public Function FTFace_ascender(ByVal face As Long) As Long
    FTFace_ascender = faces(face).ascender
End Function

Public Function FTFace_descender(ByVal face As Long) As Long
    FTFace_descender = faces(face).descender
End Function

Public Sub iV_TextInit(ByVal horizScalePercentage As Long, ByVal vertScalePercentage As Long)
    Dim horizScaleFactor As Single, vertScaleFactor As Single, horizDPI As Long, vertDPI As Long, e As String
    Dim files() As String, count As Long, i As Long, path As String, n As Long
    InitTextdrawGlobals
    If horizScalePercentage > 100 And horizScalePercentage < 200 Then
        horizScalePercentage = horizScalePercentage * 2
    End If
    If vertScalePercentage > 100 And vertScalePercentage < 200 Then
        vertScalePercentage = vertScalePercentage * 2
    End If
    horizScaleFactor = horizScalePercentage / 100!
    vertScaleFactor = vertScalePercentage / 100!
    ' assert(horizScaleFactor >= 1.0f); assert(vertScaleFactor >= 1.0f): a release build's asserts (none)

    ' Use the scaling factors to multiply the default DPI (72) to determine the desired internal font rendering DPI.
    horizScaleFactor_ = horizScaleFactor
    vertScaleFactor_ = vertScaleFactor
    horizDPI = Fix(DEFAULT_DPI * horizScaleFactor)
    vertDPI = Fix(DEFAULT_DPI * vertScaleFactor)
    debug_ 1281, LOG_WZ, "iV_TextInit", "Text-Rendering Scaling Factor: %f x %f; Internal Font DPI: %u x %u", CDbl(horizScaleFactor_), CDbl(vertScaleFactor_), horizDPI, vertDPI

    If Not glyphCache Then
        glyphCache = True
    End If

    If Not baseFonts Then
        baseFonts = True
    End If

    If textDC = 0 Then textDC = CreateCompatibleDC(0)  ' (VB) GDI's context for the faces

    e = FTFace_construct(FACE_REGULAR, "fonts/DejaVuSans.ttf", 12 * 64, horizDPI, vertDPI)
    If e = "" Then e = FTFace_construct(FACE_REGULARBOLD, "fonts/DejaVuSans-Bold.ttf", 12 * 64, horizDPI, vertDPI)
    If e = "" Then e = FTFace_construct(FACE_BOLD, "fonts/DejaVuSans-Bold.ttf", 21 * 64, horizDPI, vertDPI)
    If e = "" Then e = FTFace_construct(FACE_MEDIUM, "fonts/DejaVuSans.ttf", 16 * 64, horizDPI, vertDPI)
    If e = "" Then e = FTFace_construct(FACE_MEDIUMBOLD, "fonts/DejaVuSans-Bold.ttf", 16 * 64, horizDPI, vertDPI)
    If e = "" Then e = FTFace_construct(FACE_SMALL, "fonts/DejaVuSans.ttf", 9 * 64, horizDPI, vertDPI)
    If e = "" Then e = FTFace_construct(FACE_SMALLBOLD, "fonts/DejaVuSans-Bold.ttf", 9 * 64, horizDPI, vertDPI)
    If e <> "" Then
        ' catch (const std::exception &e)
        ' Log lots of details:
        debugOutputSearchPaths
        debug_ 1305, LOG_INFO, "iV_TextInit", "Virtual filesystem:"
        If WZ_PHYSFS_enumerateFiles("", files, count) Then
            For i = 0 To count - 1
                debug_ 1308, LOG_INFO, "iV_TextInit", " - %s", files(i)
            Next
        End If
        ' LaunchInfo::getCurrentProcessDetails().imageFileName.fullPath(): (VB) the process's image
        path = String$(1024, 0)
        n = GetModuleFileNameW(0, StrPtr(path), 1024)
        path = Left$(path, n)
        debug_ 1311, LOG_INFO, "iV_TextInit", "Path to executable: %s", path
        debugOutputSearchPathMountErrors
        ' then...
        debug_ 1314, LOG_FATAL, "iV_TextInit", "Failed to load base font:\n%s\n", e
        exit_ 3                         ' abort()
    End If

    ' Do a sanity-check here to make sure the CJK font exists
    ' (since it's only loaded on-demand, and thus might fail with a fatal error later if missing)
    If PHYSFS_exists(fontLoadPath(CJK_FONT_PATH)) = 0 Then
        debug_ 1323, LOG_FATAL, "iV_TextInit", "Missing data file: %s", CJK_FONT_PATH
    End If

    ' m_unicode_funcs_hb = hb_unicode_funcs_get_default(); hb_language_get_default(): (VB) nothing to set up

    bLoadedTextSystem = True
End Sub

Public Sub iV_TextShutdown()
    Dim f As Long
    FTCache_clear
    glyphCache = False
    For f = 0 To CJK_FACES - 1
        FTFace_destroy f
    Next
    baseFonts = False
    For f = CJK_FACES To CJK_FACES + FACE_SMALLBOLD
        FTFace_destroy f
    Next
    cjkFonts = False
    Set textureID = Nothing
    For f = 0 To 7
        Set ellipsisTexts(f) = Nothing
    Next
    clearFontDataCache
    If textDC <> 0 Then
        DeleteDC textDC                 ' (VB)
        textDC = 0
        selectedFont = 0
    End If
    bLoadedTextSystem = False
End Sub

Public Sub iV_TextUpdateScaleFactor(ByVal horizScalePercentage As Long, ByVal vertScalePercentage As Long)
    If Not bLoadedTextSystem Then
        Exit Sub
    End If
    iV_TextShutdown
    iV_TextInit horizScalePercentage, vertScalePercentage
End Sub

Private Function iV_Internal_GetEllipsis(ByVal fontID As iV_fonts) As WzText
    If ellipsisTexts(fontID) Is Nothing Then
        ' We must create + cache an ellipsis for this fontID
        Set ellipsisTexts(fontID) = WzText_new(ChrW$(&H2026), fontID)
    End If
    Set iV_Internal_GetEllipsis = ellipsisTexts(fontID)
End Function

Public Function iV_GetEllipsisWidth(ByVal fontID As iV_fonts) As Long
    iV_GetEllipsisWidth = iV_Internal_GetEllipsis(fontID).width()
End Function

' position: Vector2f (x, y)
Public Sub iV_DrawEllipsis(ByVal fontID As iV_fonts, ByVal x As Single, ByVal y As Single, ByVal colour As Long, Optional ByVal rotation_ As Single = 0!)
    iV_Internal_GetEllipsis(fontID).render x, y, colour, rotation_
End Sub

Public Function width_pixelsToPoints(ByVal widthInPixels As Long) As Long
    InitTextdrawGlobals
    width_pixelsToPoints = -Int(-(CSng(U32(widthInPixels)) / horizScaleFactor_))   ' ceil
End Function

Public Function height_pixelsToPoints(ByVal heightInPixels As Long) As Long
    InitTextdrawGlobals
    height_pixelsToPoints = -Int(-(CSng(U32(heightInPixels)) / vertScaleFactor_))    ' ceil
End Function

' Returns the text width *in points*
Public Function iV_GetTextWidth(ByVal string_ As String, ByVal fontID As iV_fonts) As Long
    Dim metrics As TextLayoutMetrics
    metrics = TextShaper_getTextMetrics(string_, fontID)
    iV_GetTextWidth = width_pixelsToPoints(metrics.width)
End Function

' Note: Is intended to be used *only* with text runs produced by iV_SplitTextParagraphIntoRuns() - rightToLeft must be properly supplied
Public Function iV_GetMaxTextRunLenForWidth(ByVal textRun_ As String, ByVal fontID As iV_fonts, ByVal maxWidthInPoints As Long, ByVal rightToLeft As Boolean) As Long
    Dim scaledWidthInPixels As Single, maxWidthInPixels As Long
    InitTextdrawGlobals
    scaledWidthInPixels = Int(CSng(U32(maxWidthInPoints)) * horizScaleFactor_)
    ' (VB) a Long holds up to 2^31 - 1 (C++: maxUint32Float, then uint32_t's max)
    If scaledWidthInPixels <= 2147483520! Then maxWidthInPixels = CLng(scaledWidthInPixels) Else maxWidthInPixels = &H7FFFFFFF
    iV_GetMaxTextRunLenForWidth = TextShaper_getTextMaxLenForWidth(textRun_, fontID, maxWidthInPixels, rightToLeft)
End Function

' Returns the counted text width *in points*
Public Function iV_GetCountedTextWidth(ByVal string_ As String, ByVal string_length As Long, ByVal fontID As iV_fonts) As Long
    iV_GetCountedTextWidth = iV_GetTextWidth(string_, fontID)
End Function

' Returns the text height *in points*
Public Function iV_GetTextHeight(ByVal string_ As String, ByVal fontID As iV_fonts) As Long
    Dim metrics As TextLayoutMetrics
    metrics = TextShaper_getTextMetrics(string_, fontID)
    iV_GetTextHeight = height_pixelsToPoints(metrics.height)
End Function

' std::vector<WzTextRun>: result(0 .. count - 1); returns the count
Public Function iV_SplitTextParagraphIntoRuns(ByVal string_ As String, ByVal fontID As iV_fonts, result() As WzTextRun) As Long
    Dim t As CodePoints, internalResult As SplitTextRunsResult, i As Long, count As Long
    toUtf32 string_, t
    internalResult = TextShaper_splitTextRuns(t, fontID)
    If internalResult.textRuns_size = 0 Then Exit Function
    ReDim result(0 To internalResult.textRuns_size - 1)

    ' The direction of the loop must change depending on the base direction
    If Not (internalResult.baseDirection = HB_DIRECTION_RTL) Then
        For i = 0 To internalResult.textRuns_size - 1
            processTextRun internalResult.textRuns(i), result(count)
            count = count + 1
        Next
    Else
        For i = internalResult.textRuns_size - 1 To 0 Step -1
            processTextRun internalResult.textRuns(i), result(count)
            count = count + 1
        Next
    End If

    iV_SplitTextParagraphIntoRuns = count
End Function

' iV_SplitTextParagraphIntoRuns's lambda processTextRun
Private Sub processTextRun(run As TextRun, r As WzTextRun)
    r.startOffset = run.startOffset
    r.endOffset = run.endOffset
    r.rightToLeft = (run.direction = HB_DIRECTION_RTL)
End Sub

' (VB) iV_SplitTextParagraphIntoRuns for the other projects: the runs' fields in arrays
Public Function iV_SplitTextParagraphIntoRuns_arrays(ByVal string_ As String, ByVal fontID As iV_fonts, startOffsets() As Long, endOffsets() As Long, rightToLeft() As Boolean) As Long
    Dim runs() As WzTextRun, n As Long, i As Long
    n = iV_SplitTextParagraphIntoRuns(string_, fontID, runs)
    If n > 0 Then
        ReDim startOffsets(0 To n - 1)
        ReDim endOffsets(0 To n - 1)
        ReDim rightToLeft(0 To n - 1)
        For i = 0 To n - 1
            startOffsets(i) = runs(i).startOffset
            endOffsets(i) = runs(i).endOffset
            rightToLeft(i) = runs(i).rightToLeft
        Next
    End If
    iV_SplitTextParagraphIntoRuns_arrays = n
End Function

' Returns the character width *in points*
Public Function iV_GetCharWidth(ByVal charCode As Long, ByVal fontID As iV_fonts) As Long
    iV_GetCharWidth = width_pixelsToPoints(FTFace_getGlyphWidth(getFTFace(fontID, hb_unicode_script(charCode)), charCode))
End Function

Public Function metricsHeight_PixelsToPoints(ByVal heightMetric As Long) As Long
    Dim ptMetric As Single
    InitTextdrawGlobals
    ptMetric = CSng(heightMetric) / vertScaleFactor_
    If ptMetric < 0 Then metricsHeight_PixelsToPoints = Int(ptMetric) Else metricsHeight_PixelsToPoints = -Int(-ptMetric)
End Function

Public Function iV_GetTextLineSize(ByVal fontID As iV_fonts) As Long
    Dim face As Long
    face = getFTFace(fontID, HB_SCRIPT_COMMON) ' TODO: Better handling of script-specific font faces?
    iV_GetTextLineSize = metricsHeight_PixelsToPoints(faces(face).ascender - faces(face).descender)
End Function

Public Function iV_GetTextAboveBase(ByVal fontID As iV_fonts) As Long
    Dim face As Long
    face = getFTFace(fontID, HB_SCRIPT_COMMON) ' TODO: Better handling of script-specific font faces?
    iV_GetTextAboveBase = metricsHeight_PixelsToPoints(-faces(face).ascender)
End Function

Public Function iV_GetTextBelowBase(ByVal fontID As iV_fonts) As Long
    Dim face As Long
    face = getFTFace(fontID, HB_SCRIPT_COMMON) ' TODO: Better handling of script-specific font faces?
    iV_GetTextBelowBase = metricsHeight_PixelsToPoints(faces(face).descender)
End Function

' colour: a PIELIGHT (r in the low byte)
Public Sub iV_SetTextColour(ByVal colour As Long)
    InitTextdrawGlobals
    font_colour(0) = (colour And &HFF&) / 255!
    font_colour(1) = ((colour And &HFF00&) \ &H100&) / 255!
    font_colour(2) = ((colour And &HFF0000) \ &H10000) / 255!
    font_colour(3) = SHRU(colour, 24) / 255!
End Sub

' optional<iV_fonts>: returns has_value, the font in result
Public Function iV_FontModifyBold(ByVal fontID As iV_fonts, ByVal bold As Boolean, result As iV_fonts) As Boolean
    If bold Then
        Select Case fontID
            ' bold fonts
            Case font_large, font_medium_bold, font_regular_bold, font_bar ' font_large is actually bold; font_bar is small_bold
                result = fontID ' already bolded
                iV_FontModifyBold = True

            ' regular fonts
            Case font_medium
                result = font_medium_bold
                iV_FontModifyBold = True
            Case font_scaled, font_regular ' font_scaled is treated the same as font_regular
                result = font_regular_bold
                iV_FontModifyBold = True
            Case font_small
                result = font_bar
                iV_FontModifyBold = True

            Case font_count
                iV_FontModifyBold = False
        End Select
    Else
        Select Case fontID
            ' bold fonts
            Case font_large ' is actually bold
                iV_FontModifyBold = False ' no non-bold option for this size
            Case font_medium_bold
                result = font_medium
                iV_FontModifyBold = True
            Case font_regular_bold
                result = font_regular
                iV_FontModifyBold = True
            Case font_bar ' small_bold
                result = font_small
                iV_FontModifyBold = True

            ' regular fonts
            Case font_medium, font_scaled, font_regular, font_small ' font_scaled is treated the same as font_regular
                result = fontID ' already non-bolded
                iV_FontModifyBold = True

            Case font_count
                iV_FontModifyBold = False
        End Select
    End If
End Function

' optional<iV_fonts>: returns has_value, the font in result
Public Function iV_ShrinkFont(ByVal fontID As iV_fonts, result As iV_fonts) As Boolean
    Select Case fontID
        ' bold fonts
        Case font_large ' is actually bold
            result = font_medium_bold
            iV_ShrinkFont = True
        Case font_medium_bold
            result = font_regular_bold
            iV_ShrinkFont = True
        Case font_regular_bold
            result = font_bar ' small_bold
            iV_ShrinkFont = True
        Case font_bar ' small_bold
            iV_ShrinkFont = False

        ' regular fonts
        Case font_medium
            result = font_regular
            iV_ShrinkFont = True
        Case font_scaled, font_regular ' font_scaled is treated the same as font_regular
            result = font_small
            iV_ShrinkFont = True
        Case font_small
            iV_ShrinkFont = False

        Case font_count
            iV_ShrinkFont = False
    End Select
End Function

Private Function breaksLine(ByVal c As Long) As Boolean
    breaksLine = (c = ASCII_NEWLINE Or c = 10)
End Function

Private Function breaksWord(ByVal c As Long) As Boolean
    breaksWord = (c = ASCII_SPACE Or breaksLine(c))
End Function

' utf8SequenceLength: (VB) the text is UTF-16 here: 2 for a surrogate pair's first unit, else 1
Private Function utf16SequenceLength(ByVal first As Long) As Long
    If first >= &HD800& And first <= &HDBFF& Then
        utf16SequenceLength = 2
    Else
        utf16SequenceLength = 1
    End If
End Function

' std::vector<TextLine>: lineDrawResults(0 .. count - 1); returns the count.
' (VB) The C++ walks the UTF-8 bytes; this walks the UTF-16 units (the same for ASCII; a
' surrogate pair is one sequence, as a multi-byte character is).
Public Function iV_FormatText(ByVal String_ As String, ByVal MaxWidth As Long, ByVal Justify As Long, ByVal fontID As iV_fonts, ByVal ignoreNewlines As Boolean, lineDrawResults() As TextLine) As Long
    Dim count As Long
    Dim FString As String
    Dim FWord As String
    Const x As Long = 0
    Const y As Long = 0
    Dim i As Long
    Dim jx As Long                      ' Default to left justify.
    Dim jy As Long
    Dim WWidth As Long
    Dim TWidth As Long
    Dim curChar As Long                 ' (VB) the 1-based index of *curChar in String_; charEnd is past the end
    Dim charEnd As Long
    Dim sequenceLen As Long
    Dim GotSpace As Boolean, NewLine As Boolean, indexWithinLine As Long, startOfWord As Long
    Dim FStringWidth As Long, c As Long, spaceWidth As Long

    jx = x
    jy = y
    curChar = 1
    charEnd = Len(String_) + 1
    sequenceLen = 1

    Do While curChar < charEnd
        GotSpace = False
        NewLine = False

        ' Reset text draw buffer
        FString = ""

        WWidth = 0

        indexWithinLine = 0

        ' Parse through the string, adding words until width is achieved.
        Do While curChar < charEnd And (WWidth = 0 Or ULT(WWidth, MaxWidth)) And Not NewLine
            startOfWord = curChar
            FStringWidth = iV_GetTextWidth(FString, fontID)

            ' Get the next word.
            i = 0
            FWord = ""
            Do While curChar < charEnd
                c = AscW(Mid$(String_, curChar, 1)) And &HFFFF&
                If Not ((indexWithinLine = 0 And Not breaksLine(c)) Or Not breaksWord(c)) Then Exit Do
                sequenceLen = utf16SequenceLength(c)
                If sequenceLen = 0 Then
                    ASSERT_FAILURE "false", "lib/ivis_opengl/textdraw.cpp:1643", "iV_FormatText", "curr_sequence_length is 0?? for string: %s", String_
                    sequenceLen = 1
                End If
                If curChar + sequenceLen > charEnd Then
                    sequenceLen = (charEnd - curChar)
                End If

                If c = ASCII_COLOURMODE Then ' If it's a colour mode toggle char then just add it to the word.
                    FWord = FWord & ChrW$(c)

                    ' this character won't be drawn so don't deal with its width
                    GoTo continue_
                End If

                ' Get the full (possibly multi-byte) sequence for this codepoint
                FWord = FWord & Mid$(String_, curChar, sequenceLen)

                ' Update this line's pixel width.
                WWidth = FStringWidth + iV_GetTextWidth(FWord, fontID)

                ' If this word doesn't fit on the current line then break out
                If indexWithinLine <> 0 And ULT(MaxWidth, WWidth) Then
                    FWord = Left$(FWord, Len(FWord) - 1)
                    Exit Do
                End If
continue_:
                i = i + sequenceLen
                curChar = curChar + sequenceLen
                indexWithinLine = indexWithinLine + sequenceLen
            Loop

            If curChar < charEnd Then c = AscW(Mid$(String_, curChar, 1)) And &HFFFF& Else c = 0
            ' Don't forget the space.
            If c = ASCII_SPACE Then
                FWord = FWord & " "
                i = i + 1
                curChar = curChar + 1
                GotSpace = True
                spaceWidth = iV_GetCharWidth(ASCII_SPACE, fontID)
                If Not ULT(MaxWidth, WWidth + spaceWidth) Then
                    WWidth = WWidth + spaceWidth
                End If
            ' Check for new line character.
            ElseIf breaksLine(c) Then
                If Not ignoreNewlines Then
                    NewLine = True
                End If
                curChar = curChar + 1
            End If

            ' If we've passed a space on this line and the word goes past the
            ' maximum width and this isn't caused by the appended space then
            ' rewind to the start of this word and finish this line.
            If GotSpace And i <> 0 Then
                If ULT(MaxWidth, WWidth) And Mid$(FWord, i, 1) <> " " Then
                    ' Skip back to the beginning of this
                    ' word and draw it on the next line
                    curChar = startOfWord
                    Exit Do
                End If
            End If

            ' And add it to the output string.
            FString = FString & FWord
        Loop

        ' Remove trailing spaces, useful when doing center alignment.
        Do While Right$(FString, 1) = " "
            FString = Left$(FString, Len(FString) - 1)  ' std::string has no pop_back().
        Loop

        TWidth = iV_GetTextWidth(FString, fontID)

        ' Do justify.
        Select Case Justify
        Case FTEXT_CENTRE
            jx = x + CLng(Fix(U32(MaxWidth - TWidth) / 2))   ' (UDWORD - int) / 2
        Case FTEXT_RIGHTJUSTIFY
            jx = x + MaxWidth - TWidth
        Case FTEXT_LEFTJUSTIFY
            jx = x
        End Select

        ' Store the line of text and its position in the bounding rect
        If count = 0 Then ReDim lineDrawResults(0 To 7)
        If count > UBound(lineDrawResults) Then ReDim Preserve lineDrawResults(0 To 2 * count - 1)
        lineDrawResults(count).text = FString
        lineDrawResults(count).dimensions.x = TWidth
        lineDrawResults(count).dimensions.y = iV_GetTextLineSize(fontID)
        lineDrawResults(count).offset.x = jx
        lineDrawResults(count).offset.y = jy
        count = count + 1

        ' and move down a line.
        jy = jy + iV_GetTextLineSize(fontID)
    Loop

    iV_FormatText = count
End Function

' (VB) iV_FormatText for the other projects: the lines' fields in arrays
Public Function iV_FormatText_arrays(ByVal String_ As String, ByVal MaxWidth As Long, ByVal Justify As Long, ByVal fontID As iV_fonts, ByVal ignoreNewlines As Boolean, lineText() As String, lineWidth() As Long, lineHeight() As Long, lineOffsetX() As Long, lineOffsetY() As Long) As Long
    Dim lines() As TextLine, n As Long, i As Long
    n = iV_FormatText(String_, MaxWidth, Justify, fontID, ignoreNewlines, lines)
    If n > 0 Then
        ReDim lineText(0 To n - 1)
        ReDim lineWidth(0 To n - 1)
        ReDim lineHeight(0 To n - 1)
        ReDim lineOffsetX(0 To n - 1)
        ReDim lineOffsetY(0 To n - 1)
        For i = 0 To n - 1
            lineText(i) = lines(i).text
            lineWidth(i) = lines(i).dimensions.x
            lineHeight(i) = lines(i).dimensions.y
            lineOffsetX(i) = lines(i).offset.x
            lineOffsetY(i) = lines(i).offset.y
        Next
    End If
    iV_FormatText_arrays = n
End Function

' Needs modification
Public Sub iV_DrawTextRotated(ByVal string_ As String, ByVal XPos As Single, ByVal YPos As Single, ByVal rotation_ As Single, ByVal fontID As iV_fonts)
    Dim color As Long, drawResult As DrawTextResult, position As Vector2f, offset As Vector2f, size As Vector2f
    If Not (StrPtr(string_) <> 0) Then ASSERT_FAILURE "string", "lib/ivis_opengl/textdraw.cpp:1756", "iV_DrawTextRotated", "Couldn't render string!": Exit Sub
    InitTextdrawGlobals

    If rotation_ <> 0! Then
        rotation_ = 180! - rotation_
    End If

    color = PIELIGHT_fromRGBA(Fix(font_colour(0) * 255!), Fix(font_colour(1) * 255!), Fix(font_colour(2) * 255!), Fix(font_colour(3) * 255!))

    drawResult = TextShaper_drawText(string_, fontID)

    If Not drawResult.text.bitmap Is Nothing Then
        If drawResult.text.bitmap.width() > 0 And drawResult.text.bitmap.height() > 0 Then
            Set textureID = Nothing
            Set textureID = gfx_api_context_createTextureForCompatibleImageUploads(1, drawResult.text.bitmap, "text::" & string_)
            textureID.upload 0, drawResult.text.bitmap
            position.x = XPos
            position.y = YPos
            offset.x = drawResult.text.offset_x / horizScaleFactor_
            offset.y = drawResult.text.offset_y / vertScaleFactor_
            size.x = drawResult.text.bitmap.width() / horizScaleFactor_
            size.y = drawResult.text.bitmap.height() / vertScaleFactor_
            iV_DrawImageText textureID, position, offset, size, rotation_, color
        End If
    End If
End Sub

Public Sub iV_DrawText(ByVal string_ As String, ByVal x As Single, ByVal y As Single, ByVal fontID As iV_fonts)
    iV_DrawTextRotated string_, x, y, 0!, fontID
End Sub

' (VB) a PIELIGHT from its bytes (PIELIGHT::fromRGBA)
Public Function PIELIGHT_fromRGBA(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
    PIELIGHT_fromRGBA = (r And &HFF&) Or ((g And &HFF&) * &H100&) Or ((b And &HFF&) * &H10000) Or SHL(a And &HFF&, 24)
End Function

' (VB) WzText(const WzString &string, iV_fonts fontID)
Public Function WzText_new(ByVal string_ As String, ByVal fontID As iV_fonts) As WzText
    Set WzText_new = New WzText
    WzText_new.setText string_, fontID
End Function

' (VB) a WzString's length (its code points)
Public Function WzString_length(ByVal s As String) As Long
    Dim t As CodePoints
    toUtf32 s, t
    WzString_length = t.size
End Function

' (VB) WzString::pop_back(): drops the last code point; false if there is none
Public Function WzString_pop_back(s As String) As Boolean
    Dim c As Long
    If Len(s) = 0 Then Exit Function
    c = AscW(Right$(s, 1)) And &HFFFF&
    If c >= &HDC00& And c <= &HDFFF& And Len(s) >= 2 Then
        c = AscW(Mid$(s, Len(s) - 1, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& Then
            s = Left$(s, Len(s) - 2)
            WzString_pop_back = True
            Exit Function
        End If
    End If
    s = Left$(s, Len(s) - 1)
    WzString_pop_back = True
End Function
