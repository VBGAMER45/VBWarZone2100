Attribute VB_Name = "texture_c"
' src/texture.cpp, texture.h: the tile textures (the decal texture arrays) and
' the tileset's radar colours. A texture array is WzGfx's gfx_api_texture_array
' (Nothing = nullptr); loadTextureArrayFromFiles' default-texture lambdas are
' GfxApi's GenerateDefaultTextureFunc numbers. WZ_PHYSFS_enumerateFiles' lambda
' is a loop over the names it returns.
Option Explicit

Private Const MIPMAP_MAX_ As Long = 512 ' MIPMAP_MAX (mipmap_max is used more)

Public Const TILES_IN_PAGE_COLUMN As Long = 16
Public Const TILES_IN_PAGE_ROW As Long = 16
Public Const TILES_IN_PAGE As Long = TILES_IN_PAGE_COLUMN * TILES_IN_PAGE_ROW
Public Const MAX_TILES As Long = TILES_IN_PAGE

Public Type TILE_TEX_INFO
    uOffset As Single                   ' Offset into texture page to left hand edge
    vOffset As Single                   ' Offset into texture page to top hand edge
    texPage As Long                     ' unsigned int: Which textpage is the tile in? TileNumber/16 basically;
End Type

' Texture page and coordinates for each tile
Public tileTexInfo(0 To MAX_TILES - 1) As TILE_TEX_INFO

Private firstPage As Long               ' size_t: the last used page before we start adding terrain textures
Public decalTexArr As gfx_api_texture_array
Public decalNormalArr As gfx_api_texture_array
Public decalSpecularArr As gfx_api_texture_array
Public decalHeightArr As gfx_api_texture_array
Private mipmap_max As Long, mipmap_levels As Long
Private maxTextureSize As Long          ' = 2048: the maximum size texture we will create

' (VB) the C++ initialisers (Sub Main calls it)
Public Sub texture_c_init()
    maxTextureSize = 2048
End Sub

Public Sub setTextureSize(ByVal texSize As Long)
    If texSize < 16 Or texSize Mod 16 <> 0 Then
        debug_ 63, LOG_ERROR, "setTextureSize", "Attempted to set bad texture size %d! Ignored.", texSize
        Exit Sub
    End If
    maxTextureSize = texSize
    debug_ 67, LOG_TEXTURE, "setTextureSize", "texture size set to %d", texSize
End Sub

Public Function getTextureSize() As Long
    getTextureSize = maxTextureSize
End Function

Public Function getCurrentTileTextureSize() As Long
    getCurrentTileTextureSize = mipmap_max
End Function

Public Function getMaxTileTextureSize(ByVal dir_ As String) As Long
    Dim res As Long
    res = MIPMAP_MAX_
    Do While res > 0
        If PHYSFS_exists(dir_ & "-" & CStr(res) & "/tile-00.ktx2") <> 0 Then Exit Do
        If PHYSFS_exists(dir_ & "-" & CStr(res) & "/tile-00.png") <> 0 Then Exit Do
        res = res \ 2
    Loop
    getMaxTileTextureSize = res
End Function

' (VB) (int)log2(n) of a positive power of two: its exponent (log2 is exact there)
Private Function intLog2(ByVal n As Long) As Long
    Dim r As Long
    Do While n > 1
        n = n \ 2
        r = r + 1
    Loop
    intLog2 = r
End Function

' (VB) a file name's suffix (the enumeration lambda's hasSuffix)
Private Function hasSuffix(ByVal fileName As String, ByVal suf As String) As Boolean
    If Len(fileName) < Len(suf) Then Exit Function
    hasSuffix = (Right$(fileName, Len(suf)) = suf)
End Function

Public Function texLoad(ByVal fileName As String) As Boolean
    Dim fullPath As String, partialPath As String, buffer() As Byte, chars As String
    Dim i As Long, j As Long, k As Long, size As Long, n As Long
    Dim maxTileTexSize As Long, max_texture_size As Long
    Dim r As Long, g As Long, b As Long, cnt As Long
    Dim xSize As Long, ySize As Long, xLimit As Long, yLimit As Long
    Dim has_nm As Boolean, has_sm As Boolean, has_hm As Boolean, has_auxillary_texture_info As Boolean
    Dim maxTileNo As Long, files() As String, filesCount As Long, tile As Long
    Dim fullPath_nm As String, fullPath_sm As String, fullPath_hm As String, fullPath_base As String
    Dim tile_base_filepaths() As String, tile_base_filepaths_size As Long
    Dim tile_nm_filepaths() As String, tile_nm_filepaths_size As Long
    Dim tile_sm_filepaths() As String, tile_sm_filepaths_size As Long
    Dim tile_hm_filepaths() As String, tile_hm_filepaths_size As Long
    Dim usedFilenames_tmp() As String, usedFilenames_tmp_size As Long

    firstPage = pie_NumberOfPages()

    ' store the filename so we can later determine which tileset we are using
    tilesetDir = fileName

    ' reset defaults
    maxTileTexSize = getMaxTileTextureSize(fileName)
    mipmap_max = maxTileTexSize
    mipmap_levels = intLog2(mipmap_max \ 8) ' 4 for 128

    max_texture_size = gfx_api_context_get_context_value(context_value_MAX_TEXTURE_SIZE)

    Do While max_texture_size < mipmap_max * TILES_IN_PAGE_COLUMN
        mipmap_max = mipmap_max \ 2
        mipmap_levels = mipmap_levels - 1
        debug_ 114, LOG_ERROR, "texLoad", "Max supported texture size %dx%d is too low - reducing texture detail to %dx%d.", max_texture_size, max_texture_size, mipmap_max, mipmap_max
        If Not (mipmap_levels > 0) Then ASSERT_FAILURE "mipmap_levels > 0", "src/texture.cpp:116", "texLoad", "Supported texture size %d is too low to load any mipmap levels!", max_texture_size
        If mipmap_levels = 0 Then
            debug_ 120, LOG_FATAL, "texLoad", "Supported texture size %d is too low to load any mipmap levels!: %s", max_texture_size, fileName
            exit_ 1
        End If
    Loop
    Do While maxTextureSize < mipmap_max
        mipmap_max = mipmap_max \ 2
        mipmap_levels = mipmap_levels - 1
        debug_ 128, LOG_TEXTURE, "texLoad", "Downgrading texture quality to %d due to user setting %d", mipmap_max, maxTextureSize
    Loop

    ' Get and set radar colours

    fullPath = fileName & ".radar"
    If Not loadFile(fullPath, buffer, size) Then
        debug_ 136, LOG_FATAL, "texLoad", "texLoad: Could not find radar colours at %s", fullPath
        exit_ 3                         ' abort(): cannot recover; we could possibly generate a random palette?
    End If
    ' (VB) the bytes as characters, one each, so that offsets are the C++ pointer's
    chars = String$(size, vbNullChar)
    For n = 0 To size - 1
        Mid$(chars, n + 1, 1) = ChrW$(buffer(n))
    Next
    i = 0 ' tile
    j = 0 ' place in buffer
    Do
        cnt = 0
        k = sscanf(Mid$(chars, j + 1), "%2x%2x%2x%n", r, g, b, cnt)
        j = j + cnt
        If k >= 3 Then
            radarColour i, r, g, b
        End If
        i = i + 1 ' next tile
    Loop While k >= 3 And j + 6 < size
    Erase buffer

    ' Now load the actual tiles
    xSize = 1
    ySize = 1
    xLimit = TILES_IN_PAGE_COLUMN * mipmap_max
    yLimit = TILES_IN_PAGE_ROW * mipmap_max

    ' pad width and height into ^2 values
    Do
        xSize = xSize * 2
    Loop While xLimit > xSize
    Do
        ySize = ySize * 2
    Loop While yLimit > ySize

    partialPath = fileName & "-" & CStr(maxTileTexSize)

    has_nm = False: has_sm = False: has_hm = False
    has_auxillary_texture_info = False
    maxTileNo = 0

    If WZ_PHYSFS_enumerateFiles(partialPath, files, filesCount) Then
        For n = 0 To filesCount - 1
            has_nm = has_nm Or hasSuffix(files(n), "_nm.png") Or hasSuffix(files(n), "_nm.ktx2")
            has_sm = has_sm Or hasSuffix(files(n), "_sm.png") Or hasSuffix(files(n), "_sm.ktx2")
            has_hm = has_hm Or hasSuffix(files(n), "_hm.png") Or hasSuffix(files(n), "_hm.ktx2")
            tile = 0
            If sscanf(files(n), "tile-%02d.png", tile) = 1 Then
                If tile > maxTileNo Then maxTileNo = tile
            End If
        Next
    End If
    has_auxillary_texture_info = has_nm Or has_sm Or has_hm
    debug_ 190, LOG_TEXTURE, "texLoad", "texLoad: Found %d tiles", maxTileNo
    Set decalTexArr = Nothing
    Set decalNormalArr = Nothing
    Set decalSpecularArr = Nothing
    Set decalHeightArr = Nothing

    ReDim tile_base_filepaths(0 To maxTileNo)
    ReDim tile_nm_filepaths(0 To maxTileNo)
    ReDim tile_sm_filepaths(0 To maxTileNo)
    ReDim tile_hm_filepaths(0 To maxTileNo)
    ReDim usedFilenames_tmp(0 To 4 * (maxTileNo + 1))

    For k = 0 To maxTileNo
        fullPath_base = gfx_api_imageLoadFilenameFromInputFilename(va("%s/tile-%02d.png", partialPath, k))
        tile_base_filepaths(tile_base_filepaths_size) = fullPath_base
        tile_base_filepaths_size = tile_base_filepaths_size + 1
        usedFilenames_tmp(usedFilenames_tmp_size) = fullPath_base
        usedFilenames_tmp_size = usedFilenames_tmp_size + 1

        If has_auxillary_texture_info Then
            fullPath_nm = gfx_api_imageLoadFilenameFromInputFilename(va("%s/tile-%02d_nm.png", partialPath, k))
            fullPath_sm = gfx_api_imageLoadFilenameFromInputFilename(va("%s/tile-%02d_sm.png", partialPath, k))
            fullPath_hm = gfx_api_imageLoadFilenameFromInputFilename(va("%s/tile-%02d_hm.png", partialPath, k))

            If has_nm Then
                usedFilenames_tmp(usedFilenames_tmp_size) = fullPath_nm
                usedFilenames_tmp_size = usedFilenames_tmp_size + 1
            End If
            If has_sm Then
                usedFilenames_tmp(usedFilenames_tmp_size) = fullPath_sm
                usedFilenames_tmp_size = usedFilenames_tmp_size + 1
            End If
            If has_hm Then
                usedFilenames_tmp(usedFilenames_tmp_size) = fullPath_hm
                usedFilenames_tmp_size = usedFilenames_tmp_size + 1
            End If

            If Not gfx_api_checkImageFilesWouldLoadFromSameParentMountPath(usedFilenames_tmp, usedFilenames_tmp_size, True) Then
                debug_ 233, LOG_INFO, "texLoad", "One of these images would load from a different path / package: %s, %s, %s, %s - defaulting to using base texture only", fullPath, fullPath_nm, fullPath_sm, fullPath_hm
                fullPath_nm = ""
                fullPath_sm = ""
                fullPath_hm = ""
            End If
            If has_nm Then
                If PHYSFS_exists(fullPath_nm) = 0 Then
                    fullPath_nm = ""
                End If
                tile_nm_filepaths(tile_nm_filepaths_size) = fullPath_nm
                tile_nm_filepaths_size = tile_nm_filepaths_size + 1
            End If
            If has_sm Then
                If PHYSFS_exists(fullPath_sm) = 0 Then
                    fullPath_sm = ""
                End If
                tile_sm_filepaths(tile_sm_filepaths_size) = fullPath_sm
                tile_sm_filepaths_size = tile_sm_filepaths_size + 1
            End If
            If has_hm Then
                If PHYSFS_exists(fullPath_hm) = 0 Then
                    fullPath_hm = ""
                End If
                tile_hm_filepaths(tile_hm_filepaths_size) = fullPath_hm
                tile_hm_filepaths_size = tile_hm_filepaths_size + 1
            End If
        End If
    Next

    Set decalTexArr = gfx_api_context_loadTextureArrayFromFiles(tile_base_filepaths, tile_base_filepaths_size, texture_type_game_texture, mipmap_max, mipmap_max, GenerateDefaultTexture_zeroed, ProgressCallback_none, "decalTexArr")
    If has_auxillary_texture_info Then
        If has_nm Then
            Set decalNormalArr = gfx_api_context_loadTextureArrayFromFiles(tile_nm_filepaths, tile_nm_filepaths_size, texture_type_normal_map, mipmap_max, mipmap_max, GenerateDefaultTexture_normalMap, ProgressCallback_none, "decalNormalArr")
            If Not (Not decalNormalArr Is Nothing) Then ASSERT_FAILURE "decalNormalArr != nullptr", "src/texture.cpp:290", "texLoad", "Failed to load tile normals"
        End If
        If has_sm Then
            Set decalSpecularArr = gfx_api_context_loadTextureArrayFromFiles(tile_sm_filepaths, tile_sm_filepaths_size, texture_type_specular_map, mipmap_max, mipmap_max, GenerateDefaultTexture_zeroed, ProgressCallback_none, "decalSpecularArr")
            If Not (Not decalSpecularArr Is Nothing) Then ASSERT_FAILURE "decalSpecularArr != nullptr", "src/texture.cpp:300", "texLoad", "Failed to load tile specular maps"
        End If
        If has_hm Then
            Set decalHeightArr = gfx_api_context_loadTextureArrayFromFiles(tile_hm_filepaths, tile_hm_filepaths_size, texture_type_height_map, mipmap_max, mipmap_max, GenerateDefaultTexture_zeroed, ProgressCallback_none, "decalHeightArr")
            If Not (Not decalHeightArr Is Nothing) Then ASSERT_FAILURE "decalHeightArr != nullptr", "src/texture.cpp:310", "texLoad", "Failed to load terrain height maps"
        End If
    End If
    ' flush all of the array textures
    If Not decalTexArr Is Nothing Then
        decalTexArr.flush
    Else
        debug_ 320, LOG_FATAL, "texLoad", "Failed to load one or more terrain decals"
    End If
    If Not decalNormalArr Is Nothing Then decalNormalArr.flush
    If Not decalSpecularArr Is Nothing Then decalSpecularArr.flush
    If Not decalHeightArr Is Nothing Then decalHeightArr.flush
    texLoad = True
End Function

Public Function reloadTileTextures() As Boolean
    Dim dir_ As String
    If StrPtr(tilesetDir) = 0 Then
        reloadTileTextures = False
        Exit Function
    End If
    dir_ = tilesetDir
    reloadTileTextures = texLoad(dir_)
End Function
