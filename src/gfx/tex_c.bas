Attribute VB_Name = "tex_c"
' lib/ivis_opengl/tex.cpp, lib/ivis_opengl/tex.h: the texture pages. A page is
' its index in TEX_PAGE_ (C++'s _TEX_PAGE, a std::vector<iTexPage>);
' _NAME_TO_TEX_PAGE_MAP is a Scripting.Dictionary (filename -> page). A page's
' texture is freed when its last reference goes (iTexPage's destructor, C++'s
' delete). optional<size_t> iV_GetTexture returns has_value and sets page.
' pie_AddTexPage's overload with a page is pie_AddTexPage_page.
' std::unordered_set<size_t> texPages (debugReloadTexturesFromDisk) is an array
' and its count, walked in array order.
Option Explicit

Private Type iTexPage
    filename As String
    id As gfx_api_texture               ' = nullptr
    textureType As texture_type         ' = gfx_api::texture_type::user_interface
End Type

Private TEX_PAGE_() As iTexPage         ' std::vector<iTexPage> _TEX_PAGE;
Private TEX_PAGE_size As Long
Private NAME_TO_TEX_PAGE_MAP_ As Object ' std::unordered_map<std::string, size_t> _NAME_TO_TEX_PAGE_MAP;

Public Const iV_TEXNAME_MAX As Long = 64
Public Const iV_TEXNAME_TCSUFFIX As String = "_tcmask"

'*************************************************************************

' (VB) the map, created on first use (a C++ global's constructor)
Private Sub NAME_TO_TEX_PAGE_MAP_init()
    If NAME_TO_TEX_PAGE_MAP_ Is Nothing Then
        Set NAME_TO_TEX_PAGE_MAP_ = CreateObject("Scripting.Dictionary")
        NAME_TO_TEX_PAGE_MAP_.CompareMode = 0  ' BinaryCompare
    End If
End Sub

' (VB) _TEX_PAGE.push_back(std::move(tex))
Private Sub TEX_PAGE_push_back(tex As iTexPage)
    If TEX_PAGE_size = 0 Then ReDim TEX_PAGE_(0 To 63)
    If TEX_PAGE_size > UBound(TEX_PAGE_) Then ReDim Preserve TEX_PAGE_(0 To 2 * TEX_PAGE_size - 1)
    TEX_PAGE_(TEX_PAGE_size).filename = tex.filename
    Set TEX_PAGE_(TEX_PAGE_size).id = tex.id
    Set tex.id = Nothing
    TEX_PAGE_(TEX_PAGE_size).textureType = tex.textureType
    TEX_PAGE_size = TEX_PAGE_size + 1
End Sub

'*************************************************************************

Public Function pie_Texture(ByVal page As Long) As gfx_api_texture
    Set pie_Texture = TEX_PAGE_(page).id
End Function

Public Function pie_NumberOfPages() As Long
    pie_NumberOfPages = TEX_PAGE_size
End Function

' Add a new texture page to the list
Public Function pie_ReserveTexture(ByVal filename As String, ByVal width As Long, ByVal height As Long) As Long
    Dim tex As iTexPage, page As Long
    NAME_TO_TEX_PAGE_MAP_init
    tex.filename = filename
    TEX_PAGE_push_back tex
    page = TEX_PAGE_size - 1
    NAME_TO_TEX_PAGE_MAP_(filename) = page
    pie_ReserveTexture = page
End Function

Public Sub pie_AssignTexture(ByVal page As Long, ByVal texture As gfx_api_texture)
    If page >= TEX_PAGE_size Then
        debug_ 111, LOG_ERROR, "pie_AssignTexture", "Invalid page (%zu), number of texpages=%zu", page, TEX_PAGE_size
        Set texture = Nothing           ' delete texture;
        Exit Sub
    End If
    Set TEX_PAGE_(page).id = texture    ' (the old texture: delete _TEX_PAGE[page].id)
End Sub

Private Function pie_AddTexPage_Impl(ByVal pTexture As gfx_api_texture, ByVal filename As String, ByVal textureType As texture_type, ByVal page As Long) As Long
    NAME_TO_TEX_PAGE_MAP_init
    If Not (Not pTexture Is Nothing And StrPtr(filename) <> 0) Then ASSERT_FAILURE "pTexture && filename", "lib/ivis_opengl/tex.cpp:122", "pie_AddTexPage_Impl", "Bad input parameter"

    NAME_TO_TEX_PAGE_MAP_(filename) = page
    debug_ 125, LOG_TEXTURE, "pie_AddTexPage_Impl", "%s page=%zu", filename, page

    Set TEX_PAGE_(page).id = pTexture   ' (the old texture: delete _TEX_PAGE[page].id)
    TEX_PAGE_(page).textureType = textureType

    ' Send back the texpage number so we can store it in the IMD
    pie_AddTexPage_Impl = page
End Function

Public Function pie_AddTexPage(ByVal pTexture As gfx_api_texture, ByVal filename As String, ByVal textureType As texture_type) As Long
    Dim tex As iTexPage, page As Long
    NAME_TO_TEX_PAGE_MAP_init
    If Not (Not pTexture Is Nothing And StrPtr(filename) <> 0) Then ASSERT_FAILURE "pTexture && filename", "lib/ivis_opengl/tex.cpp:138", "pie_AddTexPage", "Bad input parameter": pie_AddTexPage = 0: Exit Function
    If Not (Not NAME_TO_TEX_PAGE_MAP_.Exists(filename)) Then ASSERT_FAILURE "_NAME_TO_TEX_PAGE_MAP.count(filename) == 0", "lib/ivis_opengl/tex.cpp:139", "pie_AddTexPage", "tex page %s already exists", filename
    page = TEX_PAGE_size
    tex.filename = filename
    TEX_PAGE_push_back tex

    pie_AddTexPage = pie_AddTexPage_Impl(pTexture, filename, textureType, page)
End Function

' size_t pie_AddTexPage(gfx_api::texture *pTexture, const char *filename, gfx_api::texture_type textureType, size_t page)
Public Function pie_AddTexPage_page(ByVal pTexture As gfx_api_texture, ByVal filename As String, ByVal textureType As texture_type, ByVal page As Long) As Long
    NAME_TO_TEX_PAGE_MAP_init
    If Not (Not pTexture Is Nothing And StrPtr(filename) <> 0) Then ASSERT_FAILURE "pTexture && filename", "lib/ivis_opengl/tex.cpp:150", "pie_AddTexPage", "Bad input parameter": pie_AddTexPage_page = 0: Exit Function

    ' replace
    If NAME_TO_TEX_PAGE_MAP_.Exists(TEX_PAGE_(page).filename) Then NAME_TO_TEX_PAGE_MAP_.Remove TEX_PAGE_(page).filename
    TEX_PAGE_(page).filename = filename

    pie_AddTexPage_page = pie_AddTexPage_Impl(pTexture, filename, textureType, page)
End Function

'!
' * Turns page filename into a pagename + tc mask if possible
' * \param[in,out] filename Filename to pagify
Public Function pie_MakeTexPageTCMaskName(ByVal filename As String) As String
    Dim result As String, i As Long, c As Long
    result = filename
    If Left$(filename, 5) = "page-" Then
        ' filename starts with "page-"
        i = 5
        Do While i < Len(filename) - 1
            c = AscW(Mid$(filename, i + 1, 1))
            If Not (c >= 48 And c <= 57) Then Exit Do   ' isdigit
            i = i + 1
        Loop
        result = Left$(filename, i)
        result = result & iV_TEXNAME_TCSUFFIX
    End If
    pie_MakeTexPageTCMaskName = result
End Function

Public Function loadTextureHandleGraphicsOverrides(ByVal filename As String, ByVal textureType As texture_type, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As gfx_api_texture
    Dim loadPath As String, pTexture As gfx_api_texture
    ' First, try to load from the current graphics_overrides (if enabled)
    loadPath = WZ_CURRENT_GRAPHICS_OVERRIDES_PREFIX & "/texpages/"
    loadPath = loadPath & filename
    Set pTexture = gfx_api_context_loadTextureFromFile(loadPath, textureType, maxWidth, maxHeight, True)
    If pTexture Is Nothing Then
        ' Try to load it from the regular path
        loadPath = "texpages/"
        loadPath = loadPath & filename
        Set pTexture = gfx_api_context_loadTextureFromFile(loadPath, textureType, maxWidth, maxHeight)
    End If
    Set loadTextureHandleGraphicsOverrides = pTexture
End Function

' Retrieve the texture number for a given texture resource.
'
'  @note We keep textures in a separate data structure _TEX_PAGE apart from the
'        normal resource system.
'
'  @param filename The filename of the texture page to search for.
'  @param compression If we need to load it, should we use texture compression?
'  @param maxWidth If we need to load it, should we limit the texture width? (Resizes and preserves the texture image's aspect ratio)
'  @param maxHeight If we need to load it, should we limit the texture height? (Resizes and preserves the texture image's aspect ratio)
'
'  @return a non-negative index number for the texture, negative if no texture
'          with the given filename could be found
' (VB) optional<size_t>: returns has_value, the page in page
Public Function iV_GetTexture(ByVal filename As String, ByVal textureType As texture_type, page As Long, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1) As Boolean
    Dim pTexture As gfx_api_texture
    NAME_TO_TEX_PAGE_MAP_init
    If Not (StrPtr(filename) <> 0) Then ASSERT_FAILURE "filename != nullptr", "lib/ivis_opengl/tex.cpp:208", "iV_GetTexture", "filename must not be null"

    ' Have we already loaded this one then?
    If NAME_TO_TEX_PAGE_MAP_.Exists(filename) Then
        page = NAME_TO_TEX_PAGE_MAP_(filename)
        iV_GetTexture = True
        Exit Function
    End If

    Set pTexture = loadTextureHandleGraphicsOverrides(filename, textureType, maxWidth, maxHeight)
    If pTexture Is Nothing Then
        debug_ 220, LOG_ERROR, "iV_GetTexture", "Failed to load %s", filename
        iV_GetTexture = False
        Exit Function
    End If

    page = pie_AddTexPage(pTexture, filename, textureType)
    resDoResLoadCallback ' ensure loading screen doesn't freeze when loading large images
    iV_GetTexture = True
End Function

Public Function replaceTexture(ByVal oldfile As String, ByVal newfile As String) As Boolean
    Dim tmpname As String, page As Long, existingTextureType As texture_type, pTexture As gfx_api_texture
    NAME_TO_TEX_PAGE_MAP_init
    ' Load new one to replace it
    tmpname = oldfile
    ' Have we already loaded this one?
    If NAME_TO_TEX_PAGE_MAP_.Exists(tmpname) Then
        gfx_api_context_debugStringMarker "Replacing texture"
        page = NAME_TO_TEX_PAGE_MAP_(tmpname)
        existingTextureType = TEX_PAGE_(page).textureType
        If debugPartEnabled(LOG_TEXTURE) Then debug_ 240, LOG_TEXTURE, "replaceTexture", "Replacing texture %s with %s from index %zu (tex id %u)", tmpname, newfile, page, TEX_PAGE_(page).id.id()
        Set pTexture = loadTextureHandleGraphicsOverrides(newfile, existingTextureType)
        If Not pTexture Is Nothing Then
            tmpname = newfile
            pie_AddTexPage_page pTexture, tmpname, existingTextureType, page
            replaceTexture = True
            Exit Function
        End If
    End If
    debug_ 249, LOG_TEXTURE, "replaceTexture", "Nothing to replace - old (not found): %s, new (not used): %s", oldfile, newfile
    replaceTexture = False
End Function

' texPages(0 .. count - 1)
Public Function debugReloadTexturesFromDisk(texPages() As Long, ByVal count As Long) As Boolean
    Dim filename As String, i As Long, page As Long, existingTextureType As texture_type, pTexture As gfx_api_texture
    For i = 0 To count - 1
        page = texPages(i)
        existingTextureType = TEX_PAGE_(page).textureType
        filename = TEX_PAGE_(page).filename
        debug_ 260, LOG_TEXTURE, "debugReloadTexturesFromDisk", "Reloading texture %s from index %zu", filename, page
        Set pTexture = loadTextureHandleGraphicsOverrides(filename, existingTextureType)
        If pTexture Is Nothing Then
            debug_ 264, LOG_INFO, "debugReloadTexturesFromDisk", "Failed to reload texture: %s", filename
            GoTo continue_
        End If
        pie_AddTexPage_page pTexture, filename, existingTextureType, page
continue_:
    Next
    debugReloadTexturesFromDisk = True
End Function

Public Sub pie_TexShutDown()
    NAME_TO_TEX_PAGE_MAP_init
    ' TODO, lazy deletions for faster loading of next level
    debug_ 275, LOG_TEXTURE, "pie_TexShutDown", "Cleaning out %u textures", TEX_PAGE_size
    Erase TEX_PAGE_
    TEX_PAGE_size = 0
    NAME_TO_TEX_PAGE_MAP_.RemoveAll
End Sub

Public Sub pie_TexInit()
    debug_ 282, LOG_TEXTURE, "pie_TexInit", "pie_TexInit successful"
End Sub

Public Sub iV_unloadImage(ByVal image As iV_Image)
    If Not image Is Nothing Then
        image.clear
    Else
        debug_ 293, LOG_ERROR, "iV_unloadImage", "Tried to free invalid image!"
    End If
End Sub

Public Function iV_getPixelFormat(ByVal image As iV_Image) As pixel_format
    Select Case image.channels()
    Case 1
        iV_getPixelFormat = pixel_format_FORMAT_R8_UNORM
    Case 2
        iV_getPixelFormat = pixel_format_FORMAT_RG8_UNORM
    Case 3
        iV_getPixelFormat = pixel_format_FORMAT_RGB8_UNORM_PACK8
    Case 4
        iV_getPixelFormat = pixel_format_FORMAT_RGBA8_UNORM_PACK8
    Case Else
        debug_ 310, LOG_FATAL, "iV_getPixelFormat", "iV_getPixelFormat: Unsupported image channels: %u", image.channels()
        iV_getPixelFormat = pixel_format_invalid
    End Select
End Function
