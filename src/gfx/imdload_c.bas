Attribute VB_Name = "imdload_c"
' lib/ivis_opengl/imdload.cpp, imd.h: the PIE model loader and the model
' store. The models are pools (ivisdef_h): imdShapes (iIMDShape, the display
' levels, chained through nxt) and imdBaseShapes (iIMDBaseShape, the game
' state part); a handle is the 1-based index, 0 = nullptr, and the free
' slots chain through nxt and m_displayModel. ModelMap (std::unordered_map
' by name) is a Scripting.Dictionary of base shape handles, iterated in
' insertion order (C++: hash order). The file is a Byte array plus the same
' bytes as characters (IMD_FILE), a const char * an offset in it (-1 =
' nullptr), passed down so that a nested load (an EVENT's modelGet) has its
' own. sscanf is sscanf_vb's; the float arithmetic goes through glm_vb
' (rounded to float at each step, as C++ does). The vertex data goes to the
' shape's GPU buffers as C++ uploads it, and stays in the shape too (vbo_*)
' for modelDump; with no gfx context (the --vbtest runs) no buffer is made.
' The display shapes' fields piedraw and the exe's display code read have
' accessors (iIMDShape_<field>); iIMDShape::getTextures is
' iIMDShape_getTextures(shape).
' (VB) modelDump writes a model as text, for tools\run_tests.py.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' The maximum polygons for a model is influenced by several factors:
' 1. The vertexCount *must* be less than UINT16_MAX (65535), because this special vertex index value is treated as restarting the assembly of primitives in some cases (example: Vulkan on MoltenVK) - NOTE: we currently use index_type::u16 for our model VBO_INDEX buffers
' 2. Given how the game world and models are displayed, using a very large number of polygons is effectively a waste
'#define MAX_PIE_POLYGONS 21844 // floor((UINT16_MAX-1) / 3)
Private Const MAX_PIE_POLYGONS As Long = 8192

' Scale animation numbers from int to float
Private Const INT_SCALE As Long = 1000

' imd.h
Public Const PIE_NAME As String = "PIE"  ' Pumpkin image export data file
Public Const PIE_VER As Long = 2
Public Const PIE_FLOAT_VER As Long = 3
Public Const PIE_VER_4 As Long = 4

Public Const PIE_MIN_VER As Long = PIE_VER
Public Const PIE_MAX_VER As Long = 4

' PIE model flags
' premultiplied implies additive
Public Const iV_IMD_NO_ADDITIVE As Long = &H1
Public Const iV_IMD_ADDITIVE As Long = &H2
Public Const iV_IMD_PREMULTIPLIED As Long = &H4
' pitch to camera implies roll to camera
Public Const iV_IMD_ROLL_TO_CAMERA As Long = &H10
Public Const iV_IMD_PITCH_TO_CAMERA As Long = &H20
Public Const iV_IMD_NOSTRETCH As Long = &H1000
Public Const iV_IMD_TCMASK As Long = &H10000

' polygon flags	b0..b7: col, b24..b31: anim index

Public Const iV_IMD_TEX As Long = &H200          ' this is both a polygon and pie flag
Public Const iV_IMD_TEXANIM As Long = &H4000     ' iV_IMD_TEX must be set also

Private Type IMD_FILE                   ' (VB) the file data, and the end (FileDataEnd)
    data() As Byte
    chars As String                     ' the bytes, one character each
    size As Long
End Type

Private Type IMD_Line
    lineContents As String
    pNextLineBegin As Long              ' a pointer to the data after this line (including past any newline characters); -1: nullptr
End Type

' the texture commands of LevelSettings::tilesetTextureFiles (std::unordered_map<std::string, std::string> per tileset)
Private Const TEX_TEXTURE As Long = 0
Private Const TEX_TCMASK As Long = 1
Private Const TEX_NORMALMAP As Long = 2
Private Const TEX_SPECULARMAP As Long = 3

Private Type LevelSettings
    ' TYPE
    imd_flags As Long                   ' optional<uint32_t>
    has_imd_flags As Boolean

    ' INTERPOLATE
    interpolate As Boolean              ' optional<bool>
    has_interpolate As Boolean

    ' TEXTURES
    tilesetTextureFiles(0 To 2, 0 To 3) As String   ' "" when the command is not in the map
End Type

' the pools
Public imdShapes() As iIMDShape
Private imdShapesCap As Long
Private freeShapes As Long
Public imdBaseShapes() As iIMDBaseShape
Private imdBaseShapesCap As Long
Private freeBaseShapes As Long

Private models As Object                ' ModelMap
Private currentTilesetIdx As Long

Private modelLoadingErrors As Long
Private modelTextureLoadingFailures As Long

' performance hack
Private vertices() As Single
Private vertices_size As Long
Private normals() As Single
Private normals_size As Long
Private texcoords() As Single           ' texcoords + texAnim
Private texcoords_size As Long
Private tangents() As Single
Private tangents_size As Long
Private bitangents() As Single
Private bitangents_size As Long
Private indices() As Long               ' uint16_t: size is npolys * 3 * numFrames
Private indices_size As Long
Private vertexCount As Long             ' uint16_t

' (VB) modelDump's text, in pieces
Private dumpChunks() As String
Private dumpCount As Long

'--------------------------------------------------------------------------
' (VB) the pools

Private Sub InitModels()
    If models Is Nothing Then
        Set models = CreateObject("Scripting.Dictionary")
        models.CompareMode = vbBinaryCompare
    End If
End Sub

' std::make_unique<iIMDShape>(): a constructed shape
Private Function newShape() As Long
    Dim h As Long, i As Long, blank As iIMDShape
    If freeShapes = 0 Then
        ReDim Preserve imdShapes(1 To imdShapesCap * 2 + 64)
        For i = imdShapesCap + 1 To UBound(imdShapes)
            If i < UBound(imdShapes) Then imdShapes(i).nxt = i + 1 Else imdShapes(i).nxt = 0
        Next
        freeShapes = imdShapesCap + 1
        imdShapesCap = UBound(imdShapes)
    End If
    h = freeShapes
    freeShapes = imdShapes(h).nxt
    imdShapes(h) = blank
    imdShapes(h).interpolate = 1
    imdShapes(h).m_textures.texpage = iV_TEX_INVALID
    imdShapes(h).m_textures.tcmaskpage = iV_TEX_INVALID
    imdShapes(h).m_textures.normalpage = iV_TEX_INVALID
    imdShapes(h).m_textures.specularpage = iV_TEX_INVALID
    newShape = h
End Function

' a std::unique_ptr<iIMDShape> released: iIMDShape::~iIMDShape (and its next levels)
Private Sub deleteShape(ByVal h As Long)
    Dim nxt As Long, blank As iIMDShape
    Do While h <> 0
        nxt = imdShapes(h).nxt
        ' freeInternalResources: the shadow edge list and the buffers (their last references) go with the shape
        imdShapes(h) = blank
        imdShapes(h).nxt = freeShapes
        freeShapes = h
        h = nxt
    Loop
End Sub

' iIMDBaseShape::iIMDBaseShape(std::unique_ptr<iIMDShape> a, const WzString &path, const WzString &filename)
Private Function newBaseShape(ByVal a As Long, ByVal path As String, ByVal filename As String) As Long
    Dim h As Long, i As Long, blank As iIMDBaseShape
    If freeBaseShapes = 0 Then
        ReDim Preserve imdBaseShapes(1 To imdBaseShapesCap * 2 + 64)
        For i = imdBaseShapesCap + 1 To UBound(imdBaseShapes)
            If i < UBound(imdBaseShapes) Then imdBaseShapes(i).m_displayModel = i + 1 Else imdBaseShapes(i).m_displayModel = 0
        Next
        freeBaseShapes = imdBaseShapesCap + 1
        imdBaseShapesCap = UBound(imdBaseShapes)
    End If
    h = freeBaseShapes
    freeBaseShapes = imdBaseShapes(h).m_displayModel
    imdBaseShapes(h) = blank
    imdBaseShapes(h).min = imdShapes(a).min
    imdBaseShapes(h).max = imdShapes(a).max
    imdBaseShapes(h).sradius = imdShapes(a).sradius
    imdBaseShapes(h).radius = imdShapes(a).radius
    imdBaseShapes(h).ocen = imdShapes(a).ocen
    imdBaseShapes(h).connectors_size = imdShapes(a).connectors_size
    If imdShapes(a).connectors_size > 0 Then imdBaseShapes(h).connectors = imdShapes(a).connectors
    imdBaseShapes(h).path = path
    imdBaseShapes(h).filename = filename
    imdBaseShapes(h).m_displayModel = a
    newBaseShape = h
End Function

' iIMDBaseShape::~iIMDBaseShape (its display model with it)
Private Sub deleteBaseShape(ByVal h As Long)
    Dim blank As iIMDBaseShape
    ' currently nothing else to do
    deleteShape imdBaseShapes(h).m_displayModel
    imdBaseShapes(h) = blank
    imdBaseShapes(h).m_displayModel = freeBaseShapes
    freeBaseShapes = h
End Sub

Private Sub fixupDisplayModelForBaseModel(ByVal baseModel As Long, ByVal newDisplayModel As Long)
    Dim prior_override_model_connectors_count As Long, i As Long, filename As String
    ' first, some sanity checks
    If imdShapes(newDisplayModel).connectors_size < imdBaseShapes(baseModel).connectors_size Then
        ' override models should not have fewer connectors than the base model
        prior_override_model_connectors_count = imdShapes(newDisplayModel).connectors_size
        ' to avoid crashes later, copy over the extra connectors from the base model
        For i = imdShapes(newDisplayModel).connectors_size To imdBaseShapes(baseModel).connectors_size - 1
            ReDim Preserve imdShapes(newDisplayModel).connectors(0 To i)
            imdShapes(newDisplayModel).connectors(i) = imdBaseShapes(baseModel).connectors(i)
            imdShapes(newDisplayModel).connectors_size = i + 1
        Next
        ' also output a log entry, so someone can know to fix the graphics override model
        filename = imdBaseShapes(baseModel).filename
        debug_ 210, LOG_INFO, "fixupDisplayModelForBaseModel", "Graphics override model %s is missing connectors (override connectors: %zu, base model connectors: %zu)", filename, prior_override_model_connectors_count, imdBaseShapes(baseModel).connectors_size
    End If
End Sub

' iIMDBaseShape::replaceDisplayModel
Private Sub iIMDBaseShape_replaceDisplayModel(ByVal baseShape As Long, ByVal newDisplayModel As Long)
    fixupDisplayModelForBaseModel baseShape, newDisplayModel

    deleteShape imdBaseShapes(baseShape).m_displayModel
    imdBaseShapes(baseShape).m_displayModel = newDisplayModel
End Sub

' iIMDShape::operator=(iIMDShape &&other): this_ takes other's contents, other gets this_'s
Private Sub iIMDShape_moveAssign(ByVal this_ As Long, ByVal other As Long)
    Dim tmp As iIMDShape
    If this_ <> other Then
        ' Release existing resources (freeInternalResources: the shadow edge list and the buffers)
        Erase imdShapes(this_).shadowEdgeList
        Erase imdShapes(this_).buffers

        ' Transfer resources from other
        tmp = imdShapes(this_)
        imdShapes(this_) = imdShapes(other)
        imdShapes(other) = tmp
        ' (std::swap(objanimframes, objanimframes) swaps this one's with itself: both keep theirs)
        imdShapes(other).objanimframes = imdShapes(this_).objanimframes
        imdShapes(this_).objanimframes = tmp.objanimframes
        imdShapes(this_).shadowUsesAlt = (imdShapes(this_).altShadowPoints_size <> 0 And imdShapes(this_).altShadowPolys_size <> 0)

        ' Reset other's resources
        Erase imdShapes(other).shadowEdgeList
        Erase imdShapes(other).buffers
    End If
End Sub

' const iIMDShapeTextures& iIMDShape::getTextures() const (shape: an iIMDShape handle)
Public Function iIMDShape_getTextures(ByVal shape As Long) As iIMDShapeTextures
    If Not imdShapes(shape).m_textures.initialized Then
        ' Load the textures on-demand
        If Not imd_load_level_textures_(shape, currentTilesetIdx, imdShapes(shape).m_textures) Then
            modelTextureLoadingFailures = modelTextureLoadingFailures + 1
        End If
        imdShapes(shape).m_textures.initialized = True
    End If

    iIMDShape_getTextures = imdShapes(shape).m_textures
End Function

' PIELIGHT iIMDShape::getTeamColourForModel(int team) const
Public Function iIMDShape_getTeamColourForModel(ByVal shape As Long, ByVal team As Long) As Long
    iIMDShape_getTeamColourForModel = pal_GetTeamColour(team)
End Function

' void iIMDShape::reloadTexturesIfLoaded()
Private Sub iIMDShape_reloadTexturesIfLoaded(ByVal shape As Long)
    If Not imdShapes(shape).m_textures.initialized Then
        Exit Sub
    End If
    imdShapes(shape).m_textures.texpage = iV_TEX_INVALID
    imdShapes(shape).m_textures.tcmaskpage = iV_TEX_INVALID
    imdShapes(shape).m_textures.normalpage = iV_TEX_INVALID
    imdShapes(shape).m_textures.specularpage = iV_TEX_INVALID
    imd_load_level_textures_ shape, currentTilesetIdx, imdShapes(shape).m_textures
    imdShapes(shape).m_textures.initialized = True
End Sub

' iIMDBaseShape::debugReloadDisplayModel
Private Function iIMDBaseShape_debugReloadDisplayModel(ByVal baseShape As Long) As Boolean
    iIMDBaseShape_debugReloadDisplayModel = iIMDBaseShape_debugReloadDisplayModelInternal(baseShape, True)
End Function

' iIMDBaseShape::debugReloadDisplayModelInternal
Private Function iIMDBaseShape_debugReloadDisplayModelInternal(ByVal baseShape As Long, ByVal recurseAnimPie As Boolean) As Boolean
    Dim newDisplayModel As Long, i As Long, pie As Long
    If Not (imdBaseShapes(baseShape).m_displayModel <> 0) Then ASSERT_FAILURE "m_displayModel != nullptr", "lib/ivis_opengl/imdload.cpp:228", "debugReloadDisplayModelInternal", "No display model??": iIMDBaseShape_debugReloadDisplayModelInternal = False: Exit Function

    ' first, try to load a graphics override - at the same path but with a prefix
    newDisplayModel = tryLoadDisplayModelInternal(WZ_CURRENT_GRAPHICS_OVERRIDES_PREFIX & "/" & imdBaseShapes(baseShape).path, imdBaseShapes(baseShape).filename, False, True)

    If newDisplayModel = 0 Then
        ' load model from base path
        newDisplayModel = tryLoadDisplayModelInternal(imdBaseShapes(baseShape).path, imdBaseShapes(baseShape).filename, False, True)
    End If

    If newDisplayModel = 0 Then
        debug_ 241, LOG_ERROR, "debugReloadDisplayModelInternal", "Failed to reload display model?: %s%s", imdBaseShapes(baseShape).path, imdBaseShapes(baseShape).filename
        iIMDBaseShape_debugReloadDisplayModelInternal = False
        Exit Function
    End If

    fixupDisplayModelForBaseModel baseShape, newDisplayModel

    ' swap out the display model - use move assignment to avoid invalidating any cached pointers to the displaymodel iIMDShape* (as this function is likely called at runtime)
    iIMDShape_moveAssign imdBaseShapes(baseShape).m_displayModel, newDisplayModel
    deleteShape newDisplayModel             ' (the old contents, and their next levels, go with newDisplayModel)

    ' also reload the display models for any objanimpie in the newDisplayModel (as tryLoadDisplayModelInternal will not force-reload them if their names are equal to a previously-loaded model)
    If recurseAnimPie Then
        For i = 0 To ANIM_EVENT_COUNT - 1
            pie = imdShapes(imdBaseShapes(baseShape).m_displayModel).objanimpie(i)
            If pie <> 0 Then
                iIMDBaseShape_debugReloadDisplayModelInternal pie, False
            End If
        Next
    End If

    iIMDBaseShape_debugReloadDisplayModelInternal = True
End Function

Public Function getModelLoadingErrorCount() As Long
    getModelLoadingErrorCount = modelLoadingErrors
End Function

Public Function getModelTextureLoadingFailuresCount() As Long
    getModelTextureLoadingFailuresCount = modelTextureLoadingFailures
End Function

Public Sub modelShutdown()
    Dim k As Variant
    InitModels
    For Each k In models.keys
        deleteBaseShape models(k)
    Next
    models.RemoveAll
    modelLoadingErrors = 0
    modelTextureLoadingFailures = 0
End Sub

Public Function debugReloadDisplayModelsForBaseModel(ByVal baseModel As Long) As Boolean
    debugReloadDisplayModelsForBaseModel = iIMDBaseShape_debugReloadDisplayModel(baseModel)
End Function

Public Sub debugReloadAllDisplayModels()
    Dim k As Variant
    InitModels
    For Each k In models.keys
        iIMDBaseShape_debugReloadDisplayModelInternal models(k), False
    Next
End Sub

Public Sub modelUpdateTilesetIdx(ByVal tilesetIdx As Long)
    Dim k As Variant, pDisplayShape As Long
    If tilesetIdx = currentTilesetIdx Then
        Exit Sub
    End If
    currentTilesetIdx = tilesetIdx
    ' reload all initialized model textures
    InitModels
    For Each k In models.keys
        pDisplayShape = imdBaseShapes(models(k)).m_displayModel
        Do While pDisplayShape <> 0
            iIMDShape_reloadTexturesIfLoaded pDisplayShape
            pDisplayShape = imdShapes(pDisplayShape).nxt
        Loop
    Next
End Sub

Public Sub modelReloadAllModelTextures()
    Dim k As Variant, pDisplayShape As Long, textures As iIMDShapeTextures, pages(0 To 3) As Long, i As Long, j As Long
    Dim texPagesToReloadFromDisk() As Long, n As Long, found As Boolean
    ReDim texPagesToReloadFromDisk(0 To 15)
    InitModels
    For Each k In models.keys
        pDisplayShape = imdBaseShapes(models(k)).m_displayModel
        Do While pDisplayShape <> 0
            textures = iIMDShape_getTextures(pDisplayShape)
            If textures.initialized Then
                pages(0) = textures.texpage
                pages(1) = textures.tcmaskpage
                pages(2) = textures.normalpage
                pages(3) = textures.specularpage
                For i = 0 To 3
                    If pages(i) <> iV_TEX_INVALID Then
                        ' std::unordered_set<size_t>::insert
                        found = False
                        For j = 0 To n - 1
                            If texPagesToReloadFromDisk(j) = pages(i) Then found = True: Exit For
                        Next
                        If Not found Then
                            If n > UBound(texPagesToReloadFromDisk) Then ReDim Preserve texPagesToReloadFromDisk(0 To 2 * n - 1)
                            texPagesToReloadFromDisk(n) = pages(i)
                            n = n + 1
                        End If
                    End If
                Next
            End If
            pDisplayShape = imdShapes(pDisplayShape).nxt
        Loop
    Next
    debugReloadTexturesFromDisk texPagesToReloadFromDisk, n
End Sub

' (VB) enumerateLoadedModels: the number of models, and the i-th one's name and base shape
Public Function enumerateLoadedModels_count() As Long
    InitModels
    enumerateLoadedModels_count = models.count
End Function

Public Function enumerateLoadedModels_at(ByVal i As Long, modelName_ As String) As Long
    Dim keys As Variant
    keys = models.keys
    modelName_ = keys(i)
    enumerateLoadedModels_at = models(keys(i))
End Function

' (VB) a file's bytes and characters
Private Function LoadImdFile(ByVal path As String, fd As IMD_FILE) As Boolean
    Dim i As Long
    If Not loadFile(path, fd.data, fd.size) Then Exit Function
    fd.chars = String$(fd.size, vbNullChar)
    For i = 0 To fd.size - 1
        Mid$(fd.chars, i + 1, 1) = ChrW$(fd.data(i))
    Next
    LoadImdFile = True
End Function

' the shape handle, or 0
Private Function tryLoadDisplayModelInternal(ByVal path As String, ByVal filename As String, ByVal skipGPUupload As Boolean, Optional ByVal skipDuplicateLoadChecks As Boolean = False) As Long
    Dim fd As IMD_FILE, pFileDataPt As Long
    If PHYSFS_exists(path & filename) <> 0 Then
        If Not LoadImdFile(path & filename, fd) Then
            debug_ 350, LOG_ERROR, "tryLoadDisplayModelInternal", "Failed to load model file: %s", path & filename
            tryLoadDisplayModelInternal = 0
            Exit Function
        End If
        pFileDataPt = 0
        tryLoadDisplayModelInternal = iV_ProcessIMD(filename, fd, pFileDataPt, skipGPUupload, skipDuplicateLoadChecks)
        Exit Function
    End If
    tryLoadDisplayModelInternal = 0
End Function

Private Function tryLoad(ByVal path As String, ByVal filename As String) As Boolean
    Dim graphics_override_model As Long, baseModel As Long, modelName_ As String, baseShape As Long
    InitModels
    If PHYSFS_exists(path & filename) = 0 Then
        ' base model does not exist
        tryLoad = False
        Exit Function
    End If

    ' first, try to load a graphics override - at the same path but with a prefix
    graphics_override_model = tryLoadDisplayModelInternal(WZ_CURRENT_GRAPHICS_OVERRIDES_PREFIX & "/" & path, filename, False)

    ' try to load base model
    baseModel = tryLoadDisplayModelInternal(path, filename, (graphics_override_model <> 0))
    If baseModel = 0 Then
        deleteShape graphics_override_model
        tryLoad = False
        Exit Function
    End If

    ' create BaseShape from full (base) model (first level)
    ' the iIMDBaseShape then "owns" the display model iIMDShape
    modelName_ = imdShapes(baseModel).modelName
    baseShape = newBaseShape(baseModel, path, filename)
    If models.Exists(modelName_) Then
        ASSERT_FAILURE "baseInsertResult.second", "lib/ivis_opengl/imdload.cpp:384", "tryLoad", "%s: Loaded duplicate model? (%s)", filename, modelName_
        deleteBaseShape baseShape
        deleteShape graphics_override_model
        tryLoad = False
        Exit Function
    End If
    models.Add modelName_, baseShape
    ' do NOT use baseModel after this point!

    If graphics_override_model = 0 Then
        ' no graphics override - just use the base model as the display model (the default)
        tryLoad = True
        Exit Function
    End If

    ' there *is* a graphics override version of this model - swap out the base model's displayModel for the graphics override model
    iIMDBaseShape_replaceDisplayModel baseShape, graphics_override_model
    tryLoad = True
End Function

' model: an iIMDShape handle
Public Function modelName(ByVal model As Long) As String
    If model <> 0 Then
        If Not (Len(imdShapes(model).modelName) > 0) Then ASSERT_FAILURE "!model->modelName.isEmpty()", "lib/ivis_opengl/imdload.cpp:402", "modelName", "An IMD pointer could not be backtraced to a filename!"
        If Len(imdShapes(model).modelName) = 0 Then
            debug_ 405, LOG_INFO, "modelName", "Found one"
        End If
        modelName = imdShapes(model).modelName
        Exit Function
    End If
    ASSERT_FAILURE "false", "lib/ivis_opengl/imdload.cpp:409", "modelName", "Null IMD pointer??"
    modelName = ""
End Function

' the iIMDBaseShape handle, or 0
Public Function modelGet(ByVal filename As String) As Long
    Dim name_ As String, dir_ As Variant
    InitModels
    name_ = WzString_toLower(filename)
    If models.Exists(name_) Then
        modelGet = models(name_) ' cached
        Exit Function
    End If
    ' tryLoad("structs/", name) || tryLoad("misc/", name) || ...: the first that loads
    For Each dir_ In Array("structs/", "misc/", "effects/", "components/prop/", "components/weapons/", "components/bodies/", "features/", _
                           "misc/micnum/", "misc/minum/", "misc/mivnum/", "misc/researchimds/")
        If tryLoad(dir_, name_) Then
            modelGet = models(name_)
            Exit Function
        End If
    Next
    debug_ 429, LOG_ERROR, "modelGet", "Could not find: %s", name_
    modelGet = 0
End Function

' iIMDBaseShape::displayModel
Public Function iIMDBaseShape_displayModel(ByVal baseShape As Long) As Long
    iIMDBaseShape_displayModel = imdBaseShapes(baseShape).m_displayModel
End Function

' (VB) a shape's fields, for the other projects
Public Function iIMDShape_objanimpie(ByVal shape As Long, ByVal animEvent As Long) As Long
    iIMDShape_objanimpie = imdShapes(shape).objanimpie(animEvent)
End Function

Public Function iIMDShape_objanimcycles(ByVal shape As Long) As Long
    iIMDShape_objanimcycles = imdShapes(shape).objanimcycles
End Function

Public Function iIMDShape_objanimtime(ByVal shape As Long) As Long
    iIMDShape_objanimtime = imdShapes(shape).objanimtime
End Function

Public Function iIMDShape_objanimframes(ByVal shape As Long) As Long
    iIMDShape_objanimframes = imdShapes(shape).objanimframes
End Function

Public Function iIMDShape_next(ByVal shape As Long) As Long
    iIMDShape_next = imdShapes(shape).nxt
End Function

Public Function iIMDShape_points_size(ByVal shape As Long) As Long
    iIMDShape_points_size = imdShapes(shape).points_size
End Function

Public Sub iIMDShape_point(ByVal shape As Long, ByVal i As Long, x As Single, y As Single, z As Single)
    x = imdShapes(shape).points(i).x
    y = imdShapes(shape).points(i).y
    z = imdShapes(shape).points(i).z
End Sub

' (VB) a display shape's fields, for the exe's display code (shape: an iIMDShape handle)
Public Sub iIMDShape_min(ByVal shape As Long, x As Long, y As Long, z As Long)
    x = imdShapes(shape).min.x
    y = imdShapes(shape).min.y
    z = imdShapes(shape).min.z
End Sub

Public Sub iIMDShape_max(ByVal shape As Long, x As Long, y As Long, z As Long)
    x = imdShapes(shape).max.x
    y = imdShapes(shape).max.y
    z = imdShapes(shape).max.z
End Sub

Public Function iIMDShape_sradius(ByVal shape As Long) As Long
    iIMDShape_sradius = imdShapes(shape).sradius
End Function

Public Function iIMDShape_radius(ByVal shape As Long) As Long
    iIMDShape_radius = imdShapes(shape).radius
End Function

Public Sub iIMDShape_ocen(ByVal shape As Long, x As Single, y As Single, z As Single)
    x = imdShapes(shape).ocen.x
    y = imdShapes(shape).ocen.y
    z = imdShapes(shape).ocen.z
End Sub

Public Function iIMDShape_connectors_size(ByVal shape As Long) As Long
    iIMDShape_connectors_size = imdShapes(shape).connectors_size
End Function

Public Sub iIMDShape_connector(ByVal shape As Long, ByVal i As Long, x As Long, y As Long, z As Long)
    x = imdShapes(shape).connectors(i).x
    y = imdShapes(shape).connectors(i).y
    z = imdShapes(shape).connectors(i).z
End Sub

Public Function iIMDShape_flags(ByVal shape As Long) As Long
    iIMDShape_flags = imdShapes(shape).flags
End Function

Public Function iIMDShape_numFrames(ByVal shape As Long) As Long
    iIMDShape_numFrames = imdShapes(shape).numFrames
End Function

Public Function iIMDShape_animInterval(ByVal shape As Long) As Long
    iIMDShape_animInterval = imdShapes(shape).animInterval
End Function

Public Function iIMDShape_interpolate(ByVal shape As Long) As Long
    iIMDShape_interpolate = imdShapes(shape).interpolate
End Function

Public Function iIMDShape_modelLevel(ByVal shape As Long) As Long
    iIMDShape_modelLevel = imdShapes(shape).modelLevel
End Function

Public Function iIMDShape_polys_size(ByVal shape As Long) As Long
    iIMDShape_polys_size = imdShapes(shape).polys_size
End Function

' objanimdata[i]: scale, pos, rot
Public Sub iIMDShape_objanimdata(ByVal shape As Long, ByVal i As Long, scaleX As Single, scaleY As Single, scaleZ As Single, posX As Single, posY As Single, posZ As Single, rotDirection As Long, rotPitch As Long, rotRoll As Long)
    scaleX = imdShapes(shape).objanimdata(i).scale.x
    scaleY = imdShapes(shape).objanimdata(i).scale.y
    scaleZ = imdShapes(shape).objanimdata(i).scale.z
    posX = imdShapes(shape).objanimdata(i).pos.x
    posY = imdShapes(shape).objanimdata(i).pos.y
    posZ = imdShapes(shape).objanimdata(i).pos.z
    rotDirection = imdShapes(shape).objanimdata(i).rot.direction
    rotPitch = imdShapes(shape).objanimdata(i).rot.pitch
    rotRoll = imdShapes(shape).objanimdata(i).rot.roll
End Sub

' (VB) objanimdata[i] copied to the ANIMFRAME at pOut (the exe's copy of the type: 36 bytes)
Public Sub iIMDShape_objanimdata_p(ByVal shape As Long, ByVal i As Long, ByVal pOut As Long)
    CopyMemory ByVal pOut, imdShapes(shape).objanimdata(i), LenB(imdShapes(shape).objanimdata(i))
End Sub

' (VB) getTextures().texpage, for the exe (bucket3d.cpp's sort key)
Public Function iIMDShape_texpage(ByVal shape As Long) As Long
    iIMDShape_texpage = iIMDShape_getTextures(shape).texpage
End Function

' (VB) a base shape's fields, for the other projects
Public Function iIMDBaseShape_filename(ByVal baseShape As Long) As String
    iIMDBaseShape_filename = imdBaseShapes(baseShape).filename
End Function

Public Function iIMDBaseShape_path(ByVal baseShape As Long) As String
    iIMDBaseShape_path = imdBaseShapes(baseShape).path
End Function

Public Function iIMDBaseShape_sradius(ByVal baseShape As Long) As Long
    iIMDBaseShape_sradius = imdBaseShapes(baseShape).sradius
End Function

Public Function iIMDBaseShape_radius(ByVal baseShape As Long) As Long
    iIMDBaseShape_radius = imdBaseShapes(baseShape).radius
End Function

Public Sub iIMDBaseShape_min(ByVal baseShape As Long, x As Long, y As Long, z As Long)
    x = imdBaseShapes(baseShape).min.x
    y = imdBaseShapes(baseShape).min.y
    z = imdBaseShapes(baseShape).min.z
End Sub

Public Sub iIMDBaseShape_max(ByVal baseShape As Long, x As Long, y As Long, z As Long)
    x = imdBaseShapes(baseShape).max.x
    y = imdBaseShapes(baseShape).max.y
    z = imdBaseShapes(baseShape).max.z
End Sub

Public Sub iIMDBaseShape_ocen(ByVal baseShape As Long, x As Single, y As Single, z As Single)
    x = imdBaseShapes(baseShape).ocen.x
    y = imdBaseShapes(baseShape).ocen.y
    z = imdBaseShapes(baseShape).ocen.z
End Sub

Public Function iIMDBaseShape_connectors_size(ByVal baseShape As Long) As Long
    iIMDBaseShape_connectors_size = imdBaseShapes(baseShape).connectors_size
End Function

Public Sub iIMDBaseShape_connector(ByVal baseShape As Long, ByVal i As Long, x As Long, y As Long, z As Long)
    x = imdBaseShapes(baseShape).connectors(i).x
    y = imdBaseShapes(baseShape).connectors(i).y
    z = imdBaseShapes(baseShape).connectors(i).z
End Sub

'--------------------------------------------------------------------------
' the parser

Private Sub IMD_Line_reset(output As IMD_Line)
    output.lineContents = ""
    output.pNextLineBegin = -1
End Sub

Private Function imd_get_next_line(fd As IMD_FILE, ByVal pFileData As Long, output As IMD_Line) As Boolean
    Dim pCurrLineBegin As Long, pLineEnd As Long, pNextLineBegin As Long, pLineCommentStarted As Long
    Dim getNextLine As Boolean, foundNonWhitespace As Boolean, pCurr As Long, c As Long
    If pFileData = -1 Then
        IMD_Line_reset output
        imd_get_next_line = False
        Exit Function
    End If
    pCurrLineBegin = pFileData
    pLineEnd = -1
    pNextLineBegin = -1
    pLineCommentStarted = -1

    getNextLine = False
    foundNonWhitespace = False ' found non whitespace, non-comment characters
    Do
        For pCurr = pCurrLineBegin To fd.size - 1
            c = fd.data(pCurr)
            If c = 10 Or c = 13 Then
                If pLineEnd = -1 Then pLineEnd = pCurr
            ElseIf pLineEnd <> -1 Then
                ' found a prior newline, and now found the beginning of the next line
                pNextLineBegin = pCurr
                Exit For
            ElseIf c = 35 Then          ' '#'
                If pLineCommentStarted = -1 Then
                    If foundNonWhitespace Then
                        pLineCommentStarted = pCurr
                    Else
                        ' ignore whitespace before a '#' that effectively starts a line
                        pLineCommentStarted = pCurrLineBegin
                    End If
                End If
            ElseIf c <> 32 And c <> 9 Then
                ' found non-whitespace character
                If pLineCommentStarted = -1 Then
                    foundNonWhitespace = True
                End If
            End If
        Next

        If pNextLineBegin = -1 Then
            pNextLineBegin = fd.size
        End If

        If pLineCommentStarted <> -1 And pLineCommentStarted = pCurrLineBegin Then
            ' skip lines that begin with a '#'
            If pLineEnd = -1 Then
                ' last line was a comment - ignore
                IMD_Line_reset output
                imd_get_next_line = False
                Exit Function
            End If
            pCurrLineBegin = pLineEnd
            pLineEnd = -1
            pNextLineBegin = -1
            pLineCommentStarted = -1
            getNextLine = True
            foundNonWhitespace = False
        End If
    Loop While getNextLine

    If pLineCommentStarted <> -1 Then
        pLineEnd = pLineCommentStarted
    End If

    If pLineEnd = -1 Then
        If foundNonWhitespace Then
            ' last line had actual data on it
            pLineEnd = fd.size
            pNextLineBegin = fd.size
        Else
            IMD_Line_reset output
            imd_get_next_line = False ' no more lines
            Exit Function
        End If
    End If
    output.lineContents = Mid$(fd.chars, pFileData + 1, pLineEnd - pFileData)
    output.pNextLineBegin = pNextLineBegin
    imd_get_next_line = True
End Function

Private Function TexCommandIndex(ByVal command As String) As Long
    Select Case command
    Case "TEXTURE": TexCommandIndex = TEX_TEXTURE
    Case "TCMASK": TexCommandIndex = TEX_TCMASK
    Case "NORMALMAP": TexCommandIndex = TEX_NORMALMAP
    Case Else: TexCommandIndex = TEX_SPECULARMAP
    End Select
End Function

Private Function imd_load_texture_command(ByVal filename As String, lineToProcess As IMD_Line, ByVal command As String, ByVal value As Long, ByVal cnt As Long, ByVal expectedCommand As String, levelSettings_ As LevelSettings) As Boolean
    Dim pRestOfLine As String, buffer As String, outputFilename As String
    If command <> expectedCommand Then
        imd_load_texture_command = False
        Exit Function
    End If

    If value < 0 Or value >= 3 Then
        debug_ 565, LOG_ERROR, "_imd_load_texture_command", "%s: %s first parameter (%d) is invalid", filename, expectedCommand, value
        imd_load_texture_command = False
        Exit Function
    End If

    pRestOfLine = Mid$(lineToProcess.lineContents, cnt + 1)
    buffer = ""

    ' the first parameter for textures is always ignored; which is why we ignore
    ' value read in above

    If sscanf(pRestOfLine, "%255s%n", buffer, cnt) <> 1 Then
        debug_ 577, LOG_ERROR, "_imd_load_texture_command", "%s: Texture info corrupt: %s", filename, buffer
        imd_load_texture_command = False
        Exit Function
    End If

    ' verify the extension is .png
    outputFilename = buffer
    If Not strEndsWith(outputFilename, ".png") Then
        debug_ 585, LOG_ERROR, "_imd_load_texture_command", "%s: Only png textures supported", filename
        outputFilename = ""
        imd_load_texture_command = False
        Exit Function
    End If

    levelSettings_.tilesetTextureFiles(value, TexCommandIndex(expectedCommand)) = outputFilename

    ' ignoring old texture width / height directives (which were only for TEXTURE lines)

    imd_load_texture_command = True
End Function

' the lambda getNextPossibleLevelSettingLine of _imd_load_level_settings
Private Function getNextPossibleLevelSettingLine(fd As IMD_FILE, pFileData As Long, ppFileData As Long, lineToProcess As IMD_Line, buffer As String, value As Long, cnt As Long, scanResult As Long) As Boolean
    pFileData = lineToProcess.pNextLineBegin
    ppFileData = pFileData
    If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
        ' no more lines
        getNextPossibleLevelSettingLine = False
        Exit Function
    End If
    scanResult = sscanf(lineToProcess.lineContents, "%255s %d%n", buffer, value, cnt)
    If scanResult <> 2 Then
        ' does not fit format of level settings commands
        getNextPossibleLevelSettingLine = False
        Exit Function
    End If
    getNextPossibleLevelSettingLine = True
End Function

Private Function imd_load_level_settings(ByVal filename As String, fd As IMD_FILE, ppFileData As Long, ByVal pieVersion As Long, ByVal globalScope As Boolean, levelSettings_ As LevelSettings) As Boolean
    Dim pFileData As Long, buffer As String, cnt As Long, value As Long, scanResult As Long, uvalue As Long
    Dim lineToProcess As IMD_Line
    pFileData = ppFileData
    lineToProcess.pNextLineBegin = -1

    ' Every line always starts with <string> <number>
    ' The commands that this function handles (per-level options) are:
    ' - TYPE
    ' - INTERPOLATE
    ' - TEXTURE
    ' - TCMASK
    ' - NORMALMAP
    ' - SPECULARMAP
    ' They are expected in the order above (although all lines are optional, except TYPE - which is currently required at the global scope)

    ' TYPE <hex number>
    If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
        If globalScope Then
            ' TYPE is required in the global scope
            debug_ 623, LOG_ERROR, "_imd_load_level_settings", "%s: Expecting TYPE: %s", filename, buffer
            imd_load_level_settings = False
            Exit Function
        Else
            imd_load_level_settings = True
            Exit Function
        End If
    End If
    scanResult = sscanf(lineToProcess.lineContents, "%255s %x", buffer, uvalue)
    If scanResult <> 2 Or Left$(buffer, 4) <> "TYPE" Then
        If globalScope Then
            ' TYPE is required in the global scope
            debug_ 636, LOG_ERROR, "_imd_load_level_settings", "%s: Expecting TYPE: %s", filename, buffer
            imd_load_level_settings = False
            Exit Function
        Else
            ' non-global scope does not require it - just continue looking for other commands
            lineToProcess.pNextLineBegin = pFileData ' reset to re-examine this line with the standard directive format
        End If
    Else
        ' found TYPE
        levelSettings_.imd_flags = uvalue
        levelSettings_.has_imd_flags = True
    End If

    If Not getNextPossibleLevelSettingLine(fd, pFileData, ppFileData, lineToProcess, buffer, value, cnt, scanResult) Then
        imd_load_level_settings = True
        Exit Function
    End If

    ' INTERPOLATE
    If Left$(buffer, 11) = "INTERPOLATE" Then
        levelSettings_.interpolate = (value > 0)
        levelSettings_.has_interpolate = True
        If Not getNextPossibleLevelSettingLine(fd, pFileData, ppFileData, lineToProcess, buffer, value, cnt, scanResult) Then
            imd_load_level_settings = True
            Exit Function
        End If
    End If

    Do
        ' TEXTURE
        If imd_load_texture_command(filename, lineToProcess, buffer, value, cnt, "TEXTURE", levelSettings_) Then
            If Not getNextPossibleLevelSettingLine(fd, pFileData, ppFileData, lineToProcess, buffer, value, cnt, scanResult) Then
                imd_load_level_settings = True
                Exit Function
            End If
            If pieVersion >= PIE_VER_4 Then
                GoTo continue_ ' loop to allow multiple TEXTURE, TCMASK, NORMALMAP, and SPECULARMAP commands
            End If
        End If

        If pieVersion >= PIE_VER_4 Then
            ' TCMASK (WZ 4.4+ only - earlier versions automatically generated a filename based on the TEXTURE filename)
            If imd_load_texture_command(filename, lineToProcess, buffer, value, cnt, "TCMASK", levelSettings_) Then
                If Not getNextPossibleLevelSettingLine(fd, pFileData, ppFileData, lineToProcess, buffer, value, cnt, scanResult) Then
                    imd_load_level_settings = True
                    Exit Function
                End If
                If pieVersion >= PIE_VER_4 Then
                    GoTo continue_ ' loop to allow multiple TEXTURE, TCMASK, NORMALMAP, and SPECULARMAP commands
                End If
            End If
        End If

        ' NORMALMAP
        If imd_load_texture_command(filename, lineToProcess, buffer, value, cnt, "NORMALMAP", levelSettings_) Then
            If Not getNextPossibleLevelSettingLine(fd, pFileData, ppFileData, lineToProcess, buffer, value, cnt, scanResult) Then
                imd_load_level_settings = True
                Exit Function
            End If
            If pieVersion >= PIE_VER_4 Then
                GoTo continue_ ' loop to allow multiple TEXTURE, TCMASK, NORMALMAP, and SPECULARMAP commands
            End If
        End If

        ' SPECULARMAP
        If imd_load_texture_command(filename, lineToProcess, buffer, value, cnt, "SPECULARMAP", levelSettings_) Then
            If Not getNextPossibleLevelSettingLine(fd, pFileData, ppFileData, lineToProcess, buffer, value, cnt, scanResult) Then
                imd_load_level_settings = True
                Exit Function
            End If
            If pieVersion >= PIE_VER_4 Then
                GoTo continue_ ' loop to allow multiple TEXTURE, TCMASK, NORMALMAP, and SPECULARMAP commands
            End If
        End If

        ' not one of the expected commands - exit loop
        Exit Do
continue_:
    Loop While pieVersion >= PIE_VER_4 ' there may be multiple TEXTURE, TCMASK, NORMALMAP, and SPECULARMAP commands when pieVersion >= PIE_VER_4

    imd_load_level_settings = True
End Function

' Load shape level polygons
' \param ppFileData Pointer to the data (usually read from a file)
' \param s Pointer to shape level
' \return false on error (memory allocation failure/bad file format), true otherwise
Private Function imd_load_polys(ByVal filename As String, fd As IMD_FILE, ppFileData As Long, s As iIMDShape, ByVal pieVersion As Long, ByVal npoints As Long, Optional ByVal shadowPolys As Boolean = False) As Boolean
    Dim pFileData As Long, lineToProcess As IMD_Line, i As Long, flags As Long, npnts As Long, cnt As Long
    Dim pRestOfLine As String, pIdx As Long, nFrames As Long, pbRate As Long, tWidth As Single, tHeight As Single
    Dim j As Long, VertexU As Single, VertexV As Single, p0 As Vector3f, p1 As Vector3f, p2 As Vector3f
    pFileData = ppFileData
    lineToProcess.pNextLineBegin = -1

    s.numFrames = 0
    s.animInterval = 0

    For i = 0 To s.polys_size - 1
        If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
            ' no more lines
            debug_ 770, LOG_ERROR, "_imd_load_polys", "File corrupt - could not read next polygon line (EOF)"
            imd_load_polys = False
            Exit Function
        End If

        pRestOfLine = lineToProcess.lineContents
        If sscanf(pRestOfLine, "%x %u%n", flags, npnts, cnt) <> 2 Then
            debug_ 777, LOG_ERROR, "_imd_load_polys", "(_load_polys) [poly %u] error loading flags and npoints", i
        End If
        pRestOfLine = Mid$(pRestOfLine, cnt + 1)

        If shadowPolys Then
            If Not (flags = 0) Then ASSERT_FAILURE "flags == 0", "lib/ivis_opengl/imdload.cpp:783", "_imd_load_polys", "Invalid polygon flags (%d) for SHADOWPOLYGONS", flags: imd_load_polys = False: Exit Function
        End If

        s.polys(i).flags = flags
        If Not (npnts = 3) Then ASSERT_FAILURE "npnts == 3", "lib/ivis_opengl/imdload.cpp:787", "_imd_load_polys", "Invalid polygon size (%d)", npnts: imd_load_polys = False: Exit Function
        If sscanf(pRestOfLine, "%u%u%u%n", s.polys(i).pindex(0), s.polys(i).pindex(1), s.polys(i).pindex(2), cnt) <> 3 Then
            debug_ 790, LOG_ERROR, "_imd_load_polys", "failed reading triangle, point %d", i
            imd_load_polys = False
            Exit Function
        End If
        pRestOfLine = Mid$(pRestOfLine, cnt + 1)

        ' sanity check
        For pIdx = 0 To 2
            If Not (ULT(s.polys(i).pindex(pIdx), npoints)) Then ASSERT_FAILURE "(poly->pindex[pIdx] < npoints)", "lib/ivis_opengl/imdload.cpp:798", "_imd_load_polys", "Point index (%u) exceeds max index (%u)", s.polys(i).pindex(pIdx), (npoints - 1): imd_load_polys = False: Exit Function
        Next

        ' calc poly normal
        'assumes points already set
        p0 = s.points(s.polys(i).pindex(0))
        p1 = s.points(s.polys(i).pindex(1))
        p2 = s.points(s.polys(i).pindex(2))

        s.polys(i).normal = pie_SurfaceNormal3fv(p0, p1, p2)

        ' texture coord routine
        If (s.polys(i).flags And iV_IMD_TEX) <> 0 Then
            If (s.polys(i).flags And iV_IMD_TEXANIM) <> 0 Then
                If sscanf(pRestOfLine, "%d %d %f %f%n", nFrames, pbRate, tWidth, tHeight, cnt) <> 4 Then
                    debug_ 831, LOG_ERROR, "_imd_load_polys", "(_load_polys) [poly %u] error reading texanim data", i
                    imd_load_polys = False
                    Exit Function
                End If
                pRestOfLine = Mid$(pRestOfLine, cnt + 1)

                If Not (tWidth > 0.0001!) Then ASSERT_FAILURE "tWidth > 0.0001f", "lib/ivis_opengl/imdload.cpp:836", "_imd_load_polys", "%s: texture width = %f", filename, tWidth
                If Not (tHeight > 0!) Then ASSERT_FAILURE "tHeight > 0.f", "lib/ivis_opengl/imdload.cpp:837", "_imd_load_polys", "%s: texture height = %f (width=%f)", filename, tHeight, tWidth
                If Not (nFrames > 1) Then ASSERT_FAILURE "nFrames > 1", "lib/ivis_opengl/imdload.cpp:838", "_imd_load_polys", "%s: animation frames = %d", filename, nFrames
                If Not (pbRate > 0) Then ASSERT_FAILURE "pbRate > 0", "lib/ivis_opengl/imdload.cpp:839", "_imd_load_polys", "%s: animation interval = %d ms", filename, pbRate

                ' Must have same number of frames and same playback rate for all polygons
                If s.numFrames = 0 Then
                    s.numFrames = nFrames And &HFFFF&
                    s.animInterval = pbRate And &HFFFF&
                Else
                    If Not (s.numFrames = nFrames) Then ASSERT_FAILURE "s->numFrames == nFrames", "lib/ivis_opengl/imdload.cpp:849", "_imd_load_polys", _
                        "%s: varying number of frames within one PIE level: %d != %d", filename, nFrames, s.numFrames
                    If Not (s.animInterval = pbRate) Then ASSERT_FAILURE "s->animInterval == pbRate", "lib/ivis_opengl/imdload.cpp:852", "_imd_load_polys", _
                        "%s: varying animation intervals within one PIE level: %d != %d", filename, pbRate, s.animInterval
                End If

                s.polys(i).texAnim.x = tWidth
                s.polys(i).texAnim.y = tHeight

                If pieVersion < PIE_FLOAT_VER Then
                    s.polys(i).texAnim.x = fdiv(s.polys(i).texAnim.x, OLD_TEXTURE_SIZE_FIX)
                    s.polys(i).texAnim.y = fdiv(s.polys(i).texAnim.y, OLD_TEXTURE_SIZE_FIX)
                End If
            Else
                nFrames = 1
                pbRate = 1
                tWidth = 0!
                tHeight = 0!
                s.polys(i).texAnim.x = 0
                s.polys(i).texAnim.y = 0
            End If

            For j = 0 To 2
                If sscanf(pRestOfLine, "%f %f%n", VertexU, VertexV, cnt) <> 2 Then
                    debug_ 882, LOG_ERROR, "_imd_load_polys", "(_load_polys) [poly %u] error reading tex outline", i
                    imd_load_polys = False
                    Exit Function
                End If
                pRestOfLine = Mid$(pRestOfLine, cnt + 1)

                If pieVersion < PIE_FLOAT_VER Then
                    VertexU = fdiv(VertexU, OLD_TEXTURE_SIZE_FIX)
                    VertexV = fdiv(VertexV, OLD_TEXTURE_SIZE_FIX)
                End If

                s.polys(i).texCoord(j).x = VertexU
                s.polys(i).texCoord(j).y = VertexV
            Next
        Else
            If Not ((s.polys(i).flags And iV_IMD_TEXANIM) = 0) Then ASSERT_FAILURE "!(poly->flags & iV_IMD_TEXANIM)", "lib/ivis_opengl/imdload.cpp:900", "_imd_load_polys", "Polygons with texture animation must have textures!": imd_load_polys = False: Exit Function
        End If

        pFileData = lineToProcess.pNextLineBegin
    Next

    ppFileData = pFileData

    imd_load_polys = True
End Function


Private Function ReadPoints(fd As IMD_FILE, ppFileData As Long, s As iIMDShape) As Boolean
    Dim pFileData As Long, cnt As Long, lineToProcess As IMD_Line, k As Long
    pFileData = ppFileData
    lineToProcess.pNextLineBegin = pFileData

    For k = 0 To s.points_size - 1
        If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
            ' no more lines
            debug_ 924, LOG_ERROR, "ReadPoints", "File corrupt - could not read points (EOF)"
            ReadPoints = False
            Exit Function
        End If
        If sscanf(lineToProcess.lineContents, "%f %f %f%n", s.points(k).x, s.points(k).y, s.points(k).z, cnt) <> 3 Then
            debug_ 929, LOG_ERROR, "ReadPoints", "File corrupt - could not read points"
            ReadPoints = False
            Exit Function
        End If
        pFileData = lineToProcess.pNextLineBegin
    Next

    ppFileData = pFileData

    ReadPoints = True
End Function

' (C: static_cast<int>(f), truncation)
Private Function TruncF(ByVal f As Single) As Long
    TruncF = CLng(Fix(f))
End Function

Private Sub imd_calc_bounds(s As iIMDShape)
    ' (allLevels is always false here: only the level's own points)
    Dim xmax As Long, ymax As Long, zmax As Long
    Dim dx As Double, dy As Double, dz As Double, rad_sq As Double, rad As Double, old_to_p_sq As Double, old_to_p As Double, old_to_new As Double
    Dim xspan As Double, yspan As Double, zspan As Double, maxspan As Double
    Dim dia1 As Vector3f, dia2 As Vector3f, cen As Vector3f
    Dim vxmin As Vector3f, vymin As Vector3f, vzmin As Vector3f, vxmax As Vector3f, vymax As Vector3f, vzmax As Vector3f
    Dim k As Long, p As Vector3f

    s.max.x = -FP12_MULTIPLIER: s.max.y = -FP12_MULTIPLIER: s.max.z = -FP12_MULTIPLIER
    s.min.x = FP12_MULTIPLIER: s.min.y = FP12_MULTIPLIER: s.min.z = FP12_MULTIPLIER

    vxmax.x = -FP12_MULTIPLIER: vymax.y = -FP12_MULTIPLIER: vzmax.z = -FP12_MULTIPLIER
    vxmin.x = FP12_MULTIPLIER: vymin.y = FP12_MULTIPLIER: vzmin.z = FP12_MULTIPLIER

    ' set up bounding data for minimum number of vertices
    For k = 0 To s.points_size - 1
        p = s.points(k)
        If p.x > s.max.x Then
            s.max.x = TruncF(p.x)
        End If
        If p.x < s.min.x Then
            s.min.x = TruncF(p.x)
        End If

        If p.y > s.max.y Then
            s.max.y = TruncF(p.y)
        End If
        If p.y < s.min.y Then
            s.min.y = TruncF(p.y)
        End If

        If p.z > s.max.z Then
            s.max.z = TruncF(p.z)
        End If
        If p.z < s.min.z Then
            s.min.z = TruncF(p.z)
        End If

        ' for tight sphere calculations
        If p.x < vxmin.x Then
            vxmin = p
        End If

        If p.x > vxmax.x Then
            vxmax = p
        End If

        If p.y < vymin.y Then
            vymin = p
        End If

        If p.y > vymax.y Then
            vymax = p
        End If

        If p.z < vzmin.z Then
            vzmin = p
        End If

        If p.z > vzmax.z Then
            vzmax = p
        End If
    Next

    ' no need to scale an IMD shape (only FSD)
    xmax = MAX_(s.max.x, -s.min.x)
    ymax = MAX_(s.max.y, -s.min.y)
    zmax = MAX_(s.max.z, -s.min.z)

    s.radius = MAX_(xmax, MAX_(ymax, zmax))
    s.sradius = TruncF(fsqrt(F32(CDbl(xmax * xmax + ymax * ymax + zmax * zmax))))

' START: tight bounding sphere

    ' set xspan = distance between 2 points xmin & xmax (squared)
    dx = fsub(vxmax.x, vxmin.x)
    dy = fsub(vxmax.y, vxmin.y)
    dz = fsub(vxmax.z, vxmin.z)
    xspan = dx * dx + dy * dy + dz * dz

    ' same for yspan
    dx = fsub(vymax.x, vymin.x)
    dy = fsub(vymax.y, vymin.y)
    dz = fsub(vymax.z, vymin.z)
    yspan = dx * dx + dy * dy + dz * dz

    ' and ofcourse zspan
    dx = fsub(vzmax.x, vzmin.x)
    dy = fsub(vzmax.y, vzmin.y)
    dz = fsub(vzmax.z, vzmin.z)
    zspan = dx * dx + dy * dy + dz * dz

    ' set points dia1 & dia2 to maximally separated pair
    ' assume xspan biggest
    dia1 = vxmin
    dia2 = vxmax
    maxspan = xspan

    If yspan > maxspan Then
        maxspan = yspan
        dia1 = vymin
        dia2 = vymax
    End If

    If zspan > maxspan Then
        dia1 = vzmin
        dia2 = vzmax
    End If

    ' dia1, dia2 diameter of initial sphere
    cen.x = fdiv(fadd(dia1.x, dia2.x), 2!)
    cen.y = fdiv(fadd(dia1.y, dia2.y), 2!)
    cen.z = fdiv(fadd(dia1.z, dia2.z), 2!)

    ' calc initial radius
    dx = fsub(dia2.x, cen.x)
    dy = fsub(dia2.y, cen.y)
    dz = fsub(dia2.z, cen.z)

    rad_sq = dx * dx + dy * dy + dz * dz
    rad = Sqr(rad_sq)

    ' second pass (find tight sphere)
    For k = 0 To s.points_size - 1
        p = s.points(k)
        dx = fsub(p.x, cen.x)
        dy = fsub(p.y, cen.y)
        dz = fsub(p.z, cen.z)
        old_to_p_sq = dx * dx + dy * dy + dz * dz

        ' do r**2 first
        If old_to_p_sq > rad_sq Then
            ' this point outside current sphere
            old_to_p = Sqr(old_to_p_sq)
            ' radius of new sphere
            rad = (rad + old_to_p) / 2#
            ' rad**2 for next compare
            rad_sq = rad * rad
            old_to_new = old_to_p - rad
            ' centre of new sphere
            cen.x = F32((rad * cen.x + old_to_new * p.x) / old_to_p)
            cen.y = F32((rad * cen.y + old_to_new * p.y) / old_to_p)
            cen.z = F32((rad * cen.z + old_to_new * p.z) / old_to_p)
        End If
    Next

    s.ocen = cen

' END: tight bounding sphere
End Sub

Private Function MAX_(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MAX_ = a Else MAX_ = b
End Function

Private Function imd_load_points(fd As IMD_FILE, ppFileData As Long, s As iIMDShape, ByVal npoints As Long) As Boolean
    'load the points then pass through a second time to setup bounding datavalues
    s.points_size = npoints
    If npoints > 0 Then ReDim s.points(0 To npoints - 1) Else Erase s.points

    ' Read in points and remove duplicates (!)
    If ReadPoints(fd, ppFileData, s) = False Then
        Erase s.points
        s.points_size = 0
        imd_load_points = False
        Exit Function
    End If

    imd_calc_bounds s

    imd_load_points = True
End Function


' Load shape level connectors
' \param ppFileData Pointer to the data (usually read from a file)
' \param s Pointer to shape level
' \return false on error (memory allocation failure/bad file format), true otherwise
' \pre ppFileData loaded
' \pre s allocated
' \post s->connectors populated
Private Function imd_load_connectors(fd As IMD_FILE, ppFileData As Long, ByVal numConnectors As Long, s As iIMDShape) As Boolean
    Dim pFileData As Long, cnt As Long, lineToProcess As IMD_Line, newVector As Vector3i, i As Long
    pFileData = ppFileData
    lineToProcess.pNextLineBegin = pFileData

    For i = 0 To numConnectors - 1
        If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
            ' no more lines
            debug_ 1165, LOG_ERROR, "_imd_load_connectors", "(_load_connectors) file corrupt - could not read any more lines (EOF)"
            imd_load_connectors = False
            Exit Function
        End If
        If sscanf(lineToProcess.lineContents, "%d %d %d%n", newVector.x, newVector.y, newVector.z, cnt) <> 3 Then
            If sscanf(lineToProcess.lineContents, "%d%*[.0-9] %d%*[.0-9] %d%*[.0-9]%n", newVector.x, newVector.y, newVector.z, cnt) <> 3 Then
                debug_ 1171, LOG_ERROR, "_imd_load_connectors", "(_load_connectors) file corrupt -M"
                imd_load_connectors = False
                Exit Function
            End If
        End If
        pFileData = lineToProcess.pNextLineBegin
        ReDim Preserve s.connectors(0 To s.connectors_size)
        s.connectors(s.connectors_size) = newVector
        s.connectors_size = s.connectors_size + 1
    Next

    ppFileData = pFileData

    imd_load_connectors = True
End Function

Private Function imd_get_level_tileset_texture_file(ByVal typ As Long, levelTilesetSettings_ As LevelSettings, globalTilesetSettings_ As LevelSettings, ByVal tilesetIdx As Long) As String
    ' 1.) try the level tileset settings
    If Len(levelTilesetSettings_.tilesetTextureFiles(tilesetIdx, typ)) > 0 Then
        imd_get_level_tileset_texture_file = levelTilesetSettings_.tilesetTextureFiles(tilesetIdx, typ)
        Exit Function
    End If
    ' 2.) fall back to the global tileset settings
    If Len(globalTilesetSettings_.tilesetTextureFiles(tilesetIdx, typ)) > 0 Then
        imd_get_level_tileset_texture_file = globalTilesetSettings_.tilesetTextureFiles(tilesetIdx, typ)
        Exit Function
    End If
    imd_get_level_tileset_texture_file = ""
End Function

Private Sub imd_determine_tileset_texture_files(s As iIMDShape, globalLevelSettings As LevelSettings, levelSettings_ As LevelSettings)
    Dim tilesetIdx As Long
    For tilesetIdx = 0 To 2
        s.tilesetTextureFiles(tilesetIdx).texfile = imd_get_level_tileset_texture_file(TEX_TEXTURE, levelSettings_, globalLevelSettings, tilesetIdx)
        s.tilesetTextureFiles(tilesetIdx).tcmaskfile = imd_get_level_tileset_texture_file(TEX_TCMASK, levelSettings_, globalLevelSettings, tilesetIdx)
        s.tilesetTextureFiles(tilesetIdx).normalfile = imd_get_level_tileset_texture_file(TEX_NORMALMAP, levelSettings_, globalLevelSettings, tilesetIdx)
        s.tilesetTextureFiles(tilesetIdx).specfile = imd_get_level_tileset_texture_file(TEX_SPECULARMAP, levelSettings_, globalLevelSettings, tilesetIdx)
    Next
End Sub

' static bool _imd_load_level_textures(const iIMDShape& s, size_t tilesetIdx, iIMDShapeTextures& output)
' (s: an iIMDShape handle; a TilesetTextureFiles* is the index of the settings to use: 0 the default, else tilesetIdx)
Private Function imd_load_level_textures_(ByVal s As Long, ByVal tilesetIdx As Long, output As iIMDShapeTextures) As Boolean
    Dim filename As String, pLevelSettingsToUseForTextures As Long, pLevelSettingsToUseForTCMask As Long
    Dim pLevelSettingsToUseForNormals As Long, pLevelSettingsToUseForSpecular As Long
    Dim texpage As Long, tcmaskpage As Long, normalpage As Long, specpage As Long
    Dim has_texpage As Boolean, has_tcmaskpage As Boolean, has_normalpage As Boolean, has_specpage As Boolean
    Dim tcmask_name As String, texfile As String
    ' const auto& defaultSettings = s.tilesetTextureFiles[0]; const auto& tilesetSettings = s.tilesetTextureFiles[tilesetIdx];

    filename = imdShapes(s).modelName
    If Len(imdShapes(s).tilesetTextureFiles(tilesetIdx).texfile) <> 0 Then pLevelSettingsToUseForTextures = tilesetIdx Else pLevelSettingsToUseForTextures = 0
    texfile = imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForTextures).texfile
    If Len(texfile) <> 0 Then
        has_texpage = iV_GetTexture(texfile, texture_type_game_texture, texpage)

        If Not has_texpage Then ASSERT_FAILURE "texpage.has_value()", "lib/ivis_opengl/imdload.cpp:1228", "_imd_load_level_textures", "%s could not load tex page %s", filename, texfile: Exit Function

        If Len(imdShapes(s).tilesetTextureFiles(tilesetIdx).tcmaskfile) <> 0 Then pLevelSettingsToUseForTCMask = tilesetIdx Else pLevelSettingsToUseForTCMask = 0
        If Len(imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForTCMask).tcmaskfile) <> 0 Then
            ' load explicitly specified tcmask file
            debug_ 1234, LOG_TEXTURE, "_imd_load_level_textures", "Loading tcmask %s for %s", imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForTCMask).tcmaskfile, filename
            has_tcmaskpage = iV_GetTexture(imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForTCMask).tcmaskfile, texture_type_alpha_mask, tcmaskpage)
            If Not has_tcmaskpage Then ASSERT_FAILURE "tcmaskpage.has_value()", "lib/ivis_opengl/imdload.cpp:1236", "_imd_load_level_textures", "%s could not load tcmask %s", filename, imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForTCMask).tcmaskfile: Exit Function
        Else
            ' BACKWARDS-COMPATIBILITY (PIE 2/3 compatibility)
            ' check if model should use team colour mask
            If (imdShapes(s).flags And iV_IMD_TCMASK) <> 0 Then
                tcmask_name = pie_MakeTexPageTCMaskName(texfile)
                tcmask_name = tcmask_name & ".png"
                has_tcmaskpage = iV_GetTexture(tcmask_name, texture_type_alpha_mask, tcmaskpage)
                If Not has_tcmaskpage Then ASSERT_FAILURE "tcmaskpage.has_value()", "lib/ivis_opengl/imdload.cpp:1247", "_imd_load_level_textures", "%s could not load tcmask %s", filename, tcmask_name: Exit Function
            End If
        End If

        If Len(imdShapes(s).tilesetTextureFiles(tilesetIdx).normalfile) <> 0 Then pLevelSettingsToUseForNormals = tilesetIdx Else pLevelSettingsToUseForNormals = 0
        If Len(imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForNormals).normalfile) <> 0 Then
            debug_ 1254, LOG_TEXTURE, "_imd_load_level_textures", "Loading normal map %s for %s", imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForNormals).normalfile, filename
            has_normalpage = iV_GetTexture(imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForNormals).normalfile, texture_type_normal_map, normalpage)
            If Not has_normalpage Then ASSERT_FAILURE "normalpage.has_value()", "lib/ivis_opengl/imdload.cpp:1256", "_imd_load_level_textures", "%s could not load tex page %s", filename, imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForNormals).normalfile: Exit Function
        End If

        If Len(imdShapes(s).tilesetTextureFiles(tilesetIdx).specfile) <> 0 Then pLevelSettingsToUseForSpecular = tilesetIdx Else pLevelSettingsToUseForSpecular = 0
        If Len(imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForSpecular).specfile) <> 0 Then
            debug_ 1262, LOG_TEXTURE, "_imd_load_level_textures", "Loading specular map %s for %s", imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForSpecular).specfile, filename
            has_specpage = iV_GetTexture(imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForSpecular).specfile, texture_type_specular_map, specpage)
            If Not has_specpage Then ASSERT_FAILURE "specpage.has_value()", "lib/ivis_opengl/imdload.cpp:1264", "_imd_load_level_textures", "%s could not load tex page %s", filename, imdShapes(s).tilesetTextureFiles(pLevelSettingsToUseForSpecular).specfile: Exit Function
        End If

        ' assign tex pages and flags for this level
        output.texpage = texpage
        If has_tcmaskpage Then output.tcmaskpage = tcmaskpage Else output.tcmaskpage = iV_TEX_INVALID
        If has_normalpage Then output.normalpage = normalpage Else output.normalpage = iV_TEX_INVALID
        If has_specpage Then output.specularpage = specpage Else output.specularpage = iV_TEX_INVALID
    End If

    imd_load_level_textures_ = True
End Function

Private Function ReadNormals(fd As IMD_FILE, ppFileData As Long, pie_level_normals() As Vector3f, ByVal num_normal_lines As Long) As Boolean
    Dim pFileData As Long, dataCnt As Long, lineToProcess As IMD_Line, normalLine As Long, i As Long
    Dim pRestOfLine As String, k As Long
    pFileData = ppFileData
    lineToProcess.pNextLineBegin = pFileData

    For normalLine = 0 To num_normal_lines - 1
        If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
            ' no more lines
            debug_ 1300, LOG_ERROR, "ReadNormals", "File corrupt - could not read normals (EOF)"
            ReadNormals = False
            Exit Function
        End If
        pRestOfLine = lineToProcess.lineContents
        For i = 0 To 2
            k = normalLine * 3 + i
            If sscanf(pRestOfLine, "%f %f %f%n", pie_level_normals(k).x, pie_level_normals(k).y, pie_level_normals(k).z, dataCnt) <> 3 Then
                debug_ 1309, LOG_ERROR, "ReadNormals", "File corrupt - could not read normal(%zu) from normals line: %zu", i, normalLine
                ReadNormals = False
                Exit Function
            End If
            pRestOfLine = Mid$(pRestOfLine, dataCnt + 1)
        Next
        pFileData = lineToProcess.pNextLineBegin
    Next

    ppFileData = pFileData

    ReadNormals = True
End Function

Private Function imd_load_normals(fd As IMD_FILE, ppFileData As Long, pie_level_normals() As Vector3f, pie_level_normals_size As Long, ByVal num_normal_lines As Long) As Boolean
    ' We only support triangles!
    pie_level_normals_size = num_normal_lines * 3
    If pie_level_normals_size > 0 Then ReDim pie_level_normals(0 To pie_level_normals_size - 1) Else Erase pie_level_normals

    If ReadNormals(fd, ppFileData, pie_level_normals, num_normal_lines) = False Then
        Erase pie_level_normals
        pie_level_normals_size = 0
        imd_load_normals = False
        Exit Function
    End If

    imd_load_normals = True
End Function

' (VB) push_back on the static vectors
Private Sub PushF(v() As Single, size As Long, ByVal value As Single)
    If size > UBound(v) Then ReDim Preserve v(0 To size * 2 + 63)
    v(size) = value
    size = size + 1
End Sub

Private Function addVertex(s As iIMDShape, ByVal i As Long, p As iIMDPoly, ByVal polyIdx As Long, pie_level_normals() As Vector3f, ByVal pie_level_normals_size As Long) As Long
    Dim normal As Vector3f, j As Long, pt As Vector3f

    pt = s.points(p.pindex(i))
    ' Do not weld for for models with normals, those are presumed to be correct. Otherwise, it will break tangents
    If pie_level_normals_size = 0 Then
        normal = p.normal
        ' See if we already have this defined, if so, return reference to it.
        For j = 0 To vertexCount - 1
            If texcoords(j * 4 + 0) = p.texCoord(i).x Then
            If texcoords(j * 4 + 1) = p.texCoord(i).y Then
            If texcoords(j * 4 + 2) = p.texAnim.x Then
            If texcoords(j * 4 + 3) = p.texAnim.y Then
            If vertices(j * 3 + 0) = pt.x Then
            If vertices(j * 3 + 1) = pt.y Then
            If vertices(j * 3 + 2) = pt.z Then
            If normals(j * 3 + 0) = normal.x Then
            If normals(j * 3 + 1) = normal.y Then
            If normals(j * 3 + 2) = normal.z Then
                addVertex = j
                Exit Function
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
        Next
    Else
        normal = pie_level_normals(polyIdx * 3 + i)
    End If

    ' We don't have it, add it.
    PushF normals, normals_size, normal.x
    PushF normals, normals_size, normal.y
    PushF normals, normals_size, normal.z
    PushF texcoords, texcoords_size, p.texCoord(i).x
    PushF texcoords, texcoords_size, p.texCoord(i).y
    PushF texcoords, texcoords_size, p.texAnim.x
    PushF texcoords, texcoords_size, p.texAnim.y
    PushF vertices, vertices_size, pt.x
    PushF vertices, vertices_size, pt.y
    PushF vertices, vertices_size, pt.z
    vertexCount = (vertexCount + 1) And &HFFFF&

    addVertex = (vertexCount - 1) And &HFFFF&
End Function

Private Sub calculateTangentsForTriangle(ByVal ia As Long, ByVal ib As Long, ByVal ic As Long)
    Dim va As Vector3f, vb As Vector3f, vc As Vector3f, deltaPos1 As Vector3f, deltaPos2 As Vector3f
    Dim uva_x As Single, uva_y As Single, uvb_x As Single, uvb_y As Single, uvc_x As Single, uvc_y As Single
    Dim deltaUV1_x As Single, deltaUV1_y As Single, deltaUV2_x As Single, deltaUV2_y As Single
    Dim r As Single, tangent As Vector3f, bitangent As Vector3f, k As Long, idx As Variant

    ' Shortcuts for vertices
    va = vec3f(vertices(ia * 3), vertices(ia * 3 + 1), vertices(ia * 3 + 2))
    vb = vec3f(vertices(ib * 3), vertices(ib * 3 + 1), vertices(ib * 3 + 2))
    vc = vec3f(vertices(ic * 3), vertices(ic * 3 + 1), vertices(ic * 3 + 2))

    ' Shortcuts for UVs
    uva_x = texcoords(ia * 4): uva_y = texcoords(ia * 4 + 1)
    uvb_x = texcoords(ib * 4): uvb_y = texcoords(ib * 4 + 1)
    uvc_x = texcoords(ic * 4): uvc_y = texcoords(ic * 4 + 1)

    ' Edges of the triangle : postion delta
    deltaPos1 = v3sub(vb, va)
    deltaPos2 = v3sub(vc, va)

    ' UV delta
    deltaUV1_x = fsub(uvb_x, uva_x): deltaUV1_y = fsub(uvb_y, uva_y)
    deltaUV2_x = fsub(uvc_x, uva_x): deltaUV2_y = fsub(uvc_y, uva_y)

    ' check for nan
    r = fsub(fmul(deltaUV1_x, deltaUV2_y), fmul(deltaUV1_y, deltaUV2_x))
    If r <> 0! Then r = fdiv(1!, r)

    tangent = v3scale(v3sub(v3scale(deltaPos1, deltaUV2_y), v3scale(deltaPos2, deltaUV1_y)), r)
    bitangent = v3scale(v3sub(v3scale(deltaPos2, deltaUV1_x), v3scale(deltaPos1, deltaUV2_x)), r)

    For Each idx In Array(ia, ib, ic)
        k = idx
        tangents(k * 4) = fadd(tangents(k * 4), tangent.x)
        tangents(k * 4 + 1) = fadd(tangents(k * 4 + 1), tangent.y)
        tangents(k * 4 + 2) = fadd(tangents(k * 4 + 2), tangent.z)
        tangents(k * 4 + 3) = fadd(tangents(k * 4 + 3), 0!)
        bitangents(k * 3) = fadd(bitangents(k * 3), bitangent.x)
        bitangents(k * 3 + 1) = fadd(bitangents(k * 3 + 1), bitangent.y)
        bitangents(k * 3 + 2) = fadd(bitangents(k * 3 + 2), bitangent.z)
    Next
End Sub

Private Sub finishTangentsGeneration()
    Dim i As Long, n As Vector3f, b As Vector3f, t As Vector3f, d As Single
    For i = 0 To vertexCount - 1
        n = vec3f(normals(i * 3), normals(i * 3 + 1), normals(i * 3 + 2))
        b = vec3f(bitangents(i * 3), bitangents(i * 3 + 1), bitangents(i * 3 + 2))
        t = vec3f(tangents(i * 4), tangents(i * 4 + 1), tangents(i * 4 + 2))

        ' Gram-Schmidt orthogonalize
        t = v3normalize(v3sub(t, v3scale(n, v3dot(n, t))))

        ' Calculate handedness
        tangents(i * 4) = t.x
        tangents(i * 4 + 1) = t.y
        tangents(i * 4 + 2) = t.z
        d = v3dot(v3cross(n, t), b)
        If IsNaNF(d) Then                   ' (VB) a zero tangent normalizes to NaNs, which are not < 0
            tangents(i * 4 + 3) = -1!
        ElseIf d < 0! Then
            tangents(i * 4 + 3) = 1!
        Else
            tangents(i * 4 + 3) = -1!
        End If
    Next
End Sub

' the lambdas getNextPossibleIgnoredCommand and getNextPossibleCommandLine of _imd_load_level
Private Function getNextPossibleIgnoredCommand(fd As IMD_FILE, pFileData As Long, lineToProcess As IMD_Line, buffer As String, cnt As Long, scanResult As Long) As Boolean
    pFileData = lineToProcess.pNextLineBegin
    If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
        ' no more lines
        getNextPossibleIgnoredCommand = False
        Exit Function
    End If
    scanResult = sscanf(lineToProcess.lineContents, "%255s %n", buffer, cnt)
    If scanResult <> 1 Then
        getNextPossibleIgnoredCommand = False
        Exit Function
    End If
    getNextPossibleIgnoredCommand = True
End Function

Private Function getNextPossibleCommandLine(fd As IMD_FILE, pFileData As Long, lineToProcess As IMD_Line, buffer As String, value As Long, cnt As Long, scanResult As Long) As Boolean
    pFileData = lineToProcess.pNextLineBegin
    If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
        ' no more lines
        getNextPossibleCommandLine = False
        Exit Function
    End If
    scanResult = sscanf(lineToProcess.lineContents, "%255s %u%n", buffer, value, cnt)
    If scanResult <> 2 Then
        getNextPossibleCommandLine = False
        Exit Function
    End If
    getNextPossibleCommandLine = True
End Function

' (VB) if (!s.buffers[vbo]) s.buffers[vbo] = create_buffer_object(usage, static_draw, debugName);
' s.buffers[vbo]->upload(size * sizeof(gfx_api::gfxFloat), data.data()); -- skipped with no gfx context
' (the --vbtest runs, where C++'s context::get() would have none)
Private Sub uploadShapeBufferF(ByVal h As Long, ByVal vbo As VBO_TYPE, ByVal usage As buffer_usage, ByVal debugName As String, src() As Single, ByVal size As Long)
    If imdShapes(h).buffers(vbo) Is Nothing Then
        If current_backend_context = CONTEXT_NONE Then Exit Sub
        Set imdShapes(h).buffers(vbo) = gfx_api_context_create_buffer_object(usage, buffer_storage_hint_static_draw, debugName)
    End If
    If size > 0 Then
        imdShapes(h).buffers(vbo).upload size * 4, VarPtr(src(0))
    Else
        imdShapes(h).buffers(vbo).upload 0, 0
    End If
End Sub

' (VB) the index buffer: the indices as uint16_t
Private Sub uploadShapeIndexBuffer(ByVal h As Long, src() As Long, ByVal size As Long)
    Dim idx16() As Integer, i As Long
    If imdShapes(h).buffers(VBO_INDEX) Is Nothing Then
        If current_backend_context = CONTEXT_NONE Then Exit Sub
        Set imdShapes(h).buffers(VBO_INDEX) = gfx_api_context_create_buffer_object(buffer_usage_index_buffer, buffer_storage_hint_static_draw, "index buffer")
    End If
    If size > 0 Then
        ReDim idx16(0 To size - 1)
        For i = 0 To size - 1
            If src(i) > 32767 Then idx16(i) = src(i) - 65536 Else idx16(i) = src(i)
        Next
        imdShapes(h).buffers(VBO_INDEX).upload size * 2, VarPtr(idx16(0))
    Else
        imdShapes(h).buffers(VBO_INDEX).upload 0, 0
    End If
End Sub

' (VB) "uploads" a static vector into the shape (C++: into a GPU buffer)
Private Sub UploadF(dst() As Single, dstSize As Long, src() As Single, ByVal size As Long)
    dstSize = size
    If size > 0 Then
        dst = src
        ReDim Preserve dst(0 To size - 1)
    Else
        Erase dst
    End If
End Sub

' Load shape levels recursively
' \param ppFileData Pointer to the data (usually read from a file)
' \param FileDataEnd ???
' \param nlevels Number of levels to load
' \return pointer to iFSDShape structure (or NULL on error)
' \pre ppFileData loaded
' \post s allocated
Private Function imd_load_level(ByVal filename As String, fd As IMD_FILE, ppFileData As Long, ByVal pieVersion As Long, ByVal level As Long, globalLevelSettings As LevelSettings, ByVal skipGPUData As Boolean, ByVal skipDuplicateLoadChecks As Boolean) As Long
    Dim pFileData As Long, buffer As String, value As Long, cnt As Long, scanResult As Long
    Dim key As String, pAllocatedShape As Long, levelSettings_ As LevelSettings, lineToProcess As IMD_Line
    Dim npoints As Long, npolys As Long, pie_level_normals() As Vector3f, pie_level_normals_size As Long
    Dim pRestOfLine As String, i As Long, frame As Long, pos As Vector3i, rot As Vector3i
    Dim nShadowPoints As Long, nShadowPolys As Long, tmpShadowShape As iIMDShape, blankShape As iIMDShape
    Dim npol As Long, flags As Long

    pFileData = ppFileData
    buffer = ""

    ' insert model
    key = filename
    If level > 0 Then
        key = key & "_" & CStr(level)
    End If
    If Not skipDuplicateLoadChecks Then
        InitModels
        If Not (Not models.Exists(key)) Then ASSERT_FAILURE "models.count(key) == 0", "lib/ivis_opengl/imdload.cpp:1480", "_imd_load_level", "Duplicate model load for %s!", key
    End If
    pAllocatedShape = newShape()
    imdShapes(pAllocatedShape).modelName = key
    imdShapes(pAllocatedShape).modelLevel = level
    imdShapes(pAllocatedShape).shadowUsesAlt = False

    If pieVersion >= PIE_VER_4 Then
        ' While level settings are supported at the global level, each level can also specify its own settings overrides (as of WZ 4.4+) which take precedence
        If Not imd_load_level_settings(filename, fd, pFileData, pieVersion, False, levelSettings_) Then
            debug_ 1495, LOG_ERROR, "_imd_load_level", "%s: Failed to load level settings (level: %u)", filename, level
            GoTo fail
        End If
    End If

    lineToProcess.pNextLineBegin = pFileData

    getNextPossibleIgnoredCommand fd, pFileData, lineToProcess, buffer, cnt, scanResult
    If Not (scanResult = 1) Then ASSERT_FAILURE "scanResult == 1", "lib/ivis_opengl/imdload.cpp:1534", "_imd_load_level", "Bad directive following LEVEL": GoTo fail

    ' Optionally load and ignore deprecated MATERIALS directive
    If buffer = "MATERIALS" Then
        debug_ 1539, LOG_WARNING, "_imd_load_level", "MATERIALS directive no longer supported!"
        If Not getNextPossibleIgnoredCommand(fd, pFileData, lineToProcess, buffer, cnt, scanResult) Then
            debug_ 1542, LOG_ERROR, "_imd_load_level", "_imd_load_level(1a): file corrupt: %s", filename
            GoTo fail
        End If
    End If
    If buffer = "SHADERS" Then
        debug_ 1548, LOG_WARNING, "_imd_load_level", "SHADERS directive no longer supported!"
        If Not getNextPossibleIgnoredCommand(fd, pFileData, lineToProcess, buffer, cnt, scanResult) Then
            debug_ 1551, LOG_ERROR, "_imd_load_level", "_imd_load_level(1b): file corrupt: %s", filename
            GoTo fail
        End If
    End If

    lineToProcess.pNextLineBegin = pFileData ' re-process the directive
    If Not getNextPossibleCommandLine(fd, pFileData, lineToProcess, buffer, value, cnt, scanResult) Then
        debug_ 1559, LOG_ERROR, "_imd_load_level", "_imd_load_level(2): file corrupt: %s", filename
        GoTo fail
    End If

    ' load points
    If Not (buffer = "POINTS") Then ASSERT_FAILURE "strcmp(buffer, ""POINTS"") == 0", "lib/ivis_opengl/imdload.cpp:1564", "_imd_load_level", "Expecting 'POINTS' directive, got: %s", buffer: GoTo fail
    npoints = value
    If Not imd_load_points(fd, lineToProcess.pNextLineBegin, imdShapes(pAllocatedShape), npoints) Then
        debug_ 1568, LOG_ERROR, "_imd_load_level", "_imd_load_level(2a): file corrupt: %s", filename
        GoTo fail
    End If

    If Not getNextPossibleCommandLine(fd, pFileData, lineToProcess, buffer, value, cnt, scanResult) Then
        debug_ 1574, LOG_ERROR, "_imd_load_level", "_imd_load_level(3): file corrupt: %s", filename
        GoTo fail
    End If

    npolys = value

    ' It could be optional normals directive
    If buffer = "NORMALS" Then
        imd_load_normals fd, lineToProcess.pNextLineBegin, pie_level_normals, pie_level_normals_size, npolys

        ' Attemps to read polys again
        If Not getNextPossibleCommandLine(fd, pFileData, lineToProcess, buffer, value, cnt, scanResult) Then
            debug_ 1589, LOG_ERROR, "_imd_load_level", "_imd_load_level(3a): file corrupt: %s", filename
            GoTo fail
        End If
        npolys = value
        debug_ 1593, LOG_3D, "_imd_load_level", "imd[_load_level] = normals %d", pie_level_normals_size
    End If

    If Not (buffer = "POLYGONS") Then ASSERT_FAILURE "strcmp(buffer, ""POLYGONS"") == 0", "lib/ivis_opengl/imdload.cpp:1596", "_imd_load_level", "Expecting 'POLYGONS' directive, got: %s", buffer: GoTo fail
    If Not (Not ULT(MAX_PIE_POLYGONS, npolys)) Then ASSERT_FAILURE "npolys <= MAX_PIE_POLYGONS", "lib/ivis_opengl/imdload.cpp:1597", "_imd_load_level", "'POLYGONS' directive count (%u) exceeds maximum supported (%u)", npolys, MAX_PIE_POLYGONS: GoTo fail
    imdShapes(pAllocatedShape).polys_size = npolys
    If npolys > 0 Then ReDim imdShapes(pAllocatedShape).polys(0 To npolys - 1) Else Erase imdShapes(pAllocatedShape).polys

    If Not imd_load_polys(filename, fd, lineToProcess.pNextLineBegin, imdShapes(pAllocatedShape), pieVersion, npoints) Then
        debug_ 1602, LOG_ERROR, "_imd_load_level", "_imd_load_level(3b): file corrupt - invalid polys: %s", filename
        GoTo fail
    End If

    ' optional stuff : levels, object animations, connectors, shadow points + polys
    imdShapes(pAllocatedShape).objanimframes = 0
    Do While getNextPossibleCommandLine(fd, pFileData, lineToProcess, buffer, value, cnt, scanResult)
        If buffer = "LEVEL" Then ' check for next level
            debug_ 1612, LOG_3D, "_imd_load_level", "imd[_load_level] = npoints %u, npolys %u", npoints, npolys
            ' Quit this loop - parent will call _imd_load_level again to process another LEVEL section
            Exit Do
        ElseIf buffer = "CONNECTORS" Then
            'load connector stuff
            If Not imd_load_connectors(fd, lineToProcess.pNextLineBegin, value, imdShapes(pAllocatedShape)) Then
                debug_ 1621, LOG_ERROR, "_imd_load_level", "_imd_load_level(5): file corrupt - invalid shadowpoints: %s", filename
                GoTo fail
            End If
        ElseIf buffer = "ANIMOBJECT" Then
            imdShapes(pAllocatedShape).objanimtime = value
            If imdShapes(pAllocatedShape).objanimtime = 0 Then
                debug_ 1630, LOG_ERROR, "_imd_load_level", "%s bad ANIMOBJ time: %u", filename, value
                GoTo fail
            End If
            pRestOfLine = Mid$(lineToProcess.lineContents, cnt + 1)
            If sscanf(pRestOfLine, "%d %d%n", imdShapes(pAllocatedShape).objanimcycles, imdShapes(pAllocatedShape).objanimframes, cnt) <> 2 Then
                ' (pFileData as a C string: the rest of the file)
                debug_ 1636, LOG_ERROR, "_imd_load_level", "%s bad ANIMOBJ: %s", filename, Mid$(fd.chars, pFileData + 1)
                GoTo fail
            End If
            pFileData = lineToProcess.pNextLineBegin
            If imdShapes(pAllocatedShape).objanimframes > 0 Then
                ReDim imdShapes(pAllocatedShape).objanimdata(0 To imdShapes(pAllocatedShape).objanimframes - 1)
            Else
                Erase imdShapes(pAllocatedShape).objanimdata
            End If
            For i = 0 To imdShapes(pAllocatedShape).objanimframes - 1
                If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
                    ' no more lines
                    debug_ 1646, LOG_ERROR, "_imd_load_level", "%s: Missing ANIMOBJECT frame line (EOF)", filename
                    GoTo fail
                End If

                pos.x = 0: pos.y = 0: pos.z = 0
                rot.x = 0: rot.y = 0: rot.z = 0

                With imdShapes(pAllocatedShape).objanimdata(i)
                    If sscanf(lineToProcess.lineContents, "%d %d %d %d %d %d %d %f %f %f%n", _
                              frame, pos.x, pos.y, pos.z, rot.x, rot.y, rot.z, _
                              .scale.x, .scale.y, .scale.z, cnt) <> 10 Then
                        debug_ 1657, LOG_ERROR, "_imd_load_level", "%s: Invalid object animation level %u, line %d, frame %d", filename, level, i, frame
                    End If
                    If Not (frame = i) Then ASSERT_FAILURE "frame == i", "lib/ivis_opengl/imdload.cpp:1659", "_imd_load_level", "%s: Invalid frame enumeration object animation (level %u) %d: %d", filename, level, i, frame
                    .pos.x = fdiv(CSng(pos.x), CSng(INT_SCALE))
                    .pos.y = fdiv(CSng(pos.z), CSng(INT_SCALE))
                    .pos.z = fdiv(CSng(pos.y), CSng(INT_SCALE))
                    .rot.pitch = TruncF(-fdiv(fmul(CSng(rot.x), DEG_1), CSng(INT_SCALE))) And &HFFFF&
                    .rot.direction = TruncF(-fdiv(fmul(CSng(rot.z), DEG_1), CSng(INT_SCALE))) And &HFFFF&
                    .rot.roll = TruncF(-fdiv(fmul(CSng(rot.y), DEG_1), CSng(INT_SCALE))) And &HFFFF&
                End With
                pFileData = lineToProcess.pNextLineBegin
            Next
        ElseIf buffer = "SHADOWPOINTS" Then
            If Not (value <> 0) Then ASSERT_FAILURE "value > 0", "lib/ivis_opengl/imdload.cpp:1671", "_imd_load_level", "Invalid 'SHADOW_POINTS' count, got: %u", value: GoTo fail
            nShadowPoints = value

            tmpShadowShape = blankShape
            If Not imd_load_points(fd, lineToProcess.pNextLineBegin, tmpShadowShape, nShadowPoints) Then
                debug_ 1677, LOG_ERROR, "_imd_load_level", "_imd_load_level(7): file corrupt - invalid shadowpoints: %s", filename
                GoTo fail
            End If

            imdShapes(pAllocatedShape).altShadowPoints_size = tmpShadowShape.points_size
            If tmpShadowShape.points_size > 0 Then
                imdShapes(pAllocatedShape).altShadowPoints = tmpShadowShape.points
            Else
                Erase imdShapes(pAllocatedShape).altShadowPoints
            End If
        ElseIf buffer = "SHADOWPOLYGONS" Then
            If Not (imdShapes(pAllocatedShape).altShadowPoints_size > 0) Then ASSERT_FAILURE "s.altShadowPoints.size() > 0", "lib/ivis_opengl/imdload.cpp:1685", "_imd_load_level", "'SHADOW_POLYGONS' must follow a non-empty SHADOW_POINTS section": GoTo fail
            If Not (value <> 0) Then ASSERT_FAILURE "value > 0", "lib/ivis_opengl/imdload.cpp:1686", "_imd_load_level", "Invalid 'SHADOW_POLYGONS' count, got: %u", value: GoTo fail
            nShadowPolys = value

            tmpShadowShape = blankShape
            tmpShadowShape.points_size = imdShapes(pAllocatedShape).altShadowPoints_size
            tmpShadowShape.points = imdShapes(pAllocatedShape).altShadowPoints
            Erase imdShapes(pAllocatedShape).altShadowPoints   ' std::move
            imdShapes(pAllocatedShape).altShadowPoints_size = 0
            tmpShadowShape.polys_size = nShadowPolys
            ReDim tmpShadowShape.polys(0 To nShadowPolys - 1)
            If Not imd_load_polys(filename, fd, lineToProcess.pNextLineBegin, tmpShadowShape, PIE_FLOAT_VER, tmpShadowShape.points_size, True) Then
                debug_ 1694, LOG_ERROR, "_imd_load_level", "_imd_load_level(8): file corrupt - invalid shadowpolygons: %s", filename
                GoTo fail
            End If

            imdShapes(pAllocatedShape).altShadowPoints_size = tmpShadowShape.points_size
            imdShapes(pAllocatedShape).altShadowPoints = tmpShadowShape.points
            imdShapes(pAllocatedShape).altShadowPolys_size = tmpShadowShape.polys_size
            imdShapes(pAllocatedShape).altShadowPolys = tmpShadowShape.polys
        Else
            debug_ 1703, LOG_ERROR, "_imd_load_level", "(_load_level) unexpected directive %s %u", buffer, value
            Exit Do
        End If
    Loop

    ' Sanity checks
    If pie_level_normals_size <> 0 Then
        If imdShapes(pAllocatedShape).polys_size * 3 <> pie_level_normals_size Then
            debug_ 1713, LOG_ERROR, "_imd_load_level", "imd[_load_level] = got %u npolys, but there are only %zu normals! Discarding normals...", _
                npolys, pie_level_normals_size
            ' This will force usage of calculated normals in addVertex
            Erase pie_level_normals
            pie_level_normals_size = 0
        End If
    End If
    If imdShapes(pAllocatedShape).altShadowPoints_size <> 0 Then
        If imdShapes(pAllocatedShape).altShadowPolys_size <> 0 Then
            imdShapes(pAllocatedShape).shadowUsesAlt = True
        Else
            debug_ 1728, LOG_ERROR, "_imd_load_level", "imd[_load_level] = %zu SHADOWPOINTS specified, but there are no SHADOWPOLYGONS! Discarding shadow points...", _
                imdShapes(pAllocatedShape).altShadowPoints_size
            Erase imdShapes(pAllocatedShape).altShadowPoints
            imdShapes(pAllocatedShape).altShadowPoints_size = 0
            Erase imdShapes(pAllocatedShape).altShadowPolys
            imdShapes(pAllocatedShape).altShadowPolys_size = 0
        End If
    Else
        Erase imdShapes(pAllocatedShape).altShadowPoints
        imdShapes(pAllocatedShape).altShadowPoints_size = 0
        Erase imdShapes(pAllocatedShape).altShadowPolys
        imdShapes(pAllocatedShape).altShadowPolys_size = 0
    End If

    ' FINALLY, massage the data into what can stream directly to GPU buffers
    If Not skipGPUData Then
        ReDim vertices(0 To 63): vertices_size = 0
        ReDim normals(0 To 63): normals_size = 0
        ReDim texcoords(0 To 63): texcoords_size = 0
        ReDim indices(0 To imdShapes(pAllocatedShape).polys_size * 3 + 2): indices_size = 0
        vertexCount = 0
        ' Go through all polygons for each frame
        For npol = 0 To imdShapes(pAllocatedShape).polys_size - 1
            ' Do we already have the vertex data for this polygon?
            For i = 0 To 2
                indices(indices_size) = addVertex(imdShapes(pAllocatedShape), i, imdShapes(pAllocatedShape).polys(npol), npol, pie_level_normals, pie_level_normals_size)
                indices_size = indices_size + 1
            Next
        Next

        imdShapes(pAllocatedShape).vertexCount = vertexCount

        ' Tangents are optional, only if normals were loaded and passed sanity check above
        If pie_level_normals_size <> 0 Then
            tangents_size = vertexCount * 4
            bitangents_size = vertexCount * 3
            If vertexCount > 0 Then
                ReDim tangents(0 To tangents_size - 1)
                ReDim bitangents(0 To bitangents_size - 1)
            End If

            For i = 0 To indices_size - 1 Step 3
                calculateTangentsForTriangle indices(i), indices(i + 1), indices(i + 2)
            Next
            finishTangentsGeneration

            If tangents_size <> 0 Then
                UploadF imdShapes(pAllocatedShape).vbo_tangents, imdShapes(pAllocatedShape).vbo_tangents_size, tangents, tangents_size
                uploadShapeBufferF pAllocatedShape, VBO_TANGENT, buffer_usage_vertex_buffer, "tangent buffer", tangents, tangents_size
            End If
        End If

        If vertices_size = 0 Then
            debug_ 1778, LOG_ERROR, "_imd_load_level", "_imd_load_level: file corrupt? - no vertices?: %s (key: %s)", filename, key
        End If
        UploadF imdShapes(pAllocatedShape).vbo_vertices, imdShapes(pAllocatedShape).vbo_vertices_size, vertices, vertices_size
        uploadShapeBufferF pAllocatedShape, VBO_VERTEX, buffer_usage_vertex_buffer, "vertex buffer", vertices, vertices_size

        If normals_size = 0 Then
            debug_ 1786, LOG_ERROR, "_imd_load_level", "_imd_load_level: file corrupt? - no normals?: %s (key: %s)", filename, key
        End If
        UploadF imdShapes(pAllocatedShape).vbo_normals, imdShapes(pAllocatedShape).vbo_normals_size, normals, normals_size
        uploadShapeBufferF pAllocatedShape, VBO_NORMAL, buffer_usage_vertex_buffer, "normals buffer", normals, normals_size

        If indices_size = 0 Then
            debug_ 1794, LOG_ERROR, "_imd_load_level", "_imd_load_level: file corrupt? - no indices?: %s (key: %s)", filename, key
        End If
        imdShapes(pAllocatedShape).vbo_indices_size = indices_size
        If indices_size > 0 Then
            imdShapes(pAllocatedShape).vbo_indices = indices
            ReDim Preserve imdShapes(pAllocatedShape).vbo_indices(0 To indices_size - 1)
        End If
        uploadShapeIndexBuffer pAllocatedShape, indices, indices_size

        If texcoords_size = 0 Then
            debug_ 1802, LOG_ERROR, "_imd_load_level", "_imd_load_level: file corrupt? - no texcoords?: %s (key: %s)", filename, key
        End If
        UploadF imdShapes(pAllocatedShape).vbo_texcoords, imdShapes(pAllocatedShape).vbo_texcoords_size, texcoords, texcoords_size
        uploadShapeBufferF pAllocatedShape, VBO_TEXCOORD, buffer_usage_vertex_buffer, "tex coords buffer", texcoords, texcoords_size
    End If

    Erase indices: indices_size = 0
    Erase vertices: vertices_size = 0
    Erase texcoords: texcoords_size = 0
    Erase normals: normals_size = 0
    Erase tangents: tangents_size = 0
    Erase bitangents: bitangents_size = 0

    ppFileData = pFileData

    ' decide which flags to use (local level override or global)
    If levelSettings_.has_imd_flags Then
        flags = levelSettings_.imd_flags
    ElseIf globalLevelSettings.has_imd_flags Then
        flags = globalLevelSettings.imd_flags
    Else
        flags = 0
    End If
    imdShapes(pAllocatedShape).flags = flags

    ' default for interpolation is on
    If levelSettings_.has_interpolate Then
        imdShapes(pAllocatedShape).interpolate = -CLng(levelSettings_.interpolate)
    ElseIf globalLevelSettings.has_interpolate Then
        imdShapes(pAllocatedShape).interpolate = -CLng(globalLevelSettings.interpolate)
    Else
        imdShapes(pAllocatedShape).interpolate = 1
    End If

    imd_determine_tileset_texture_files imdShapes(pAllocatedShape), globalLevelSettings, levelSettings_

    imd_load_level = pAllocatedShape
    Exit Function
fail:
    deleteShape pAllocatedShape
    imd_load_level = 0
End Function

' Load ppFileData into a shape
' \param ppFileData Data from the IMD file
' \param FileDataEnd Endpointer
' \return The shape, constructed from the data read
' ppFileData is incremented to the end of the file on exit!
Private Function iV_ProcessIMD(ByVal filename As String, fd As IMD_FILE, ppFileData As Long, ByVal skipGPUData As Boolean, Optional ByVal skipDuplicateLoadChecks As Boolean = False) As Long
    Dim pFileData As Long, buffer As String, cnt As Long, value As Long, nlevels As Long, imd_version As Long
    Dim objanimpie(0 To ANIM_EVENT_COUNT - 1) As Long, lineToProcess As IMD_Line, globalLevelSettings As LevelSettings
    Dim i As Long, animpie As String, pRestOfLine As String, firstLevel As Long, lastLevel As Long, level As Long, shape As Long
    Dim scanResult As Long

    pFileData = ppFileData
    buffer = ""

    If Not imd_get_next_line(fd, pFileData, lineToProcess) Then
        GoTo bad_version
    ElseIf sscanf(lineToProcess.lineContents, "%255s %d", buffer, imd_version) <> 2 Then
        GoTo bad_version
    End If
    pFileData = lineToProcess.pNextLineBegin

    If PIE_NAME <> buffer Then
        debug_ 1857, LOG_ERROR, "iV_ProcessIMD", "%s: Not an IMD file (%s %d)", filename, buffer, imd_version
        modelLoadingErrors = modelLoadingErrors + 1
        iV_ProcessIMD = 0
        Exit Function
    End If

    'Now supporting version PIE_VER and PIE_FLOAT_VER files
    If imd_version < PIE_MIN_VER Or imd_version > PIE_MAX_VER Then
        debug_ 1865, LOG_ERROR, "iV_ProcessIMD", "%s: Version %d not supported", filename, imd_version
        modelLoadingErrors = modelLoadingErrors + 1
        iV_ProcessIMD = 0
        Exit Function
    End If

    If Not imd_load_level_settings(filename, fd, pFileData, imd_version, True, globalLevelSettings) Then
        debug_ 1873, LOG_ERROR, "iV_ProcessIMD", "%s: Failed to load level settings", filename
        modelLoadingErrors = modelLoadingErrors + 1
        iV_ProcessIMD = 0
        Exit Function
    End If
    ' TYPE is required in the global scope
    If Not (globalLevelSettings.has_imd_flags) Then ASSERT_FAILURE "globalLevelSettings.imd_flags.has_value()", "lib/ivis_opengl/imdload.cpp:1878", "iV_ProcessIMD", "%s: Missing TYPE line", filename: iV_ProcessIMD = 0: Exit Function

    lineToProcess.pNextLineBegin = pFileData

    ' Next: either EVENT or LEVELS
    If Not getNextPossibleCommandLine(fd, pFileData, lineToProcess, buffer, value, cnt, scanResult) Then
        debug_ 1900, LOG_ERROR, "iV_ProcessIMD", "%s: Expecting EVENT or LEVELS: %s", filename, buffer
        modelLoadingErrors = modelLoadingErrors + 1
        iV_ProcessIMD = 0
        Exit Function
    End If

    For i = 0 To ANIM_EVENT_COUNT - 1
        objanimpie(i) = 0
    Next
    Do While Left$(buffer, 5) = "EVENT"
        If Not (ULT(value, ANIM_EVENT_COUNT)) Then ASSERT_FAILURE "value < ANIM_EVENT_COUNT", "lib/ivis_opengl/imdload.cpp:1913", "iV_ProcessIMD", "Invalid event type %u", value
        pRestOfLine = Mid$(lineToProcess.lineContents, cnt + 1)
        If sscanf(pRestOfLine, "%255s%n", animpie, cnt) <> 1 Then
            debug_ 1917, LOG_ERROR, "iV_ProcessIMD", "%s animation model corrupt: %s", filename, buffer
            modelLoadingErrors = modelLoadingErrors + 1
            iV_ProcessIMD = 0
            Exit Function
        End If

        objanimpie(value) = modelGet(animpie)

        ' Try -yet again- to read in LEVELS directive
        If Not getNextPossibleCommandLine(fd, pFileData, lineToProcess, buffer, value, cnt, scanResult) Then
            debug_ 1927, LOG_ERROR, "iV_ProcessIMD", "%s: Bad levels info: %s", filename, buffer
            modelLoadingErrors = modelLoadingErrors + 1
            iV_ProcessIMD = 0
            Exit Function
        End If
    Loop

    If Left$(buffer, 6) <> "LEVELS" Then
        debug_ 1935, LOG_ERROR, "iV_ProcessIMD", "%s: Expecting 'LEVELS' directive (%s)", filename, buffer
        modelLoadingErrors = modelLoadingErrors + 1
        iV_ProcessIMD = 0
        Exit Function
    End If
    nlevels = value

    firstLevel = 0
    lastLevel = 0
    For level = 0 To nlevels - 1
        ' Read LEVEL directive
        If Not getNextPossibleCommandLine(fd, pFileData, lineToProcess, buffer, value, cnt, scanResult) Then
            debug_ 1948, LOG_ERROR, "iV_ProcessIMD", "(_load_level) file corrupt -J"
            GoTo fail
        End If
        If Left$(buffer, 5) <> "LEVEL" Then
            debug_ 1954, LOG_ERROR, "iV_ProcessIMD", "%s: Expecting 'LEVEL' directive (%s)", filename, buffer
            GoTo fail
        End If
        If value <> (level + 1) Then
            debug_ 1960, LOG_ERROR, "iV_ProcessIMD", "%s: LEVEL %u is invalid - expecting LEVEL %u (LEVELS must be sequential, starting at 1)", filename, value, level
            GoTo fail
        End If

        shape = imd_load_level(filename, fd, lineToProcess.pNextLineBegin, imd_version, level, globalLevelSettings, skipGPUData, skipDuplicateLoadChecks)
        If shape = 0 Then
            debug_ 1968, LOG_ERROR, "iV_ProcessIMD", "%s: Unsuccessful loading level %u", filename, (level + 1)
            GoTo fail
        End If

        If lastLevel <> 0 Then
            imdShapes(lastLevel).nxt = shape
            lastLevel = shape
        Else
            If Not (level = 0) Then ASSERT_FAILURE "level == 0", "lib/ivis_opengl/imdload.cpp:1980", "iV_ProcessIMD", "Invalid level"
            firstLevel = shape
            lastLevel = firstLevel
        End If

        pFileData = lineToProcess.pNextLineBegin
    Next

    If Not (firstLevel <> 0) Then ASSERT_FAILURE "firstLevel != nullptr", "lib/ivis_opengl/imdload.cpp:1988", "iV_ProcessIMD", "%s: Has no levels?", filename: iV_ProcessIMD = 0: Exit Function

    ' copy over model-wide animation information, stored only in the first level
    For i = 0 To ANIM_EVENT_COUNT - 1
        imdShapes(firstLevel).objanimpie(i) = objanimpie(i)
    Next

    ' TODO: once all levels have been loaded, re-calculate the bounds of the first level using all the levels' points?

    ppFileData = pFileData
    iV_ProcessIMD = firstLevel
    Exit Function
bad_version:
    debug_ 1848, LOG_ERROR, "iV_ProcessIMD", "%s: bad PIE version: (%s)", filename, buffer
    modelLoadingErrors = modelLoadingErrors + 1
    ' (assert(false): a release build has none)
    iV_ProcessIMD = 0
    Exit Function
fail:
    modelLoadingErrors = modelLoadingErrors + 1
    deleteShape firstLevel                  ' (std::unique_ptr firstLevel and its next levels)
    iV_ProcessIMD = 0
End Function

'--------------------------------------------------------------------------
' (VB) modelDump: a model as text, for tools\run_tests.py, which compares it
' with tools\statsref's (the same lines from the C++ structures): the base
' shape, then each display level; a float is its bits in hex.

Private Sub dumpPut(ByVal s As String)
    If dumpCount > UBound(dumpChunks) Then ReDim Preserve dumpChunks(0 To dumpCount * 2)
    dumpChunks(dumpCount) = s
    dumpCount = dumpCount + 1
End Sub

Private Function v3iStr(v As Vector3i) As String
    v3iStr = v.x & " " & v.y & " " & v.z
End Function

Private Function v3fStr(v As Vector3f) As String
    v3fStr = FloatHex(v.x) & " " & FloatHex(v.y) & " " & FloatHex(v.z)
End Function

Private Sub dumpConnectors(c() As Vector3i, ByVal n As Long)
    Dim i As Long
    dumpPut " connectors " & n & ":"
    For i = 0 To n - 1
        dumpPut " " & v3iStr(c(i)) & ";"
    Next
    dumpPut vbLf
End Sub

Private Sub dumpPoints(ByVal label As String, p() As Vector3f, ByVal n As Long)
    Dim i As Long
    dumpPut " " & label & " " & n & ":"
    For i = 0 To n - 1
        dumpPut " " & v3fStr(p(i)) & ";"
    Next
    dumpPut vbLf
End Sub

Private Sub dumpPolys(ByVal label As String, p() As iIMDPoly, ByVal n As Long)
    Dim i As Long
    dumpPut " " & label & " " & n & vbLf
    For i = 0 To n - 1
        With p(i)
            dumpPut "  " & i & ": flags=" & Hex$(.flags) & " p=" & U32(.pindex(0)) & " " & U32(.pindex(1)) & " " & U32(.pindex(2)) _
                & " tc=" & FloatHex(.texCoord(0).x) & " " & FloatHex(.texCoord(0).y) & " " & FloatHex(.texCoord(1).x) & " " & FloatHex(.texCoord(1).y) _
                & " " & FloatHex(.texCoord(2).x) & " " & FloatHex(.texCoord(2).y) _
                & " anim=" & FloatHex(.texAnim.x) & " " & FloatHex(.texAnim.y) & " n=" & v3fStr(.normal) & vbLf
        End With
    Next
End Sub

Private Sub dumpFloats(ByVal label As String, v() As Single, ByVal n As Long)
    Dim i As Long
    dumpPut " " & label & " " & n & ":"
    For i = 0 To n - 1
        dumpPut " " & FloatHex(v(i))
    Next
    dumpPut vbLf
End Sub

Private Function baseName(ByVal baseShape As Long) As String
    If baseShape = 0 Then baseName = "null" Else baseName = imdBaseShapes(baseShape).filename
End Function

Public Function modelDump(ByVal baseShape As Long) As String
    Dim s As Long, k As Long, i As Long, t As String
    ReDim dumpChunks(0 To 255)
    dumpCount = 0
    With imdBaseShapes(baseShape)
        dumpPut "base filename=" & .filename & " path=" & .path & " min=" & v3iStr(.min) & " max=" & v3iStr(.max) _
            & " sradius=" & .sradius & " radius=" & .radius & " ocen=" & v3fStr(.ocen) & vbLf
        dumpConnectors .connectors, .connectors_size
        s = .m_displayModel
    End With
    k = 0
    Do While s <> 0
        With imdShapes(s)
            dumpPut "level " & k & " name=" & .modelName & " modelLevel=" & U32(.modelLevel) & " flags=" & Hex$(.flags) _
                & " numFrames=" & .numFrames & " animInterval=" & .animInterval & " interpolate=" & .interpolate _
                & " shadowUsesAlt=" & -CLng(.shadowUsesAlt) & " vertexCount=" & .vertexCount & vbLf
            dumpPut " bounds min=" & v3iStr(.min) & " max=" & v3iStr(.max) & " sradius=" & .sradius & " radius=" & .radius & " ocen=" & v3fStr(.ocen) & vbLf
            dumpConnectors .connectors, .connectors_size
            t = " textures"
            For i = 0 To 2
                t = t & " " & .tilesetTextureFiles(i).texfile & "|" & .tilesetTextureFiles(i).tcmaskfile & "|" & .tilesetTextureFiles(i).normalfile & "|" & .tilesetTextureFiles(i).specfile
            Next
            dumpPut t & vbLf
            dumpPut " objanim time=" & .objanimtime & " cycles=" & .objanimcycles & " frames=" & .objanimframes & " pie=" _
                & baseName(.objanimpie(0)) & " " & baseName(.objanimpie(1)) & " " & baseName(.objanimpie(2)) & " " & baseName(.objanimpie(3)) & vbLf
            For i = 0 To .objanimframes - 1
                dumpPut "  " & i & ": scale=" & v3fStr(.objanimdata(i).scale) & " pos=" & v3fStr(.objanimdata(i).pos) _
                    & " rot=" & .objanimdata(i).rot.direction & " " & .objanimdata(i).rot.pitch & " " & .objanimdata(i).rot.roll & vbLf
            Next
            dumpPoints "points", .points, .points_size
            dumpPolys "polys", .polys, .polys_size
            dumpPoints "altpoints", .altShadowPoints, .altShadowPoints_size
            dumpPolys "altpolys", .altShadowPolys, .altShadowPolys_size
            dumpFloats "vertices", .vbo_vertices, .vbo_vertices_size
            dumpFloats "normals", .vbo_normals, .vbo_normals_size
            dumpFloats "texcoords", .vbo_texcoords, .vbo_texcoords_size
            dumpFloats "tangents", .vbo_tangents, .vbo_tangents_size
            t = " indices " & .vbo_indices_size & ":"
            dumpPut t
            For i = 0 To .vbo_indices_size - 1
                dumpPut " " & .vbo_indices(i)
            Next
            dumpPut vbLf
            s = .nxt
        End With
        k = k + 1
    Loop
    ReDim Preserve dumpChunks(0 To dumpCount - 1)
    modelDump = Join(dumpChunks, "")
    Erase dumpChunks
End Function
