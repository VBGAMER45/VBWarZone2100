Attribute VB_Name = "ivisdef_h"
' lib/ivis_opengl/ivisdef.h: the model structures (with glm's vector types
' as UDTs, lib/framework/vector.h's Rotation, pietypes.h's
' OLD_TEXTURE_SIZE_FIX and tex.h's iV_TEX_INVALID). An iIMDShape * or
' iIMDBaseShape * is a 1-based handle into imdload_c's pools (0 =
' nullptr); a std::unique_ptr's owner frees it. iIMDShape's
' pShadowPoints / pShadowPolys (pointers to either the points and polys or
' their alternative shadow copies) are shadowUsesAlt. The GPU buffers
' (buffers, gfx_api_buffer objects: Nothing = nullptr; they go with the
' shape) are made by _imd_load_level from the vertex data, which also stays
' in the shape (vbo_*) for the tests' model dumps. iClip and iSurface are
' piemode_c's.
Option Explicit

Public Const OLD_TEXTURE_SIZE_FIX As Single = 256!   ' pietypes.h
Public Const iV_TEX_INVALID As Long = 0             ' tex.h

Public Const WZ_CURRENT_GRAPHICS_OVERRIDES_PREFIX As String = "graphics_overrides/curr"

' lib/framework/vector.h (glm's vectors are src/shared/gfx's glm_h)
Public Type Rotation
    direction As Long                   ' uint16_t: Object rotation in 0..64k range
    pitch As Long
    roll As Long
End Type

' Stores the from and to verticles from an edge
Public Type EDGE
    from As Long                        ' uint32_t
    to_ As Long                         ' uint32_t
    sort_key As Double                  ' uint64_t (from << 32 | to; exact below 2^53)
End Type

Public Type ANIMFRAME
    scale As Vector3f
    pos As Vector3f
    rot As Rotation
End Type

Public Type iIMDPoly
    texCoord(0 To 2) As Vector2f
    texAnim As Vector2f
    flags As Long                       ' uint32_t
    zcentre As Long
    normal As Vector3f
    pindex(0 To 2) As Long              ' uint32_t
End Type

Public Enum VBO_TYPE
    VBO_VERTEX
    VBO_TEXCOORD
    VBO_MINIMAL
    VBO_NORMAL = VBO_MINIMAL
    VBO_INDEX
    VBO_TANGENT
    VBO_COUNT
End Enum

Public Enum ANIMATION_EVENTS
    ANIM_EVENT_NONE
    ANIM_EVENT_ACTIVE
    ANIM_EVENT_FIRING ' should not be combined with fire-on-move, as this will look weird
    ANIM_EVENT_DYING
    ANIM_EVENT_COUNT
End Enum

' A game model will have (potentially) two sets of information:
' 1. Information used for game state calculations (this is always from the "base" / core model file)
'		- This is accessible in the `iIMDBaseShape`
' 2. Information used for display (this may be from the "base" / core model, or a graphics mod overlay display-only model replacement)
'		- This is accessible in the `iIMDShape` that can be obtained via `iIMDBaseShape::displayModel()`
Public Type iIMDBaseShape
    ' SAFE FOR USE IN GAME STATE CALCULATIONS

    min As Vector3i                     ' used by: establishTargetHeight (game state calculation), alignStructure (game state calculation), etc
    max As Vector3i
    sradius As Long
    radius As Long

    ocen As Vector3f

    connectors() As Vector3i            ' used by: muzzle base location, fire line, actionVisibleTarget, etc)
    connectors_size As Long

    path As String
    filename As String

    m_displayModel As Long              ' std::unique_ptr<iIMDShape>: the display shape used for rendering (*NOT* for any game state calculations!)
End Type

Public Type iIMDShapeTextures
    initialized As Boolean
    texpage As Long                     ' size_t
    tcmaskpage As Long
    normalpage As Long
    specularpage As Long
End Type

Public Type TilesetTextureFiles
    texfile As String
    tcmaskfile As String
    normalfile As String
    specfile As String
End Type

' DISPLAY-ONLY
' NOTE: Do *NOT* use any data from iIMDShape in game state calculations - instead, use the data in an iIMDBaseShape
Public Type iIMDShape
    min As Vector3i
    max As Vector3i
    sradius As Long
    radius As Long

    ocen As Vector3f

    connectors() As Vector3i
    connectors_size As Long

    flags As Long                       ' unsigned int

    numFrames As Long                   ' unsigned short
    animInterval As Long                ' unsigned short

    shadowEdgeList() As EDGE            ' EDGE *shadowEdgeList = nullptr (nShadowEdges of them)
    nShadowEdges As Long                ' size_t

    ' The old rendering data
    points() As Vector3f                ' NOTE: This is used to calculate some of the values above on imd load (in _imd_calc_bounds)
    points_size As Long
    polys() As iIMDPoly
    polys_size As Long

    ' Data used for stencil shadows
    altShadowPoints() As Vector3f
    altShadowPoints_size As Long
    altShadowPolys() As iIMDPoly
    altShadowPolys_size As Long
    shadowUsesAlt As Boolean            ' (VB) pShadowPoints / pShadowPolys point to the alternative ones

    ' The new rendering data
    buffers(0 To VBO_COUNT - 1) As gfx_api_buffer ' std::array<gfx_api::buffer*, VBO_COUNT> buffers = { };
    vertexCount As Long                 ' uint16_t

    ' (VB) the data the buffers get (C++: uploaded while loading): gfxFloat and uint16_t arrays
    vbo_vertices() As Single
    vbo_vertices_size As Long
    vbo_normals() As Single
    vbo_normals_size As Long
    vbo_texcoords() As Single
    vbo_texcoords_size As Long
    vbo_tangents() As Single
    vbo_tangents_size As Long
    vbo_indices() As Long
    vbo_indices_size As Long

    ' object animation (animating a level, rather than its texture)
    objanimdata() As ANIMFRAME          ' (objanimframes of them)
    objanimframes As Long

    ' more object animation, but these are only set for the first level
    objanimtime As Long                 ' total time to render all animation frames
    objanimcycles As Long               ' Number of cycles to render, zero means infinitely many
    objanimpie(0 To ANIM_EVENT_COUNT - 1) As Long ' non-owned pointer to loaded base shape

    interpolate As Long                 ' if the model wants to be interpolated

    modelName As String
    modelLevel As Long                  ' uint32_t

    tilesetTextureFiles(0 To 2) As TilesetTextureFiles

    nxt As Long                         ' std::unique_ptr<iIMDShape> next: next pie in multilevel pies (NULL for non multilevel !)

    m_textures As iIMDShapeTextures
End Type

' ivisdef.h: the image sheets (bitimage_c). IMAGEFILE * is a handle, the 1-based index of
' bitimage_c's imageFiles (0 = nullptr); AtlasImageDef * is file * &H10000 + index + 1 (0 = nullptr)
Public Type AtlasImageDef
    TPageID As Long                     ' size_t: Which associated file to read our info from
    Tu As Long                          ' unsigned int: First vertex coordinate
    Tv As Long                          ' unsigned int: Second vertex coordinate
    Width As Long                       ' unsigned int: Width of image
    Height As Long                      ' unsigned int: Height of image
    XOffset As Long                     ' X offset into source position
    YOffset As Long                     ' Y offset into source position

    textureId As Long                   ' size_t: duplicate of below, fix later
    invTextureSize As Single            ' gfx_api::gfxFloat
End Type

' IMAGEFILE::Page
Public Type IMAGEFILE_Page
    id As Long                          ' size_t: OpenGL texture ID. (a tex_c page)
    size As Long                        ' Size of texture in pixels. (Should be square.)
End Type

Public Type IMAGEFILE
    pages() As IMAGEFILE_Page           ' Texture pages.
    pages_size As Long
    imageDefs() As AtlasImageDef        ' Stored images.
    imageDefs_size As Long
    imageNamesMap As Object             ' std::unordered_map<WzString, AtlasImageDef *>: a Scripting.Dictionary, name -> index in imageDefs
    inUse As Boolean                    ' (VB) the pool slot holds a file
End Type
