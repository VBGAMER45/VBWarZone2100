Attribute VB_Name = "pieblitfunc_c"
' lib/ivis_opengl/pieblitfunc.cpp/.h: 2D drawing (lines, boxes, images of
' the image sheets, the radar, the backdrop's choice); pietypes.h's
' PIERECT_DrawRequest(_f). The class GFX is GFX.cls (made by GFX_new), the
' batched renderers BatchedMultiRectRenderer.cls and
' BatchedImageDrawRequests.cls (made by BatchedImageDrawRequests_new).
' Templates over the PSO take the pso_type (CONVENTIONS: the type as a
' parameter). IMAGEFILE * is bitimage_c's handle (imageFiles(h)), an
' AtlasImageDef * its packed handle (file * &H10000 + index + 1), a texture
' page tex_c's index (pie_Texture). The PIELIGHTs are Longs (piestate_c).
' The functions the exe calls with a const glm::mat4& take the mat4 here;
' GfxApi hands them the exe's address (0 = the C++ default argument,
' defaultProjectionMatrix(): defaultProjectionMatrix_or). C++'s vector and
' pointer-to-array arguments the exe passes are addresses in the _p forms.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Const NUM_BACKDROPS As Long = 7
Public Const BACKDROP_HACK_WIDTH As Long = 512
Public Const BACKDROP_HACK_HEIGHT As Long = 512

' pietypes.h
Public Type PIERECT_DrawRequest
    x0 As Long
    y0 As Long
    x1 As Long
    y1 As Long
    color As Long                       ' PIELIGHT
End Type

Public Type PIERECT_DrawRequest_f
    x0 As Single
    y0 As Single
    x1 As Single
    y1 As Single
    color As Long                       ' PIELIGHT
End Type

' Screen rectangle.
Public Type PIERECT
    x As Single
    y As Single
    w As Single
    h As Single
End Type

Public Type PieDrawImageRequest
    rendMode As REND_MODE
    imageFile_ As Long                   ' IMAGEFILE *
    ID As Long
    size As Vector2i
    dest As PIERECT
    colour As Long                      ' PIELIGHT
    modelViewProjection As mat4
    textureInset As Vector2i            ' = Vector2i(0, 0)
End Type

'***************************************************************************
'
'	Local Variables
'
'***************************************************************************

Private Const NUM_RADAR_TEXTURES As Long = 2
Private radarGfx(0 To NUM_RADAR_TEXTURES - 1) As GFX   ' = {nullptr}
Private currRadarGfx As Long            ' = 0

' iV_DrawImageRepeatX's and iV_DrawImageRepeatY's static BatchedImageDrawRequests localBatch;
Private repeatXLocalBatch As BatchedImageDrawRequests
Private repeatYLocalBatch As BatchedImageDrawRequests

'***************************************************************************
'
'	Source
'
'***************************************************************************

' (VB) new GFX(type, coordsPerVertex)
Public Function GFX_new(ByVal type_ As GFXTYPE, Optional ByVal coordsPerVertex As Long = 3) As GFX
    Set GFX_new = New GFX
    GFX_new.init type_, coordsPerVertex
End Function

' (VB) BatchedImageDrawRequests(bool deferRender = false)
Public Function BatchedImageDrawRequests_new(Optional ByVal deferRender As Boolean = False) As BatchedImageDrawRequests
    Set BatchedImageDrawRequests_new = New BatchedImageDrawRequests
    BatchedImageDrawRequests_new.deferRender = deferRender
End Function

' (VB) the mat4 at address p, or defaultProjectionMatrix() for 0 (the C++ default argument)
Public Function defaultProjectionMatrix_or(ByVal p As Long) As mat4
    If p = 0 Then
        defaultProjectionMatrix_or = defaultProjectionMatrix()
    Else
        defaultProjectionMatrix_or = mat4_fromPtr(p)
    End If
End Function

Public Sub iV_Line(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal colour As Long)
    Dim color As Vector4f, mat As mat4, cb As constant_buffer_type_SHADER_LINE
    color = pielightToRGBAVec4(colour)
    mat = glm_ortho(0!, CSng(pie_GetVideoBufferWidth()), CSng(pie_GetVideoBufferHeight()), 0!)
    gfx_api_pso_get(LinePSO).bind
    cb.mat = mat
    cb.p0 = glm_vec2(x0, y0)
    cb.p1 = glm_vec2(x1, y1)
    cb.colour = color
    gfx_api_pso_get(LinePSO).bind_constants VarPtr(cb), LenB(cb)
    gfx_api_pso_get(LinePSO).bind_vertex_buffers pie_internal_rectBuffer
    gfx_api_pso_get(LinePSO).draw 2, 0
    gfx_api_pso_get(LinePSO).unbind_vertex_buffers pie_internal_rectBuffer
End Sub

' lines: lines(0 .. count - 1) (glm::ivec4: x0, y0, x1, y1)
Public Sub iV_Lines(lines() As Vector4i, ByVal count As Long, ByVal colour As Long)
    Dim color As Vector4f, mat As mat4, cb As constant_buffer_type_SHADER_LINE, i As Long
    color = pielightToRGBAVec4(colour)
    mat = glm_ortho(0!, CSng(pie_GetVideoBufferWidth()), CSng(pie_GetVideoBufferHeight()), 0!)
    gfx_api_pso_get(LinePSO).bind
    gfx_api_pso_get(LinePSO).bind_vertex_buffers pie_internal_rectBuffer
    For i = 0 To count - 1
        cb.mat = mat
        cb.p0 = glm_vec2(lines(i).x, lines(i).y)
        cb.p1 = glm_vec2(lines(i).z, lines(i).w)
        cb.colour = color
        gfx_api_pso_get(LinePSO).bind_constants VarPtr(cb), LenB(cb)
        gfx_api_pso_get(LinePSO).draw 2, 0
    Next
    gfx_api_pso_get(LinePSO).unbind_vertex_buffers pie_internal_rectBuffer
End Sub

' (VB) iV_Lines for the exe: pLines is the address of count Vector4i (glm::ivec4)
Public Sub iV_Lines_p(ByVal pLines As Long, ByVal count As Long, ByVal colour As Long)
    Dim lines() As Vector4i
    If count <= 0 Then
        ReDim lines(0 To 0)
    Else
        ReDim lines(0 To count - 1)
        CopyMemory lines(0), ByVal pLines, count * LenB(lines(0))
    End If
    iV_Lines lines, count, colour
End Sub

'*
' *	Assumes render mode set up externally, draws filled rectangle.
'
' template<typename PSO>: a PSO with SHADER_RECT's constants
Private Sub pie_DrawRect(ByVal pso As pso_type, ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal colour As Long)
    Dim center As Vector2f, mvp As mat4, t As Single, cb As constant_buffer_type_SHADER_RECT
    If x0 > x1 Then
        t = x0: x0 = x1: x1 = t
    End If
    If y0 > y1 Then
        t = y0: y0 = y1: y1 = t
    End If
    center = glm_vec2(x0, y0)
    mvp = glm_mat4_mul(glm_mat4_mul(defaultProjectionMatrix(), glm_translate(glm_vec3(center.x, center.y, 0!))), glm_scale(glm_vec3(x1 - x0, y1 - y0, 1!)))

    gfx_api_pso_get(pso).bind
    cb.transform_matrix = mvp
    cb.colour = pielightToRGBAVec4(colour)
    gfx_api_pso_get(pso).bind_constants VarPtr(cb), LenB(cb)
    gfx_api_pso_get(pso).bind_vertex_buffers pie_internal_rectBuffer
    gfx_api_pso_get(pso).draw 4, 0
    gfx_api_pso_get(pso).unbind_vertex_buffers pie_internal_rectBuffer
End Sub

' rects: rects(0 .. count - 1)
Public Sub pie_DrawMultiRect(rects() As PIERECT_DrawRequest, ByVal count As Long)
    Dim projectionMatrix As mat4, didEnableRect As Boolean, it As Long, t As Long
    Dim colour As Long, center As Vector2f, mvp As mat4, cb As constant_buffer_type_SHADER_RECT
    If count = 0 Then Exit Sub

    projectionMatrix = defaultProjectionMatrix()
    didEnableRect = False
    gfx_api_pso_get(BoxFillPSO).bind

    For it = 0 To count - 1
        If rects(it).x0 > rects(it).x1 Then
            t = rects(it).x0: rects(it).x0 = rects(it).x1: rects(it).x1 = t
        End If
        If rects(it).y0 > rects(it).y1 Then
            t = rects(it).y0: rects(it).y0 = rects(it).y1: rects(it).y1 = t
        End If
        colour = rects(it).color
        center = glm_vec2(rects(it).x0, rects(it).y0)
        mvp = glm_mat4_mul(glm_mat4_mul(projectionMatrix, glm_translate(glm_vec3(center.x, center.y, 0!))), glm_scale(glm_vec3(rects(it).x1 - rects(it).x0, rects(it).y1 - rects(it).y0, 1!)))
        cb.transform_matrix = mvp
        cb.colour = pielightToRGBAVec4(colour)
        gfx_api_pso_get(BoxFillPSO).bind_constants VarPtr(cb), LenB(cb)
        If Not didEnableRect Then
            gfx_api_pso_get(BoxFillPSO).bind_vertex_buffers pie_internal_rectBuffer
            didEnableRect = True
        End If

        gfx_api_pso_get(BoxFillPSO).draw 4, 0
    Next
    gfx_api_pso_get(BoxFillPSO).unbind_vertex_buffers pie_internal_rectBuffer
End Sub

' (VB) pie_DrawMultiRect for the exe: pRects is the address of count PIERECT_DrawRequest (5 Longs each;
' the vector is passed by value: the caller's copy is not changed)
Public Sub pie_DrawMultiRect_p(ByVal pRects As Long, ByVal count As Long)
    Dim rects() As PIERECT_DrawRequest
    If count <= 0 Then Exit Sub
    ReDim rects(0 To count - 1)
    CopyMemory rects(0), ByVal pRects, count * LenB(rects(0))
    pie_DrawMultiRect rects, count
End Sub

Public Sub pie_DrawMultiRect_NonInstanced(instances() As MultiRectPerInstanceInterleavedData, ByVal count As Long)
    Dim projectionMatrix As mat4, didEnableRect As Boolean, it As Long, mvp As mat4
    Dim cb As constant_buffer_type_SHADER_RECT, c As Long
    If count = 0 Then Exit Sub

    projectionMatrix = defaultProjectionMatrix()
    didEnableRect = False
    gfx_api_pso_get(BoxFillPSO).bind

    For it = 0 To count - 1
        mvp = glm_mat4_mul(projectionMatrix, instances(it).TransformationMatrix)
        c = instances(it).colour
        cb.transform_matrix = mvp
        cb.colour = glm_vec4((c And &HFF&) / 255!, ((c And &HFF00&) \ &H100&) / 255!, ((c And &HFF0000) \ &H10000) / 255!, PIELIGHT_a(c) / 255!)
        gfx_api_pso_get(BoxFillPSO).bind_constants VarPtr(cb), LenB(cb)
        If Not didEnableRect Then
            gfx_api_pso_get(BoxFillPSO).bind_vertex_buffers pie_internal_rectBuffer
            didEnableRect = True
        End If

        gfx_api_pso_get(BoxFillPSO).draw 4, 0
    Next
    gfx_api_pso_get(BoxFillPSO).unbind_vertex_buffers pie_internal_rectBuffer
End Sub

Public Sub iV_ShadowBox(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal pad As Long, ByVal first As Long, ByVal second As Long, ByVal fill As Long)
    pie_DrawRect ShadowBox2DPSO, x0 + pad, y0 + pad, x1 - pad, y1 - pad, fill
    iV_Box2 x0, y0, x1, y1, first, second
End Sub

'***************************************************************************

Public Sub iV_Box2(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal first As Long, ByVal second As Long)
    Dim mat As mat4, firstColor As Vector4f, secondColor As Vector4f, cb As constant_buffer_type_SHADER_LINE
    mat = glm_ortho(0!, CSng(pie_GetVideoBufferWidth()), CSng(pie_GetVideoBufferHeight()), 0!)
    firstColor = pielightToRGBAVec4(first)
    gfx_api_pso_get(LinePSO).bind
    gfx_api_pso_get(LinePSO).bind_vertex_buffers pie_internal_rectBuffer
    cb.mat = mat
    cb.p0 = glm_vec2(x0, y1): cb.p1 = glm_vec2(x0, y0): cb.colour = firstColor
    gfx_api_pso_get(LinePSO).bind_constants VarPtr(cb), LenB(cb)
    gfx_api_pso_get(LinePSO).draw 2, 0
    cb.p0 = glm_vec2(x0, y0): cb.p1 = glm_vec2(x1, y0): cb.colour = firstColor
    gfx_api_pso_get(LinePSO).bind_constants VarPtr(cb), LenB(cb)
    gfx_api_pso_get(LinePSO).draw 2, 0

    secondColor = pielightToRGBAVec4(second)
    cb.p0 = glm_vec2(x1, y0): cb.p1 = glm_vec2(x1, y1): cb.colour = secondColor
    gfx_api_pso_get(LinePSO).bind_constants VarPtr(cb), LenB(cb)
    gfx_api_pso_get(LinePSO).draw 2, 0
    cb.p0 = glm_vec2(x0, y1): cb.p1 = glm_vec2(x1, y1): cb.colour = secondColor
    gfx_api_pso_get(LinePSO).bind_constants VarPtr(cb), LenB(cb)
    gfx_api_pso_get(LinePSO).draw 2, 0
    gfx_api_pso_get(LinePSO).unbind_vertex_buffers pie_internal_rectBuffer
End Sub

' pieblitfunc.h: static inline void iV_Box(int x0, int y0, int x1, int y1, PIELIGHT first)
Public Sub iV_Box(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal first As Long)
    iV_Box2 x0, y0, x1, y1, first, first
End Sub

'***************************************************************************

Public Sub pie_BoxFill(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal colour As Long)
    pie_DrawRect BoxFillPSO, x0, y0, x1, y1, colour
End Sub

Public Sub pie_BoxFillf(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal colour As Long)
    pie_DrawRect BoxFillPSO, x0, y0, x1, y1, colour
End Sub

Public Sub pie_BoxFill_alpha(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal colour As Long)
    pie_DrawRect BoxFillAlphaPSO, x0, y0, x1, y1, colour
End Sub

'***************************************************************************

Public Sub iV_TransBoxFill(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single)
    pie_UniTransBoxFill x0, y0, x1, y1, psPalette(95)   ' WZCOL_TRANSPARENT_BOX
End Sub

'***************************************************************************

Public Sub pie_UniTransBoxFill(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal light_ As Long)
    pie_DrawRect UniTransBoxPSO, x0, y0, x1, y1, light_
End Sub

'***************************************************************************

Public Function assertValidImage(ByVal imageFile_ As Long, ByVal id As Long) As Boolean
    If Not (imageFile_ <> 0) Then ASSERT_FAILURE "imageFile != nullptr", "lib/ivis_opengl/pieblitfunc.cpp:533", "assertValidImage", "Null imageFile (id: %u)", id: assertValidImage = False: Exit Function
    If Not (id < imageFiles(imageFile_).imageDefs_size) Then ASSERT_FAILURE "id < imageFile->imageDefs.size()", "lib/ivis_opengl/pieblitfunc.cpp:534", "assertValidImage", "Out of range 1: %u/%d", id, imageFiles(imageFile_).imageDefs_size: assertValidImage = False: Exit Function
    If Not (imageFiles(imageFile_).imageDefs(id).TPageID < imageFiles(imageFile_).pages_size) Then ASSERT_FAILURE "imageFile->imageDefs[id].TPageID < imageFile->pages.size()", "lib/ivis_opengl/pieblitfunc.cpp:535", "assertValidImage", "Out of range 2: %zu", imageFiles(imageFile_).imageDefs(id).TPageID: assertValidImage = False: Exit Function
    assertValidImage = True
End Function

' template<typename PSO>: DrawImagePSO, DrawImageAnisotropicPSO (SHADER_TEXRECT) or DrawImageTextPSO (SHADER_TEXT)
Private Sub iv_DrawImageImpl(ByVal pso As pso_type, ByVal TextureID As gfx_api_texture, offset As Vector2f, size As Vector2f, TextureUV As Vector2f, TextureSize As Vector2f, ByVal colour As Long, modelViewProjection As mat4, Optional ByVal mode As SHADER_MODE = SHADER_TEXRECT)
    Dim transformMat As mat4, cb As constant_buffer_type_SHADER_TEXRECT, cbt As constant_buffer_type_SHADER_TEXT
    transformMat = glm_mat4_mul(glm_mat4_mul(modelViewProjection, glm_translate(glm_vec3(offset.x, offset.y, 0!))), glm_scale(glm_vec3(size.x, size.y, 1!)))

    gfx_api_pso_get(pso).bind
    If pso = DrawImageTextPSO Then
        cbt.transform_matrix = transformMat
        cbt.offset = TextureUV
        cbt.size = TextureSize
        cbt.color = pielightToRGBAVec4(colour)
        cbt.texture = 0
        gfx_api_pso_get(pso).bind_constants VarPtr(cbt), LenB(cbt)
    Else
        cb.transform_matrix = transformMat
        cb.offset = TextureUV
        cb.size = TextureSize
        cb.color = pielightToRGBAVec4(colour)
        cb.texture = 0
        gfx_api_pso_get(pso).bind_constants VarPtr(cb), LenB(cb)
    End If
    gfx_api_pso_get(pso).bind_textures TextureID
    gfx_api_pso_get(pso).bind_vertex_buffers pie_internal_rectBuffer
    gfx_api_pso_get(pso).draw 4, 0
    gfx_api_pso_get(pso).unbind_vertex_buffers pie_internal_rectBuffer
End Sub

' angle: in radians (glm::rotate)
Public Sub iV_DrawImageAnisotropic(ByVal TextureID As gfx_api_texture, Position As Vector2i, offset As Vector2f, size As Vector2f, ByVal angle As Single, ByVal colour As Long)
    Dim mvp As mat4
    mvp = glm_mat4_mul(glm_mat4_mul(defaultProjectionMatrix(), glm_translate(glm_vec3(Position.x, Position.y, 0!))), glm_rotate(angle, glm_vec3(0!, 0!, 1!)))

    iv_DrawImageImpl DrawImageAnisotropicPSO, TextureID, offset, size, glm_vec2(0!, 0!), glm_vec2(1!, 1!), colour, mvp
End Sub

' angle: in degrees
Public Sub iV_DrawImageText(ByVal TextureID As gfx_api_texture, Position As Vector2f, offset As Vector2f, size As Vector2f, ByVal angle As Single, ByVal colour As Long)
    Dim mvp As mat4
    mvp = glm_mat4_mul(glm_mat4_mul(defaultProjectionMatrix(), glm_translate(glm_vec3(Position.x, Position.y, 0!))), glm_rotate(RADIANS(angle), glm_vec3(0!, 0!, 1!)))

    iv_DrawImageImpl DrawImageTextPSO, TextureID, offset, size, glm_vec2(0!, 0!), glm_vec2(1!, 1!), colour, mvp, SHADER_TEXT
End Sub

' clippingRect: (VB) the WzClippingRectF as its x(), y(), width(), height()
Public Sub iV_DrawImageTextClipped(ByVal TextureID As gfx_api_texture, textureSize As Vector2i, Position As Vector2f, offset As Vector2f, size As Vector2f, ByVal angle As Single, ByVal colour As Long, ByVal clippingRect_x As Single, ByVal clippingRect_y As Single, ByVal clippingRect_width As Single, ByVal clippingRect_height As Single)
    Dim mvp As mat4, invTextureSizeX As Single, invTextureSizeY As Single
    Dim tu As Single, tv As Single, su As Single, sv As Single
    mvp = glm_mat4_mul(glm_mat4_mul(defaultProjectionMatrix(), glm_translate(glm_vec3(Position.x, Position.y, 0!))), glm_rotate(RADIANS(angle), glm_vec3(0!, 0!, 1!)))

    invTextureSizeX = fdiv(1!, textureSize.x)
    invTextureSizeY = fdiv(1!, textureSize.y)
    tu = clippingRect_x * invTextureSizeX
    tv = clippingRect_y * invTextureSizeY
    su = clippingRect_width * invTextureSizeX
    sv = clippingRect_height * invTextureSizeY

    iv_DrawImageImpl DrawImageTextPSO, TextureID, offset, size, glm_vec2(tu, tv), glm_vec2(su, sv), colour, mvp, SHADER_TEXT
End Sub

' template<typename PSO>
Private Sub pie_DrawImageTemplate(ByVal pso As pso_type, ByVal imageFile_ As Long, ByVal id As Long, size As Vector2i, dest As PIERECT, ByVal colour As Long, modelViewProjection As mat4, textureInset As Vector2i)
    Dim texPage As Long, invTextureSize As Single, tu As Single, tv As Single, su As Single, sv As Single
    Dim mvp As mat4, image2 As AtlasImageDef
    image2 = imageFiles(imageFile_).imageDefs(id)
    texPage = imageFiles(imageFile_).pages(image2.TPageID).id
    invTextureSize = fdiv(1!, CSng(imageFiles(imageFile_).pages(image2.TPageID).size))
    tu = CSng(image2.Tu + textureInset.x) * invTextureSize
    tv = CSng(image2.Tv + textureInset.y) * invTextureSize
    su = CSng(size.x - (textureInset.x * 2)) * invTextureSize
    sv = CSng(size.y - (textureInset.y * 2)) * invTextureSize

    mvp = glm_mat4_mul(modelViewProjection, glm_translate(glm_vec3(dest.x, dest.y, 0!)))

    iv_DrawImageImpl pso, pie_Texture(texPage), glm_vec2(0!, 0!), glm_vec2(dest.w, dest.h), glm_vec2(tu, tv), glm_vec2(su, sv), colour, mvp
End Sub

Private Sub pie_DrawImage(ByVal imageFile_ As Long, ByVal id As Long, size As Vector2i, dest As PIERECT, ByVal colour As Long, modelViewProjection As mat4, textureInset As Vector2i)
    pie_DrawImageTemplate DrawImagePSO, imageFile_, id, size, dest, colour, modelViewProjection, textureInset
End Sub

' requests: requests(0 .. count - 1) (C++: a std::list)
Public Sub pie_DrawMultipleImages(requests() As PieDrawImageRequest, ByVal count As Long)
    Dim didEnableRect As Boolean, r As Long, image2 As AtlasImageDef, texPage As Long, invTextureSize As Single
    Dim tu As Single, tv As Single, su As Single, sv As Single, mvp As mat4, TextureID As gfx_api_texture
    Dim offset As Vector2f, size As Vector2f, TextureUV As Vector2f, TextureSize As Vector2f, transformMat As mat4
    Dim cb As constant_buffer_type_SHADER_TEXRECT
    If count = 0 Then Exit Sub

    didEnableRect = False
    gfx_api_pso_get(DrawImagePSO).bind

    For r = 0 To count - 1
        ' The following is the equivalent of:
        ' pie_DrawImage(request.imageFile, request.ID, request.size, &request.dest, request.colour, request.modelViewProjection, request.textureInset)
        ' but is tweaked to use custom implementation of iv_DrawImageImpl that does not disable the shader after every glDrawArrays call.

        image2 = imageFiles(requests(r).imageFile_).imageDefs(requests(r).ID)
        texPage = imageFiles(requests(r).imageFile_).pages(image2.TPageID).id
        invTextureSize = fdiv(1!, CSng(imageFiles(requests(r).imageFile_).pages(image2.TPageID).size))
        tu = CSng(image2.Tu + requests(r).textureInset.x) * invTextureSize
        tv = CSng(image2.Tv + requests(r).textureInset.y) * invTextureSize
        su = CSng(requests(r).size.x - (requests(r).textureInset.x * 2)) * invTextureSize
        sv = CSng(requests(r).size.y - (requests(r).textureInset.y * 2)) * invTextureSize

        mvp = glm_mat4_mul(requests(r).modelViewProjection, glm_translate(glm_vec3(requests(r).dest.x, requests(r).dest.y, 0!)))

        Set TextureID = pie_Texture(texPage)
        offset = glm_vec2(0!, 0!)
        size = glm_vec2(requests(r).dest.w, requests(r).dest.h)
        TextureUV = glm_vec2(tu, tv)
        TextureSize = glm_vec2(su, sv)
        transformMat = glm_mat4_mul(glm_mat4_mul(mvp, glm_translate(glm_vec3(offset.x, offset.y, 0!))), glm_scale(glm_vec3(size.x, size.y, 1!)))

        cb.transform_matrix = transformMat
        cb.offset = TextureUV
        cb.size = TextureSize
        cb.color = pielightToRGBAVec4(requests(r).colour)
        cb.texture = 0
        gfx_api_pso_get(DrawImagePSO).bind_constants VarPtr(cb), LenB(cb)

        gfx_api_pso_get(DrawImagePSO).bind_textures TextureID

        If Not didEnableRect Then
            gfx_api_pso_get(DrawImagePSO).bind_vertex_buffers pie_internal_rectBuffer
            didEnableRect = True
        End If

        gfx_api_pso_get(DrawImagePSO).draw 4, 0
    Next
    gfx_api_pso_get(DrawImagePSO).unbind_vertex_buffers pie_internal_rectBuffer
End Sub

' dest: Optional (C++: PIERECT *dest = nullptr): hasDest
Private Function makePieImage(ByVal imageFile_ As Long, ByVal id As Long, ByVal hasDest As Boolean, dest As PIERECT, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0) As Vector2i
    Dim image As AtlasImageDef, pieImage As Vector2i
    image = imageFiles(imageFile_).imageDefs(id)
    pieImage.x = image.Width
    pieImage.y = image.Height
    If hasDest Then
        dest.x = x + image.XOffset
        dest.y = y + image.YOffset
        dest.w = image.Width
        dest.h = image.Height
    End If
    makePieImage = pieImage
End Function

Private Function makePieImagef(ByVal imageFile_ As Long, ByVal id As Long, ByVal hasDest As Boolean, dest As PIERECT, ByVal x As Single, ByVal y As Single) As Vector2i
    Dim image As AtlasImageDef, pieImage As Vector2i
    image = imageFiles(imageFile_).imageDefs(id)
    pieImage.x = image.Width
    pieImage.y = image.Height
    If hasDest Then
        dest.x = x + CSng(image.XOffset)
        dest.y = y + CSng(image.YOffset)
        dest.w = CSng(image.Width)
        dest.h = CSng(image.Height)
    End If
    makePieImagef = pieImage
End Function

Public Sub iV_DrawImage2(ByVal filename As String, ByVal x As Single, ByVal y As Single, Optional ByVal width As Single = 0!, Optional ByVal height As Single = 0!)
    Dim image As Long
    If Len(filename) = 0 Then Exit Sub
    image = iV_GetImage(filename)
    If Not (image <> 0) Then ASSERT_FAILURE "image != nullptr", "lib/ivis_opengl/pieblitfunc.cpp:690", "iV_DrawImage2", "No image found for filename: %s", filename: Exit Sub
    iV_DrawImage2_image image, x, y, width, height
End Sub

' void iV_DrawImage2(const AtlasImageDef *image, float x, float y, float width = -0.0f, float height = -0.0f)
' image: the AtlasImageDef's packed handle
Public Sub iV_DrawImage2_image(ByVal image As Long, ByVal x As Single, ByVal y As Single, Optional ByVal width As Single = 0!, Optional ByVal height As Single = 0!)
    Dim imgDef As AtlasImageDef, invTextureSize As Single, tu As Long, tv As Long, w As Long, h As Long, mvp As mat4
    If Not (image <> 0) Then ASSERT_FAILURE "(image) != nullptr", "lib/ivis_opengl/pieblitfunc.cpp:696", "iV_DrawImage2", "%s should not be nullptr", "image": Exit Sub
    imgDef = AtlasImageDef_get(image)
    invTextureSize = imgDef.invTextureSize
    tu = imgDef.Tu
    tv = imgDef.Tv
    If width > 0 Then w = Fix(width) Else w = imgDef.Width
    If height > 0 Then h = Fix(height) Else h = imgDef.Height
    x = x + imgDef.XOffset
    y = y + imgDef.YOffset

    mvp = glm_mat4_mul(defaultProjectionMatrix(), glm_translate(glm_vec3(x, y, 0!)))
    iv_DrawImageImpl DrawImagePSO, pie_Texture(imgDef.textureId), glm_vec2(0!, 0!), glm_vec2(w, h), _
        glm_vec2(tu * invTextureSize, tv * invTextureSize), _
        glm_vec2(imgDef.Width * invTextureSize, imgDef.Height * invTextureSize), _
        psPalette(1), mvp                ' WZCOL_WHITE
End Sub

' pBatchedRequests: Nothing = nullptr
Public Sub iV_DrawImage(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, modelViewProjection As mat4, Optional ByVal pBatchedRequests As BatchedImageDrawRequests, Optional ByVal alpha As Long = 255)
    Dim dest As PIERECT, pieImage As Vector2i, noInset As Vector2i
    If Not assertValidImage(imageFile_, ID) Then
        Exit Sub
    End If

    pieImage = makePieImage(imageFile_, ID, True, dest, x, y)

    If pBatchedRequests Is Nothing Then
        gfx_api_pso_get(DrawImagePSO).bind
        pie_DrawImage imageFile_, ID, pieImage, dest, pal_RGBA(255, 255, 255, alpha), modelViewProjection, noInset
    Else
        pBatchedRequests.queuePieImageDraw_args REND_ALPHA, imageFile_, ID, pieImage, dest, pal_RGBA(255, 255, 255, alpha), modelViewProjection, noInset
        pBatchedRequests.draw    ' draw only if not deferred
    End If
End Sub

' size: optional<Vector2f> (has_size)
Public Sub iV_DrawImageTint(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Single, ByVal y As Single, ByVal color As Long, ByVal has_size As Boolean, size As Vector2f, modelViewProjection As mat4, Optional ByVal pBatchedRequests As BatchedImageDrawRequests)
    Dim dest As PIERECT, pieImage As Vector2i, noInset As Vector2i
    If Not assertValidImage(imageFile_, ID) Then
        Exit Sub
    End If

    pieImage = makePieImagef(imageFile_, ID, True, dest, x, y)
    If has_size Then
        dest.w = size.x
        dest.h = size.y
    End If

    If pBatchedRequests Is Nothing Then
        gfx_api_pso_get(DrawImagePSO).bind
        pie_DrawImage imageFile_, ID, pieImage, dest, color, modelViewProjection, noInset
    Else
        pBatchedRequests.queuePieImageDraw_args REND_ALPHA, imageFile_, ID, pieImage, dest, color, modelViewProjection, noInset
        pBatchedRequests.draw    ' draw only if not deferred
    End If
End Sub

Public Sub iV_DrawImageFileAnisotropicTint(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, size As Vector2f, ByVal color As Long, modelViewProjection As mat4)
    Dim dest As PIERECT, pieImage As Vector2i, noInset As Vector2i
    If Not assertValidImage(imageFile_, ID) Then
        Exit Sub
    End If

    pieImage = makePieImage(imageFile_, ID, True, dest, x, y)
    dest.w = size.x
    dest.h = size.y

    gfx_api_pso_get(DrawImageAnisotropicPSO).bind
    pie_DrawImageTemplate DrawImageAnisotropicPSO, imageFile_, ID, pieImage, dest, color, modelViewProjection, noInset
End Sub

Public Sub iV_DrawImageFileAnisotropic(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, size As Vector2f, modelViewProjection As mat4, Optional ByVal alpha As Long = 255)
    iV_DrawImageFileAnisotropicTint imageFile_, ID, x, y, size, pal_RGBA(255, 255, 255, alpha), modelViewProjection
End Sub

' image, imageTc: AtlasImage as (images, id)
Public Sub iV_DrawImageTc(ByVal image_images As Long, ByVal image_id As Long, ByVal imageTc_images As Long, ByVal imageTc_id As Long, ByVal x As Long, ByVal y As Long, ByVal colour As Long, modelViewProjection As mat4)
    Dim dest As PIERECT, pieImage As Vector2i, pieImageTc As Vector2i, noDest As PIERECT, noInset As Vector2i
    If Not assertValidImage(image_images, image_id) Then
        Exit Sub
    End If
    If Not assertValidImage(imageTc_images, imageTc_id) Then
        Exit Sub
    End If

    pieImage = makePieImage(image_images, image_id, True, dest, x, y)
    pieImageTc = makePieImage(imageTc_images, imageTc_id, False, noDest)

    gfx_api_pso_get(DrawImagePSO).bind

    pie_DrawImage image_images, image_id, pieImage, dest, psPalette(1), modelViewProjection, noInset  ' WZCOL_WHITE
    pie_DrawImage imageTc_images, imageTc_id, pieImageTc, dest, colour, modelViewProjection, noInset
End Sub

' pieblitfunc.h: static inline void iV_DrawImageImage(AtlasImage image, int x, int y, uint8_t alpha = 255)
Public Sub iV_DrawImageImage(ByVal image_images As Long, ByVal image_id As Long, ByVal x As Long, ByVal y As Long, Optional ByVal alpha As Long = 255)
    Dim m As mat4
    m = defaultProjectionMatrix()
    iV_DrawImage image_images, image_id, x, y, m, Nothing, alpha
End Sub

' pieblitfunc.h: static inline void iV_DrawImageTc(IMAGEFILE *imageFile, unsigned id, unsigned idTc, int x, int y, PIELIGHT colour)
Public Sub iV_DrawImageTc_ids(ByVal imageFile_ As Long, ByVal id As Long, ByVal idTc As Long, ByVal x As Long, ByVal y As Long, ByVal colour As Long)
    Dim m As mat4
    m = defaultProjectionMatrix()
    iV_DrawImageTc imageFile_, id, imageFile_, idTc, x, y, colour, m
End Sub

' Repeat a texture
Public Sub iV_DrawImageRepeatX(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, modelViewProjection As mat4, Optional ByVal enableHorizontalTilingSeamWorkaround As Boolean = False, Optional ByVal pBatchedRequests As BatchedImageDrawRequests)
    Dim Image As AtlasImageDef, mode As REND_MODE, dest As PIERECT, pieImage As Vector2i
    Dim usableImageWidth As Long, imageXInset As Long, hRemainder As Long, hRep As Long, inset As Vector2i
    If width <= 0 Then
        Exit Sub
    End If

    If pBatchedRequests Is Nothing Then
        If repeatXLocalBatch Is Nothing Then Set repeatXLocalBatch = New BatchedImageDrawRequests
        repeatXLocalBatch.clear
        Set pBatchedRequests = repeatXLocalBatch
    End If

    assertValidImage imageFile_, ID
    Image = imageFiles(imageFile_).imageDefs(ID)

    mode = REND_OPAQUE

    pieImage = makePieImage(imageFile_, ID, True, dest, x, y)

    usableImageWidth = Image.Width
    imageXInset = 0
    If enableHorizontalTilingSeamWorkaround Then
        ' Inset the portion of the image that is used by 1 on both the left + right sides
        usableImageWidth = usableImageWidth - 2
        imageXInset = 1
        dest.w = dest.w - 2
    End If
    ' assert(usableImageWidth > 0); (a release build has none)

    hRemainder = (width Mod usableImageWidth)
    inset.x = imageXInset
    inset.y = 0

    For hRep = 0 To (width \ usableImageWidth) - 1
        pBatchedRequests.queuePieImageDraw_args mode, imageFile_, ID, pieImage, dest, psPalette(1), modelViewProjection, inset  ' WZCOL_WHITE
        dest.x = dest.x + usableImageWidth
    Next

    ' draw remainder
    If hRemainder > 0 Then
        pieImage.x = hRemainder
        dest.w = hRemainder
        pBatchedRequests.queuePieImageDraw_args mode, imageFile_, ID, pieImage, dest, psPalette(1), modelViewProjection, inset  ' WZCOL_WHITE
    End If

    ' draw batched requests (unless batch is deferred)
    pBatchedRequests.draw
End Sub

Public Sub iV_DrawImageRepeatY(ByVal imageFile_ As Long, ByVal ID As Long, ByVal x As Long, ByVal y As Long, ByVal height As Long, modelViewProjection As mat4, Optional ByVal pBatchedRequests As BatchedImageDrawRequests)
    Dim Image As AtlasImageDef, mode As REND_MODE, dest As PIERECT, pieImage As Vector2i
    Dim vRemainder As Long, vRep As Long, noInset As Vector2i
    If height <= 0 Then
        Exit Sub
    End If

    If pBatchedRequests Is Nothing Then
        If repeatYLocalBatch Is Nothing Then Set repeatYLocalBatch = New BatchedImageDrawRequests
        repeatYLocalBatch.clear
        Set pBatchedRequests = repeatYLocalBatch
    End If

    assertValidImage imageFile_, ID
    Image = imageFiles(imageFile_).imageDefs(ID)

    mode = REND_OPAQUE

    pieImage = makePieImage(imageFile_, ID, True, dest, x, y)

    vRemainder = height Mod Image.Height

    For vRep = 0 To (height \ Image.Height) - 1
        pBatchedRequests.queuePieImageDraw_args mode, imageFile_, ID, pieImage, dest, psPalette(1), modelViewProjection, noInset  ' WZCOL_WHITE
        dest.y = dest.y + Image.Height
    Next

    ' draw remainder
    If vRemainder > 0 Then
        pieImage.y = vRemainder
        dest.h = vRemainder
        pBatchedRequests.queuePieImageDraw_args mode, imageFile_, ID, pieImage, dest, psPalette(1), modelViewProjection, noInset  ' WZCOL_WHITE
    End If

    ' draw batched requests (unless batch is deferred)
    pBatchedRequests.draw
End Sub

Public Function pie_InitRadar() As Boolean
    Dim i As Long
    For i = 0 To NUM_RADAR_TEXTURES - 1
        Set radarGfx(i) = GFX_new(GFX_TEXTURE, 2)
    Next
    currRadarGfx = 0
    pie_InitRadar = True
End Function

Public Function pie_ShutdownRadar() As Boolean
    Dim i As Long
    For i = 0 To NUM_RADAR_TEXTURES - 1
        If radarGfx(i) Is Nothing Then
            GoTo continue_
        End If
        Set radarGfx(i) = Nothing
continue_:
    Next
    currRadarGfx = 0
    pie_ViewingWindow_Shutdown
    pie_ShutdownRadar = True
End Function

Public Sub pie_SetRadar(ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal twidth As Long, ByVal theight As Long)
    Dim i As Long, texcoords(0 To 7) As Single, vertices(0 To 7) As Single
    For i = 0 To NUM_RADAR_TEXTURES - 1
        If radarGfx(i) Is Nothing Then
            GoTo continue_
        End If
        radarGfx(i).makeTexture twidth, theight, pixel_format_FORMAT_RGBA8_UNORM_PACK8, "mem::radarTexture[" & CStr(i) & "]"
        '	glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST);  // Want GL_LINEAR (or GL_LINEAR_MIPMAP_NEAREST) for min filter, but GL_NEAREST for mag filter. // TODO: Add a gfx_api::sampler_type to handle this case? bilinear, but nearest for mag?
        texcoords(0) = 0!: texcoords(1) = 0!: texcoords(2) = 1!: texcoords(3) = 0!
        texcoords(4) = 0!: texcoords(5) = 1!: texcoords(6) = 1!: texcoords(7) = 1!
        vertices(0) = x: vertices(1) = y: vertices(2) = x + width: vertices(3) = y
        vertices(4) = x: vertices(5) = y + height: vertices(6) = x + width: vertices(7) = y + height
        radarGfx(i).buffers 4, VarPtr(vertices(0)), VarPtr(texcoords(0))
continue_:
    Next
End Sub

'* Store radar texture with given width and height.
Public Sub pie_DownLoadRadar(ByVal bitmap As iV_Image)
    currRadarGfx = currRadarGfx + 1
    If currRadarGfx >= NUM_RADAR_TEXTURES Then
        currRadarGfx = 0
    End If
    radarGfx(currRadarGfx).updateTexture bitmap
End Sub

'* Display radar texture using the given height and width, depending on zoom level.
Public Sub pie_RenderRadar(modelViewProjectionMatrix As mat4)
    Dim cb As constant_buffer_type_SHADER_GFX_TEXT
    gfx_api_pso_get(RadarPSO).bind
    cb.transform_matrix = modelViewProjectionMatrix
    cb.offset = glm_vec2(0!, 0!)
    cb.size = glm_vec2(0!, 0!)
    cb.color = glm_vec4(1!, 1!, 1!, 1!)
    cb.texture = 0
    gfx_api_pso_get(RadarPSO).bind_constants VarPtr(cb), LenB(cb)
    radarGfx(currRadarGfx).draw RadarPSO, modelViewProjectionMatrix
End Sub

'/ Load and display a random backdrop picture.
Public Sub pie_LoadBackDrop(ByVal screenType_ As SCREENTYPE)
    Select Case screenType_
    Case SCREEN_RANDOMBDROP
        screen_SetRandomBackdrop "texpages/bdrops/", "backdrop"
    Case SCREEN_MISSIONEND
        screen_SetRandomBackdrop "texpages/bdrops/", "missionend"
    End Select
End Sub

Public Function defaultProjectionMatrix() As mat4
    Dim w As Single, h As Single
    w = pie_GetVideoBufferWidth()
    h = pie_GetVideoBufferHeight()

    defaultProjectionMatrix = glm_ortho(0!, w, h, 0!)
End Function

' Debug: Draw a texture to a screen quad, given a position, offset, size (in screen coordinates), angle, and colour
' Optionally takes a channelSwizzle - glm::ivec4(0,0,0,3) for example, to just use the red channel (preserving the last channel)
' (VB) channelSwizzle = glm::ivec4(0,1,2,3) and textureUVTransform = glm::mat4(1.f) are the C++ defaults
Public Sub iV_DebugDrawTextureToQuad(ByVal texture As gfx_api_abstract_texture, ByVal layer As Long, Position As Vector2f, offset As Vector2f, size As Vector2f, ByVal angle As Single, ByVal colour As Long, channelSwizzle As Vector4i, textureUVTransform As mat4)
    Dim modelViewProjection As mat4, mvpFinal As mat4, TextureUV As Vector2f, TextureSize As Vector2f, uvTransformMatrix As mat4
    Dim cba As constant_buffer_type_SHADER_DEBUG_TEXTURE2DARRAY_QUAD, cb As constant_buffer_type_SHADER_DEBUG_TEXTURE2D_QUAD
    modelViewProjection = glm_mat4_mul(glm_mat4_mul(defaultProjectionMatrix(), glm_translate(glm_vec3(Position.x, Position.y, 0!))), glm_rotate(RADIANS(angle), glm_vec3(0!, 0!, 1!)))
    mvpFinal = glm_mat4_mul(glm_mat4_mul(modelViewProjection, glm_translate(glm_vec3(offset.x, offset.y, 0!))), glm_scale(glm_vec3(size.x, size.y, 1!)))

    TextureUV = glm_vec2(0!, 0!)
    TextureSize = glm_vec2(1!, 1!)

    uvTransformMatrix = glm_mat4_mul(glm_mat4_mul(glm_mat4_mul(glm_mat4(1!), glm_scale(glm_vec3(TextureSize.x, TextureSize.y, 1!))), glm_translate(glm_vec3(TextureUV.x, TextureUV.y, 0!))), textureUVTransform)

    If texture.isArray() Then
        gfx_api_pso_get(DebugDrawTexture2DArrayToQuad).bind
        cba.transform_matrix = mvpFinal
        cba.uv_transform_matrix = uvTransformMatrix
        cba.swizzle = channelSwizzle
        cba.color = pielightToRGBAVec4(colour)
        cba.layer = layer
        cba.texture = 0
        gfx_api_pso_get(DebugDrawTexture2DArrayToQuad).bind_constants VarPtr(cba), LenB(cba)
        gfx_api_pso_get(DebugDrawTexture2DArrayToQuad).bind_textures texture
        gfx_api_pso_get(DebugDrawTexture2DArrayToQuad).bind_vertex_buffers pie_internal_rectBuffer
        gfx_api_pso_get(DebugDrawTexture2DArrayToQuad).draw 4, 0
        gfx_api_pso_get(DebugDrawTexture2DArrayToQuad).unbind_vertex_buffers pie_internal_rectBuffer
    Else
        gfx_api_pso_get(DebugDrawTexture2DToQuad).bind
        cb.transform_matrix = mvpFinal
        cb.uv_transform_matrix = uvTransformMatrix
        cb.swizzle = channelSwizzle
        cb.color = pielightToRGBAVec4(colour)
        cb.texture = 0
        gfx_api_pso_get(DebugDrawTexture2DToQuad).bind_constants VarPtr(cb), LenB(cb)
        gfx_api_pso_get(DebugDrawTexture2DToQuad).bind_textures texture
        gfx_api_pso_get(DebugDrawTexture2DToQuad).bind_vertex_buffers pie_internal_rectBuffer
        gfx_api_pso_get(DebugDrawTexture2DToQuad).draw 4, 0
        gfx_api_pso_get(DebugDrawTexture2DToQuad).unbind_vertex_buffers pie_internal_rectBuffer
    End If
End Sub
