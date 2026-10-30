Attribute VB_Name = "bitimage_c"
' lib/ivis_opengl/bitimage.cpp, lib/ivis_opengl/bitimage.h, lib/ivis_opengl/ivisdef.h:
' the image sheets (.img files: one "x,y,name" line per image, the images in
' the folder of the same name), packed onto texture pages at load.
' IMAGEFILE * is a handle into imageFiles (1-based, 0 = nullptr; ivisdef_h's
' IMAGEFILE); files is the C++ std::vector<IMAGEFILE *> of the loaded ones.
' AtlasImageDef * is the handle file * &H10000 + index + 1 (0 = nullptr):
' AtlasImageDef_get gives WzGfx the struct, the AtlasImageDef_<field> procedures
' give the exe its fields. AtlasImage {images, id} is a UDT of the exe, which
' calls AtlasImage_width ... here with its two members.
' ImageMerge::arrange's std::multiset is a sorted array (an insert goes after
' the equal elements, as the multiset's); std::sort of the images is a stable
' insertion sort (std::sort orders equal rectangles in its own way, so the
' packing can differ from the C++ build's where images have equal sizes: the
' images' texture coordinates follow the packing either way).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public imageFiles() As IMAGEFILE        ' (VB) the pool of IMAGEFILEs (1-based)
Private imageFiles_size As Long

Private files() As Long                 ' static std::vector<IMAGEFILE *> files;
Private files_size As Long

Private Type ImageMergeRectangle
    index As Long                       ' Index in ImageDefs array.
    page As Long                        ' size_t: Texture page index.
    loc As Vector2i                     ' = Vector2i(0, 0)
    siz As Vector2i                     ' = Vector2i(0, 0)
    data As iV_Image
End Type

Private Type ImageMerge
    images() As ImageMergeRectangle
    images_size As Long
    pages() As Long                     ' List of page sizes, normally all pageSize, unless an image is too large for a normal page.
    pages_size As Long
End Type

Private Const ImageMerge_pageSize As Long = 256    ' static const int pageSize = 256;

' Used to provide empty space between sprites arranged on the page, to avoid display glitches when scaling + drawing
Private Const SPRITE_BUFFER_PIXELS As Long = 1

' bool ImageMergeRectangle::operator <(ImageMergeRectangle const &b) const
Private Function ImageMergeRectangle_less(a As ImageMergeRectangle, b As ImageMergeRectangle) As Boolean
    Dim amax As Long, bmax As Long, amin As Long, bmin As Long
    If a.siz.x > a.siz.y Then amax = a.siz.x: amin = a.siz.y Else amax = a.siz.y: amin = a.siz.x
    If b.siz.x > b.siz.y Then bmax = b.siz.x: bmin = b.siz.y Else bmax = b.siz.y: bmin = b.siz.x
    If amax <> bmax Then
        ImageMergeRectangle_less = amax < bmax
        Exit Function
    End If
    If amin <> bmin Then
        ImageMergeRectangle_less = amin < bmin
        Exit Function
    End If
    ImageMergeRectangle_less = a.siz.x < b.siz.x
End Function

' (VB) the handle of imageDefs[index] of a file
Public Function AtlasImageDef_handle(ByVal imgFile As Long, ByVal index As Long) As Long
    AtlasImageDef_handle = imgFile * &H10000 + index + 1
End Function

' (VB) *def: the struct an AtlasImageDef handle points at
Public Function AtlasImageDef_get(ByVal def As Long) As AtlasImageDef
    AtlasImageDef_get = imageFiles(def \ &H10000).imageDefs((def And &HFFFF&) - 1)
End Function

' (VB) the fields of an AtlasImageDef handle, for the exe
Public Function AtlasImageDef_Width(ByVal def As Long) As Long
    AtlasImageDef_Width = imageFiles(def \ &H10000).imageDefs((def And &HFFFF&) - 1).Width
End Function

Public Function AtlasImageDef_Height(ByVal def As Long) As Long
    AtlasImageDef_Height = imageFiles(def \ &H10000).imageDefs((def And &HFFFF&) - 1).Height
End Function

Public Function AtlasImageDef_XOffset(ByVal def As Long) As Long
    AtlasImageDef_XOffset = imageFiles(def \ &H10000).imageDefs((def And &HFFFF&) - 1).XOffset
End Function

Public Function AtlasImageDef_YOffset(ByVal def As Long) As Long
    AtlasImageDef_YOffset = imageFiles(def \ &H10000).imageDefs((def And &HFFFF&) - 1).YOffset
End Function

Public Function AtlasImageDef_textureId(ByVal def As Long) As Long
    AtlasImageDef_textureId = imageFiles(def \ &H10000).imageDefs((def And &HFFFF&) - 1).textureId
End Function

' ivisdef.h: AtlasImage's members (images: an IMAGEFILE handle, id)
Public Function AtlasImage_isNull(ByVal images As Long) As Boolean
    AtlasImage_isNull = (images = 0)
End Function

Public Function AtlasImage_width(ByVal images As Long, ByVal id As Long) As Long
    AtlasImage_width = imageFiles(images).imageDefs(id).Width
End Function

Public Function AtlasImage_height(ByVal images As Long, ByVal id As Long) As Long
    AtlasImage_height = imageFiles(images).imageDefs(id).Height
End Function

Public Function AtlasImage_xOffset(ByVal images As Long, ByVal id As Long) As Long
    AtlasImage_xOffset = imageFiles(images).imageDefs(id).XOffset
End Function

Public Function AtlasImage_yOffset(ByVal images As Long, ByVal id As Long) As Long
    AtlasImage_yOffset = imageFiles(images).imageDefs(id).YOffset
End Function

' bitimage.h
Public Function iV_GetImageWidth(ByVal imgFile As Long, ByVal ID As Long) As Long
    ID = ID And &HFFFF&                 ' const unsigned short ID
    If Not (ID < imageFiles(imgFile).imageDefs_size) Then ASSERT_FAILURE "ID < ImageFile->imageDefs.size()", "lib/ivis_opengl/bitimage.h:29", "iV_GetImageWidth", "Invalid image ID: %u", ID: iV_GetImageWidth = 0: Exit Function
    iV_GetImageWidth = AtlasImage_width(imgFile, ID) And &HFFFF&
End Function

Public Function iV_GetImageHeight(ByVal imgFile As Long, ByVal ID As Long) As Long
    ID = ID And &HFFFF&
    If Not (ID < imageFiles(imgFile).imageDefs_size) Then ASSERT_FAILURE "ID < ImageFile->imageDefs.size()", "lib/ivis_opengl/bitimage.h:36", "iV_GetImageHeight", "Invalid image ID: %u", ID: iV_GetImageHeight = 0: Exit Function
    iV_GetImageHeight = AtlasImage_height(imgFile, ID) And &HFFFF&
End Function

Public Function iV_GetImageXOffset(ByVal imgFile As Long, ByVal ID As Long) As Long
    ID = ID And &HFFFF&
    ' assert(ID < ImageFile->imageDefs.size()): a release build's assert (none)
    iV_GetImageXOffset = AtlasImage_xOffset(imgFile, ID)
End Function

Public Function iV_GetImageYOffset(ByVal imgFile As Long, ByVal ID As Long) As Long
    ID = ID And &HFFFF&
    ' assert(ID < ImageFile->imageDefs.size()): a release build's assert (none)
    iV_GetImageYOffset = AtlasImage_yOffset(imgFile, ID)
End Function

' inline size_t IMAGEFILE::numImages() const
Public Function IMAGEFILE_numImages(ByVal imgFile As Long) As Long
    IMAGEFILE_numImages = imageFiles(imgFile).imageDefs_size
End Function

Public Function iV_GetImage(ByVal filename As String) As Long
    Dim i As Long, pImageDef As Long
    For i = 0 To files_size - 1
        pImageDef = IMAGEFILE_find(files(i), filename)
        If pImageDef <> 0 Then
            iV_GetImage = pImageDef
            Exit Function
        End If
    Next
    debug_ 76, LOG_ERROR, "iV_GetImage", "%s not found in image list!", filename
    iV_GetImage = 0
End Function

Private Sub checkRect(rect As ImageMergeRectangle, ByVal pageSize As Long)
    If (rect.loc.x + rect.siz.x) > pageSize Then
        debug_ 87, LOG_ERROR, "checkRect", "Merge rectangle bounds extend outside of pageSize bounds: %d > %d", (rect.loc.x + rect.siz.x), pageSize
    End If
    If (rect.loc.y + rect.siz.y) > pageSize Then
        debug_ 91, LOG_ERROR, "checkRect", "Merge rectangle bounds extend outside of pageSize bounds: %d > %d", (rect.loc.y + rect.siz.y), pageSize
    End If
    ' assert(rect.loc.x >= 0); assert(rect.loc.y >= 0): a release build's asserts (none)
End Sub

' (VB) std::multiset<ImageMergeRectangle>::insert: after the elements equal to r; returns the position
Private Function freeSpace_insert(freeSpace() As ImageMergeRectangle, freeSpace_size As Long, r As ImageMergeRectangle) As Long
    Dim pos As Long, i As Long
    pos = freeSpace_size
    For i = 0 To freeSpace_size - 1
        If ImageMergeRectangle_less(r, freeSpace(i)) Then
            pos = i
            Exit For
        End If
    Next
    If freeSpace_size > UBound(freeSpace) Then ReDim Preserve freeSpace(0 To 2 * freeSpace_size + 15)
    For i = freeSpace_size To pos + 1 Step -1
        freeSpace(i) = freeSpace(i - 1)
    Next
    freeSpace(pos) = r
    freeSpace_size = freeSpace_size + 1
    freeSpace_insert = pos
End Function

Private Sub freeSpace_erase(freeSpace() As ImageMergeRectangle, freeSpace_size As Long, ByVal pos As Long)
    Dim i As Long, blank As ImageMergeRectangle
    For i = pos To freeSpace_size - 2
        freeSpace(i) = freeSpace(i + 1)
    Next
    freeSpace_size = freeSpace_size - 1
    freeSpace(freeSpace_size) = blank
End Sub

' inline void ImageMerge::arrange()
Private Sub ImageMerge_arrange(m As ImageMerge)
    Dim freeSpace() As ImageMergeRectangle, freeSpace_size As Long  ' std::multiset<ImageMergeRectangle> freeSpace;
    Dim r As Long, f As Long, s As Long, i As Long, j As Long, p As Long
    Dim newPage As ImageMergeRectangle, spRight As ImageMergeRectangle, spDown As ImageMergeRectangle
    Dim blank As ImageMergeRectangle, t As ImageMergeRectangle
    ReDim freeSpace(0 To 15)
    m.pages_size = 0

    ' std::sort(images.begin(), images.end()); (VB: stable insertion sort)
    For i = 1 To m.images_size - 1
        t = m.images(i)
        j = i - 1
        Do While j >= 0
            If Not ImageMergeRectangle_less(t, m.images(j)) Then Exit Do
            m.images(j + 1) = m.images(j)
            j = j - 1
        Loop
        m.images(j + 1) = t
    Next

    r = m.images_size
    Do While r <> 0
        r = r - 1

        ' Find smallest free rectangle which is large enough. (lower_bound: the first not less than *r)
        f = freeSpace_size
        For i = 0 To freeSpace_size - 1
            If Not ImageMergeRectangle_less(freeSpace(i), m.images(r)) Then
                f = i
                Exit For
            End If
        Next
        Do While f <> freeSpace_size
            If Not (freeSpace(f).siz.x < m.images(r).siz.x Or freeSpace(f).siz.y < m.images(r).siz.y) Then Exit Do
            f = f + 1  ' Rectangle has wrong shape.
        Loop
        If f = freeSpace_size Then
            ' No free space, make new page.
            s = ImageMerge_pageSize
            Do While s < (m.images(r).siz.x Or m.images(r).siz.y)
                s = s * 2
            Loop
            newPage = blank
            newPage.page = m.pages_size
            newPage.loc.x = 0: newPage.loc.y = 0
            newPage.siz.x = s: newPage.siz.y = s
            If m.pages_size > UBound(m.pages) Then ReDim Preserve m.pages(0 To 2 * m.pages_size + 3)
            m.pages(m.pages_size) = s
            m.pages_size = m.pages_size + 1
            f = freeSpace_insert(freeSpace, freeSpace_size, newPage)
        End If
        m.images(r).page = freeSpace(f).page
        m.images(r).loc = freeSpace(f).loc
        spRight = blank
        spDown = blank
        spRight.page = freeSpace(f).page
        spDown.page = freeSpace(f).page
        spRight.loc.x = freeSpace(f).loc.x + (m.images(r).siz.x + SPRITE_BUFFER_PIXELS): spRight.loc.y = freeSpace(f).loc.y
        spDown.loc.x = freeSpace(f).loc.x: spDown.loc.y = freeSpace(f).loc.y + (m.images(r).siz.y + SPRITE_BUFFER_PIXELS)
        spRight.siz.x = freeSpace(f).siz.x - (m.images(r).siz.x + SPRITE_BUFFER_PIXELS): spRight.siz.y = m.images(r).siz.y
        spDown.siz.x = m.images(r).siz.x: spDown.siz.y = freeSpace(f).siz.y - (m.images(r).siz.y + SPRITE_BUFFER_PIXELS)
        If spRight.siz.x <= spDown.siz.y Then
            ' Split horizontally.
            spDown.siz.x = freeSpace(f).siz.x
        Else
            ' Split vertically.
            spRight.siz.y = freeSpace(f).siz.y
        End If
        checkRect spDown, m.pages(spDown.page)
        checkRect spRight, m.pages(spRight.page)
        If spRight.siz.x > 0 And spRight.siz.y > 0 Then
            p = freeSpace_insert(freeSpace, freeSpace_size, spRight)
            If p <= f Then f = f + 1    ' (VB) the iterator f keeps pointing at its element
        End If
        If spDown.siz.x > 0 And spDown.siz.y > 0 Then
            p = freeSpace_insert(freeSpace, freeSpace_size, spDown)
            If p <= f Then f = f + 1
        End If
        freeSpace_erase freeSpace, freeSpace_size, f
    Loop
End Sub

' (VB) new IMAGEFILE: a free slot of the pool
Private Function IMAGEFILE_new() As Long
    Dim i As Long, blank As IMAGEFILE
    For i = 1 To imageFiles_size
        If Not imageFiles(i).inUse Then Exit For
    Next
    If i > imageFiles_size Then
        imageFiles_size = imageFiles_size + 1
        ReDim Preserve imageFiles(1 To imageFiles_size)
        i = imageFiles_size
    End If
    imageFiles(i) = blank
    imageFiles(i).inUse = True
    Set imageFiles(i).imageNamesMap = CreateObject("Scripting.Dictionary")
    imageFiles(i).imageNamesMap.CompareMode = 0    ' BinaryCompare
    IMAGEFILE_new = i
End Function

' (VB) delete imgFile: the destructor, then the slot is free
Private Sub IMAGEFILE_delete(ByVal imgFile As Long)
    Dim blank As IMAGEFILE
    IMAGEFILE_destructor imgFile
    imageFiles(imgFile) = blank
End Sub

Public Function iV_LoadImageFile(ByVal fileName As String) As Long
    Dim imageDir As String, pFileData() As Byte, pFileSize As Long, chars As String
    Dim ptr As Long, numImages As Long, imgFile As Long, pageLayout As ImageMerge
    Dim temp As Long, retval As Long, tmpName As String, spriteName As String, nameBytes() As Byte
    Dim xo As Long, yo As Long, i As Long, p As Long, size As Long
    Dim ivImages() As iV_Image, srcImage As iV_Image, dstImage As iV_Image
    Dim srcChannels As Long, srcStride As Long, srcBytes As Long, dstChannels As Long, dstStride As Long, dstBytes As Long
    Dim x As Long, y As Long, rgba(0 To 3) As Byte, arbitraryName As String, pTexture As gfx_api_texture

    ' Find the directory of images.
    imageDir = fileName
    If InStrRev(imageDir, ".") > 0 Then
        imageDir = Left$(imageDir, InStrRev(imageDir, ".") - 1)
    End If
    imageDir = imageDir & "/"

    ' Load the image list file.
    If Not loadFile(fileName, pFileData, pFileSize) Then
        debug_ 178, LOG_ERROR, "iV_LoadImageFile", "iV_LoadImageFile: failed to open %s", fileName
        iV_LoadImageFile = 0
        Exit Function
    End If
    ' (VB) the bytes as characters, one each, so that offsets are the C++ pointer's
    chars = String$(pFileSize, vbNullChar)
    For i = 0 To pFileSize - 1
        Mid$(chars, i + 1, 1) = ChrW$(pFileData(i))
    Next

    ptr = 0
    ' count lines, which is identical to number of images
    numImages = 0
    Do While ptr < pFileSize
        If pFileData(ptr) = 0 Then Exit Do
        If pFileData(ptr) = 10 Then numImages = numImages + 1
        ptr = ptr + 1
    Loop
    imgFile = IMAGEFILE_new()
    imageFiles(imgFile).imageDefs_size = numImages
    If numImages > 0 Then ReDim imageFiles(imgFile).imageDefs(0 To numImages - 1)
    pageLayout.images_size = numImages
    If numImages > 0 Then ReDim pageLayout.images(0 To numImages - 1)
    ReDim pageLayout.pages(0 To 3)
    ptr = 0
    numImages = 0
    Do While ptr < pFileSize
        ' (VB) a last line without its newline: C++ writes past the vectors' ends; here they grow
        If numImages >= imageFiles(imgFile).imageDefs_size Then
            ReDim Preserve imageFiles(imgFile).imageDefs(0 To numImages)
            ReDim Preserve pageLayout.images(0 To numImages)
        End If

        tmpName = ""
        temp = 0
        retval = sscanf(Mid$(chars, ptr + 1), "%d,%d,%255[^" & vbCr & vbLf & """,]%n", xo, yo, tmpName, temp)
        If retval <> 3 Then
            debug_ 206, LOG_ERROR, "iV_LoadImageFile", "Bad line in ""%s"".", fileName
            IMAGEFILE_delete imgFile
            iV_LoadImageFile = 0
            Exit Function
        End If
        imageFiles(imgFile).imageDefs(numImages).XOffset = xo
        imageFiles(imgFile).imageDefs(numImages).YOffset = yo
        ' WzString::fromUtf8(tmpName): the name's bytes are its characters' codes
        ReDim nameBytes(0 To Len(tmpName))
        For i = 1 To Len(tmpName)
            nameBytes(i - 1) = AscW(Mid$(tmpName, i, 1)) And &HFF&
        Next
        If Not imageFiles(imgFile).imageNamesMap.Exists(utf8_decode(nameBytes, 0, Len(tmpName))) Then
            imageFiles(imgFile).imageNamesMap.Add utf8_decode(nameBytes, 0, Len(tmpName)), numImages
        End If
        spriteName = imageDir & utf8_decode(nameBytes, 0, Len(tmpName))

        pageLayout.images(numImages).index = numImages
        Set pageLayout.images(numImages).data = New iV_Image
        If Not iV_loadImage_PNG2(spriteName, pageLayout.images(numImages).data, True) Then
            debug_ 219, LOG_ERROR, "iV_LoadImageFile", "Failed to find image ""%s"" listed in ""%s"".", spriteName, fileName
            IMAGEFILE_delete imgFile
            iV_LoadImageFile = 0
            Exit Function
        End If
        pageLayout.images(numImages).siz.x = pageLayout.images(numImages).data.width()
        pageLayout.images(numImages).siz.y = pageLayout.images(numImages).data.height()
        numImages = numImages + 1
        ptr = ptr + temp
        ' skip rest of line
        Do While ptr < pFileSize
            ptr = ptr + 1
            If pFileData(ptr - 1) = 10 Then Exit Do
        Loop
    Loop
    Erase pFileData
    If numImages > imageFiles(imgFile).imageDefs_size Then
        imageFiles(imgFile).imageDefs_size = numImages
        pageLayout.images_size = numImages
    End If

    ImageMerge_arrange pageLayout  ' Arrange all the images onto texture pages (attempt to do so with as few pages as possible).
    imageFiles(imgFile).pages_size = pageLayout.pages_size
    If pageLayout.pages_size > 0 Then ReDim imageFiles(imgFile).pages(0 To pageLayout.pages_size - 1)

    If pageLayout.pages_size > 0 Then ReDim ivImages(0 To pageLayout.pages_size - 1)

    For p = 0 To pageLayout.pages_size - 1
        size = pageLayout.pages(p)
        Set ivImages(p) = New iV_Image
        ivImages(p).allocate size, size, 4, True
        imageFiles(imgFile).pages(p).size = size
        ' Must set imgFile->pages[p].id later, after filling out ivImages[p].bmp.
    Next

    For i = 0 To pageLayout.images_size - 1
        imageFiles(imgFile).imageDefs(pageLayout.images(i).index).TPageID = pageLayout.images(i).page
        imageFiles(imgFile).imageDefs(pageLayout.images(i).index).Tu = pageLayout.images(i).loc.x
        imageFiles(imgFile).imageDefs(pageLayout.images(i).index).Tv = pageLayout.images(i).loc.y
        imageFiles(imgFile).imageDefs(pageLayout.images(i).index).Width = pageLayout.images(i).siz.x
        imageFiles(imgFile).imageDefs(pageLayout.images(i).index).Height = pageLayout.images(i).siz.y

        ' Copy image data onto texture page.
        Set srcImage = pageLayout.images(i).data
        srcChannels = srcImage.channels()
        srcStride = srcImage.width() * srcChannels ' Not sure whether to pad in the case that srcDepth ? 4, however this apparently cannot happen.
        srcBytes = srcImage.bmp() + 0 * srcChannels + 0 * srcStride
        Set dstImage = ivImages(pageLayout.images(i).page)
        dstChannels = dstImage.channels()
        dstStride = dstImage.width() * dstChannels
        dstBytes = dstImage.bmp_w() + pageLayout.images(i).loc.x * dstChannels + pageLayout.images(i).loc.y * dstStride
        If srcChannels = dstChannels Then
            ' (VB) the per-pixel copy below, a row at a time (the same bytes)
            For y = 0 To pageLayout.images(i).siz.y - 1
                CopyMemory ByVal (dstBytes + y * dstStride), ByVal (srcBytes + y * srcStride), pageLayout.images(i).siz.x * dstChannels
            Next
        Else
            rgba(0) = &H0: rgba(1) = &H0: rgba(2) = &H0: rgba(3) = &HFF
            For y = 0 To pageLayout.images(i).siz.y - 1
                For x = 0 To pageLayout.images(i).siz.x - 1
                    CopyMemory rgba(0), ByVal (srcBytes + x * srcChannels + y * srcStride), srcChannels
                    CopyMemory ByVal (dstBytes + x * dstChannels + y * dstStride), rgba(0), dstChannels
                Next
            Next
        End If

        ' Finished reading the image data and copying it to the texture page, delete it.
        pageLayout.images(i).data.clear
        Set pageLayout.images(i).data = Nothing
        Set srcImage = Nothing
        Set dstImage = Nothing
    Next

    ' Upload texture pages and free image data.
    For p = 0 To pageLayout.pages_size - 1
        arbitraryName = va("%s-%03u", fileName, p)
        ' Now we can set imgFile->pages[p].id. This free()s the ivImages[p].bmp array!
        Set pTexture = gfx_api_context_loadTextureFromUncompressedImage(ivImages(p), texture_type_user_interface, arbitraryName)
        Set ivImages(p) = Nothing
        If pTexture Is Nothing Then
            ' Failed to load texture from image
            debug_ 300, LOG_ERROR, "iV_LoadImageFile", "Failed to load texture from image page [%u] listed in ""%s"".", p, fileName
            ' (VB) C++ leaks the IMAGEFILE (not in files: no destructor runs); its slot is freed silently
            imageFiles(imgFile).inUse = False
            iV_LoadImageFile = 0
            Exit Function
        End If
        imageFiles(imgFile).pages(p).id = pie_AddTexPage(pTexture, arbitraryName, texture_type_user_interface)
        Set pTexture = Nothing
    Next
    Erase ivImages

    ' duplicate some data, since we want another access point to these data structures now, FIXME
    For i = 0 To imageFiles(imgFile).imageDefs_size - 1
        imageFiles(imgFile).imageDefs(i).textureId = imageFiles(imgFile).pages(imageFiles(imgFile).imageDefs(i).TPageID).id
        imageFiles(imgFile).imageDefs(i).invTextureSize = 1! / imageFiles(imgFile).pages(imageFiles(imgFile).imageDefs(i).TPageID).size
    Next

    If files_size = 0 Then ReDim files(0 To 7)
    If files_size > UBound(files) Then ReDim Preserve files(0 To 2 * files_size + 7)
    files(files_size) = imgFile
    files_size = files_size + 1

    iV_LoadImageFile = imgFile
End Function

Public Sub iV_FreeImageFile(ByVal imgFile As Long)
    If imgFile = 0 Then Exit Sub      ' delete nullptr
    IMAGEFILE_delete imgFile
End Sub

' IMAGEFILE::~IMAGEFILE()
Private Sub IMAGEFILE_destructor(ByVal imgFile As Long)
    Dim it As Long, i As Long
    it = files_size
    For i = 0 To files_size - 1
        If files(i) = imgFile Then
            it = i
            Exit For
        End If
    Next
    If Not (it <> files_size) Then ASSERT_FAILURE "it != files.end()", "lib/ivis_opengl/bitimage.cpp:327", "~IMAGEFILE", "Calling iV_FreeImageFile for IMAGEFILE that wasn't loaded??"
    If it <> files_size Then
        For i = it To files_size - 2
            files(i) = files(i + 1)
        Next
        files_size = files_size - 1
    End If
End Sub

' AtlasImageDef* IMAGEFILE::find(WzString const &filename): an AtlasImageDef handle
Public Function IMAGEFILE_find(ByVal imgFile As Long, ByVal filename As String) As Long
    If imageFiles(imgFile).imageNamesMap.Exists(filename) Then
        IMAGEFILE_find = AtlasImageDef_handle(imgFile, imageFiles(imgFile).imageNamesMap(filename))
        Exit Function

    End If
'    debug(LOG_ERROR, "%s not found in image list!", filename.toUtf8().c_str());
    IMAGEFILE_find = 0 ' Error, image not found.
End Function

Public Sub iV_ImageFileShutdown()
    Dim i As Long, file As Long
    If Not (files_size = 0) Then ASSERT_FAILURE "files.empty()", "lib/ivis_opengl/bitimage.cpp:348", "iV_ImageFileShutdown", "%zu IMAGEFILEs were not properly unloaded", files_size
    ' for (auto file : files) delete file; (each delete erases it from files: the loop runs over a copy)
    Do While files_size > 0
        file = files(0)
        IMAGEFILE_delete file
    Loop
    files_size = 0
End Sub
