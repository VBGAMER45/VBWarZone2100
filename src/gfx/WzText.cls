VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/textdraw.h, textdraw.cpp: class WzText, a text drawn once
' into a texture (textdraw_gdi's TextShaper) and rendered from it. `New
' WzText` is WzText(); WzText(text, fontID) is textdraw_gdi's WzText_new
' (GfxApi). The C++ move (operator=(WzText&&), WzText(WzText&&)) is moveFrom;
' the destructor is the last reference going. A Vector2f position is (x, y),
' a WzRect (left, top, width, height); a colour is a PIELIGHT (Long); a font
' is an iV_fonts as a Long (GfxApi's WzText_new returns this class, so its
' public methods use none of GfxApi's enums: "Circular dependencies").
Option Explicit

Private mText As String
Private texture As gfx_api_texture              ' = nullptr
Private mPtsAboveBase As Long                   ' = 0
Private mPtsBelowBase As Long                   ' = 0
Private mPtsLineSize As Long                    ' = 0
Private offsets As Vector2i                     ' = Vector2i(0, 0)
Private dimensions As Vector2i                  ' = Vector2i(0, 0)
Private mRenderingHorizScaleFactor As Single    ' = 0.f
Private mRenderingVertScaleFactor As Single     ' = 0.f
Private mFontID As Long                         ' iV_fonts = font_count
Private layoutMetrics As Vector2i               ' = Vector2i(0, 0)

Private Sub Class_Initialize()
    mFontID = font_count
End Sub

Public Function width() As Long
    updateCacheIfNecessary
    width = width_pixelsToPoints(layoutMetrics.x)
End Function

Public Function height() As Long
    updateCacheIfNecessary
    height = height_pixelsToPoints(layoutMetrics.y)
End Function

' (in points)
Public Function aboveBase() As Long
    updateCacheIfNecessary
    aboveBase = mPtsAboveBase
End Function

' (in points)
Public Function belowBase() As Long
    updateCacheIfNecessary
    belowBase = mPtsBelowBase
End Function

' (in points)
Public Function lineSize() As Long
    updateCacheIfNecessary
    lineSize = mPtsLineSize
End Function

Public Sub setText(ByVal text As String, ByVal fontID As Long)
    If mText = text And fontID = mFontID Then
        Exit Sub ' cached
    End If
    drawAndCacheText text, fontID
End Sub

Public Function getText() As String
    getText = mText
End Function

Public Function getFontID() As Long
    getFontID = mFontID
End Function

Private Sub drawAndCacheText(ByVal string_ As String, ByVal fontID As Long)
    Dim face As Long, drawResult As DrawTextResult
    mFontID = fontID
    mText = string_
    mRenderingHorizScaleFactor = iV_GetHorizScaleFactor()
    mRenderingVertScaleFactor = iV_GetVertScaleFactor()

    face = getFTFace(fontID, HB_SCRIPT_COMMON)

    ' TODO: Better handling here of multiple font faces?
    mPtsAboveBase = metricsHeight_PixelsToPoints(-FTFace_ascender(face))
    mPtsLineSize = metricsHeight_PixelsToPoints(FTFace_ascender(face) - FTFace_descender(face))
    mPtsBelowBase = metricsHeight_PixelsToPoints(FTFace_descender(face))

    drawResult = TextShaper_drawText(string_, fontID)
    If Not drawResult.text.bitmap Is Nothing Then
        dimensions.x = drawResult.text.bitmap.width()
        dimensions.y = drawResult.text.bitmap.height()
    Else
        dimensions.x = 0
        dimensions.y = 0
    End If
    offsets.x = drawResult.text.offset_x
    offsets.y = drawResult.text.offset_y
    layoutMetrics.x = drawResult.layoutMetrics.width
    layoutMetrics.y = drawResult.layoutMetrics.height

    Set texture = Nothing

    If dimensions.x > 0 And dimensions.y > 0 Then
        Set texture = gfx_api_context_createTextureForCompatibleImageUploads(1, drawResult.text.bitmap, "text::" & string_)
        texture.upload 0, drawResult.text.bitmap
    End If
End Sub

Private Sub redrawAndCacheText()
    drawAndCacheText mText, mFontID
End Sub

' operator=(WzText&& other), WzText(WzText&& other)
Public Sub moveFrom(ByVal other As WzText)
    If other Is Me Then Exit Sub
    ' Free the existing texture, if any; get the other data; reset other's texture
    other.moveTo texture, mFontID, mText, mPtsAboveBase, mPtsBelowBase, mPtsLineSize, offsets.x, offsets.y, dimensions.x, dimensions.y, mRenderingHorizScaleFactor, mRenderingVertScaleFactor, layoutMetrics.x, layoutMetrics.y
End Sub

' (VB) moveFrom's other half: hands over this text's data, keeping none of its texture
Friend Sub moveTo(texture_ As gfx_api_texture, fontID As Long, text As String, ptsAboveBase As Long, ptsBelowBase As Long, ptsLineSize As Long, offsets_x As Long, offsets_y As Long, dimensions_x As Long, dimensions_y As Long, renderingHorizScaleFactor As Single, renderingVertScaleFactor As Single, layoutMetrics_x As Long, layoutMetrics_y As Long)
    Set texture_ = texture
    fontID = mFontID
    text = mText
    mText = ""                          ' std::move(other.mText)
    ptsAboveBase = mPtsAboveBase
    ptsBelowBase = mPtsBelowBase
    ptsLineSize = mPtsLineSize
    offsets_x = offsets.x
    offsets_y = offsets.y
    dimensions_x = dimensions.x
    dimensions_y = dimensions.y
    renderingHorizScaleFactor = mRenderingHorizScaleFactor
    renderingVertScaleFactor = mRenderingVertScaleFactor
    layoutMetrics_x = layoutMetrics.x
    layoutMetrics_y = layoutMetrics.y
    Set texture = Nothing
End Sub

Private Sub updateCacheIfNecessary()
    If Len(mText) = 0 Then
        Exit Sub ' string is empty (or hasn't yet been set), thus changes have no effect
    End If
    If mRenderingHorizScaleFactor <> iV_GetHorizScaleFactor() Or mRenderingVertScaleFactor <> iV_GetVertScaleFactor() Then
        ' The text rendering subsystem's scale factor has changed, so the rendered (cached) text must be re-rendered.
        redrawAndCacheText
        ' debug(LOG_WZ, "Redrawing / re-calculating WzText text - scale factor has changed.");
    End If
End Sub

' position: Vector2f (x, y); screenClippingRect: WzRect (left, top, width, height)
Public Sub renderClipped(ByVal x As Single, ByVal y As Single, ByVal colour As Long, ByVal clipLeft As Long, ByVal clipTop As Long, ByVal clipWidth As Long, ByVal clipHeight As Long, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1)
    Dim visualOrigin As Vector2f, logicalDisplayWidth As Long, logicalDisplayHeight As Long
    Dim drawLeft As Long, drawTop As Long, drawRight As Long, drawBottom As Long
    Dim clipRight As Long, clipBottom As Long
    Dim cl As Long, ct As Long, cr As Long, cb As Long, clippingWidth As Long, clippingHeight As Long
    Dim clip_x As Single, clip_y As Single, clip_width As Single, clip_height As Single ' WzClippingRectF clippingRectInPixels
    Dim logicalDrawSize As Vector2f, position As Vector2f, offs As Vector2f
    updateCacheIfNecessary

    If texture Is Nothing Then
        ' A texture will not always be created. (For example, if the rendered text is empty.)
        ' No need to render if there's nothing to render.
        Exit Sub
    End If

    visualOrigin.x = x
    visualOrigin.y = y + mPtsAboveBase
    logicalDisplayWidth = Fix(dimensions.x / mRenderingHorizScaleFactor)
    logicalDisplayHeight = Fix(dimensions.y / mRenderingVertScaleFactor)
    If maxWidth > 0 Then
        If maxWidth < logicalDisplayWidth Then logicalDisplayWidth = maxWidth
    End If
    ' WzRect screenDrawRect(visualOrigin.x, visualOrigin.y, logicalDisplayWidth, logicalDisplayHeight); // screen coordinates
    drawLeft = Fix(visualOrigin.x)
    drawTop = Fix(visualOrigin.y)
    drawRight = drawLeft + logicalDisplayWidth
    drawBottom = drawTop + logicalDisplayHeight
    ' WzRect clippingRect = screenDrawRect.intersectionWith(screenClippingRect.setWidth(screenClippingRect.width() - 1).setHeight(screenClippingRect.height() - 1));
    clipRight = clipLeft + (clipWidth - 1)
    clipBottom = clipTop + (clipHeight - 1)
    If drawLeft > clipLeft Then cl = drawLeft Else cl = clipLeft
    If drawTop > clipTop Then ct = drawTop Else ct = clipTop
    If drawRight < clipRight Then cr = drawRight Else cr = clipRight
    If drawBottom < clipBottom Then cb = drawBottom Else cb = clipBottom
    If cr - cl <= 0 Or cb - ct <= 0 Then
        Exit Sub
    End If
    ' clippingRect.translateBy(-visualOrigin.x, -visualOrigin.y); // translate to 0,0 origin
    cl = cl - Fix(visualOrigin.x)
    cr = cr - Fix(visualOrigin.x)
    ct = ct - Fix(visualOrigin.y)
    cb = cb - Fix(visualOrigin.y)
    clippingWidth = cr - cl
    clippingHeight = cb - ct

    clip_x = cl * mRenderingHorizScaleFactor
    clip_y = ct * mRenderingVertScaleFactor
    clip_width = clippingWidth * mRenderingHorizScaleFactor
    clip_height = clippingHeight * mRenderingVertScaleFactor

    If logicalDisplayWidth < clippingWidth Then logicalDrawSize.x = logicalDisplayWidth Else logicalDrawSize.x = clippingWidth
    If logicalDisplayHeight < clippingHeight Then logicalDrawSize.y = logicalDisplayHeight Else logicalDrawSize.y = clippingHeight

    position.x = x + cl
    position.y = y + ct
    offs.x = offsets.x / mRenderingHorizScaleFactor
    offs.y = offsets.y / mRenderingVertScaleFactor
    iV_DrawImageTextClipped texture, dimensions, position, offs, logicalDrawSize, 0!, colour, clip_x, clip_y, clip_width, clip_height
End Sub

' render(Vector2f position, ...) and render(float x, float y, ...)
Public Sub render(ByVal x As Single, ByVal y As Single, ByVal colour As Long, Optional ByVal rotation_ As Single = 0!, Optional ByVal maxWidth As Long = -1, Optional ByVal maxHeight As Long = -1)
    Dim position As Vector2f, offs As Vector2f, size As Vector2f
    Dim clip_x As Single, clip_y As Single, clip_width As Single, clip_height As Single ' WzClippingRectF clippingRectInPixels
    updateCacheIfNecessary

    If texture Is Nothing Then
        ' A texture will not always be created. (For example, if the rendered text is empty.)
        ' No need to render if there's nothing to render.
        Exit Sub
    End If

    If rotation_ <> 0! Then
        rotation_ = 180! - rotation_
    End If

    position.x = x
    position.y = y
    offs.x = offsets.x / mRenderingHorizScaleFactor
    offs.y = offsets.y / mRenderingVertScaleFactor
    If maxWidth <= 0 And maxHeight <= 0 Then
        size.x = dimensions.x / mRenderingHorizScaleFactor
        size.y = dimensions.y / mRenderingVertScaleFactor
        iV_DrawImageText texture, position, offs, size, rotation_, colour
    Else
        clip_x = 0!
        clip_y = 0!
        If maxWidth > 0 Then clip_width = CSng(maxWidth) * mRenderingHorizScaleFactor Else clip_width = dimensions.x
        If maxHeight > 0 Then clip_height = CSng(maxHeight) * mRenderingVertScaleFactor Else clip_height = dimensions.y
        If maxWidth > 0 Then size.x = maxWidth Else size.x = dimensions.x / mRenderingHorizScaleFactor
        If maxHeight > 0 Then size.y = maxHeight Else size.y = dimensions.y / mRenderingVertScaleFactor
        iV_DrawImageTextClipped texture, dimensions, position, offs, size, rotation_, colour, clip_x, clip_y, clip_width, clip_height
    End If
End Sub

Public Sub renderOutlined(ByVal x As Long, ByVal y As Long, ByVal colour As Long, ByVal outlineColour As Long)
    Dim i As Long, j As Long
    For i = -1 To 1
        For j = -1 To 1
            render x + i, y + j, outlineColour
        Next
    Next
    render x, y, colour
End Sub
