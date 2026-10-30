Attribute VB_Name = "intimage_c"
' src/intimage.cpp, intimage.h: the interface's window frames
' (RenderWindowFrame), imageInitBitmaps and class IntListTabWidget (its field:
' IntListTabWidget.cls). IMAGEFILE *IntImages is still stubs_display's Long
' (design: until cluster G); imageInitBitmaps sets it. The frame tables
' FrameNormal / FrameRadar are filled on first use (VB has no initialised
' UDTs). RenderWindowFrame's modelViewProjection is a mat4's VarPtr (0 =
' defaultProjectionMatrix()), as GfxApi's.
' (VB) widget class IntListTabWidget : ListTabWidget
Option Explicit

Public Const FILLRED As Long = 16
Public Const FILLGREEN As Long = 16
Public Const FILLBLUE As Long = 128
Public Const FILLTRANS As Long = 128

Public Enum FRAMETYPE
    FRAME_NORMAL
    FRAME_RADAR
End Enum

' struct TABDEF
Public Type TABDEF
    MajorUp As Integer                  ' < Index of image to use for tab not pressed.
    MajorDown As Integer                ' < Index of image to use for tab pressed.
    MajorHilight As Integer             ' < Index of image to use for tab hilighted by mouse.
    MajorSelected As Integer            ' < Index of image to use for tab selected (same as pressed).
End Type

Private Const INCEND As Long = 0

Private Enum FR_DRAWTYPE
    FR_SOLID                            ' < Bitmap drawn solid.
    FR_KEYED                            ' < Bitmap drawn with colour 0 transparent.
End Enum

Private Enum FRAMERECTTYPE
    FR_IGNORE                           ' < Fill rect is ignored.
    FR_FRAME                            ' < Fill rect drawn relative to frame.
    FR_LEFT                             ' < Fill rect drawn relative to left of frame.
    FR_RIGHT                            ' < Fill rect drawn relative to right of frame.
    FR_TOP                              ' < Fill rect drawn relative to top of frame.
    FR_BOTTOM                           ' < Fill rect drawn relative to bottom of frame.
End Enum

Private Type FRAMERECT
    typ As FRAMERECTTYPE                ' < One of the FR_... values.
    TLXOffset As Long                   ' < Offsets for the rect fill.
    TLYOffset As Long
    BRXOffset As Long
    BRYOffset As Long
    ColourIndex As Long                 ' < Hackish index into the WZCOLOR palette
End Type

' Frame definition structure.
Private Type IMAGEFRAME
    OffsetX0 As Integer                 ' < Offset top left of frame.
    OffsetY0 As Integer
    OffsetX1 As Integer                 ' < Offset bottom right of frame.
    OffsetY1 As Integer
    TopLeft As Integer                  ' < Image indecies for the corners ( -1 = don't draw).
    TopRight As Integer
    BottomLeft As Integer
    BottomRight As Integer
    TopEdge As Integer                  ' < Image indecies for the edges ( -1 = don't draw). Type ie FR_SOLID or FR_KEYED.
    TopType As Integer
    RightEdge As Integer
    RightType As Integer
    BottomEdge As Integer
    BottomType As Integer
    LeftEdge As Integer
    LeftType As Integer
    FRect(4) As FRAMERECT               ' < Fill rectangles.
End Type

' Form frame definition for normal frames.
Private FrameNormal As IMAGEFRAME
' Form frame definition for radar frames.
Private FrameRadar As IMAGEFRAME
Private framesFilled As Boolean

Private Sub setFrameRect(r As FRAMERECT, ByVal typ As FRAMERECTTYPE, ByVal TLXOffset As Long, ByVal TLYOffset As Long, ByVal BRXOffset As Long, ByVal BRYOffset As Long, ByVal ColourIndex As Long)
    r.typ = typ
    r.TLXOffset = TLXOffset
    r.TLYOffset = TLYOffset
    r.BRXOffset = BRXOffset
    r.BRYOffset = BRYOffset
    r.ColourIndex = ColourIndex
End Sub

Private Sub setFrame(f As IMAGEFRAME)
    Dim i As Long
    f.OffsetX0 = 0: f.OffsetY0 = 0: f.OffsetX1 = 0: f.OffsetY1 = 0
    f.TopLeft = IMAGE_FRAME_C0
    f.TopRight = IMAGE_FRAME_C1
    f.BottomLeft = IMAGE_FRAME_C3
    f.BottomRight = IMAGE_FRAME_C2
    f.TopEdge = IMAGE_FRAME_HT: f.TopType = FR_SOLID
    f.RightEdge = IMAGE_FRAME_VR: f.RightType = FR_SOLID
    f.BottomEdge = IMAGE_FRAME_HB: f.BottomType = FR_SOLID
    f.LeftEdge = IMAGE_FRAME_VL: f.LeftType = FR_SOLID
    For i = 0 To 4
        setFrameRect f.FRect(i), FR_IGNORE, 0, 0, 0, 0, 0
    Next i
End Sub

Private Sub fillFrames()
    setFrame FrameNormal
    setFrameRect FrameNormal.FRect(0), FR_FRAME, 0, 1, 0, -1, 33
    setFrame FrameRadar
    framesFilled = True
End Sub

' Read bitmaps used by the interface.
'
Public Function imageInitBitmaps() As Boolean
    Dim data As Variant, errorMessage As String, mods() As WzMods_LoadedMod
    data = resGetData("IMG", "intfac.img")
    If IsEmpty(data) Then IntImages = 0 Else IntImages = data
    If IntImages = 0 Then
        errorMessage = va("Unable to load: %s.", "intfac.img")
        If getLoadedMods(mods) > 0 Then
            errorMessage = errorMessage & " "
            errorMessage = errorMessage & "Please remove all incompatible mods."
        End If
        debug_ 133, LOG_FATAL, "imageInitBitmaps", "%s", errorMessage
        imageInitBitmaps = False
        Exit Function
    End If

    imageInitBitmaps = True
End Function

' Render a window frame.
'
Public Sub RenderWindowFrame(ByVal frame As FRAMETYPE, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal pModelViewProjectionMatrix As Long = 0)
    Dim WTopRight As Long, WTopLeft As Long, WBottomRight As Long, WBottomLeft As Long
    Dim HTopRight As Long, HTopLeft As Long, HBottomRight As Long, HBottomLeft As Long
    Dim RectI As Long, Frame_ As IMAGEFRAME, imageDrawBatch As BatchedImageDrawRequests
    If Width = 0 Or Height = 0 Then
        Exit Sub
    End If

    If Not framesFilled Then fillFrames
    If frame = FRAME_NORMAL Then Frame_ = FrameNormal Else Frame_ = FrameRadar

    x = x + Frame_.OffsetX0
    y = y + Frame_.OffsetY0
    Width = Width - (Frame_.OffsetX1 + Frame_.OffsetX0)
    Height = Height - (Frame_.OffsetY1 + Frame_.OffsetY0)

    For RectI = 0 To 4
        With Frame_.FRect(RectI)
            Select Case .typ
            Case FR_FRAME
                iV_TransBoxFill x + .TLXOffset, y + .TLYOffset, x + Width - INCEND + .BRXOffset, y + Height - INCEND + .BRYOffset
            Case FR_LEFT
                iV_TransBoxFill x + .TLXOffset, y + .TLYOffset, x + .BRXOffset, y + Height - INCEND + .BRYOffset
            Case FR_RIGHT
                iV_TransBoxFill x + Width - INCEND + .TLXOffset, y + .TLYOffset, x + Width - INCEND + .BRXOffset, y + Height - INCEND + .BRYOffset
            Case FR_TOP
                iV_TransBoxFill x + .TLXOffset, y + .TLYOffset, x + Width - INCEND + .BRXOffset, y + .BRYOffset
            Case FR_BOTTOM
                iV_TransBoxFill x + .TLXOffset, y + Height - INCEND + .TLYOffset, x + Width - INCEND + .BRXOffset, y + Height - INCEND + .BRYOffset
            Case FR_IGNORE
                ' ignored
            End Select
        End With
    Next RectI

    Set imageDrawBatch = BatchedImageDrawRequests_new(True) ' defer drawing

    If Frame_.TopLeft >= 0 Then
        WTopLeft = iV_GetImageWidth(IntImages, Frame_.TopLeft)
        HTopLeft = iV_GetImageHeight(IntImages, Frame_.TopLeft)
        iV_DrawImage IntImages, Frame_.TopLeft, x, y, pModelViewProjectionMatrix, imageDrawBatch
    End If

    If Frame_.TopRight >= 0 Then
        WTopRight = iV_GetImageWidth(IntImages, Frame_.TopRight)
        HTopRight = iV_GetImageHeight(IntImages, Frame_.TopRight)
        iV_DrawImage IntImages, Frame_.TopRight, x + Width - WTopRight, y, pModelViewProjectionMatrix, imageDrawBatch
    End If

    If Frame_.BottomRight >= 0 Then
        WBottomRight = iV_GetImageWidth(IntImages, Frame_.BottomRight)
        HBottomRight = iV_GetImageHeight(IntImages, Frame_.BottomRight)
        iV_DrawImage IntImages, Frame_.BottomRight, x + Width - WBottomRight, y + Height - HBottomRight, pModelViewProjectionMatrix, imageDrawBatch
    End If

    If Frame_.BottomLeft >= 0 Then
        WBottomLeft = iV_GetImageWidth(IntImages, Frame_.BottomLeft)
        HBottomLeft = iV_GetImageHeight(IntImages, Frame_.BottomLeft)
        iV_DrawImage IntImages, Frame_.BottomLeft, x, y + Height - HBottomLeft, pModelViewProjectionMatrix, imageDrawBatch
    End If

    If Frame_.TopEdge >= 0 Then
        iV_DrawImageRepeatX IntImages, Frame_.TopEdge, x + iV_GetImageWidth(IntImages, Frame_.TopLeft), y, Width - WTopLeft - WTopRight, pModelViewProjectionMatrix, False, imageDrawBatch
    End If

    If Frame_.BottomEdge >= 0 Then
        iV_DrawImageRepeatX IntImages, Frame_.BottomEdge, x + WBottomLeft, y + Height - iV_GetImageHeight(IntImages, Frame_.BottomEdge), Width - WBottomLeft - WBottomRight, pModelViewProjectionMatrix, False, imageDrawBatch
    End If

    If Frame_.LeftEdge >= 0 Then
        iV_DrawImageRepeatY IntImages, Frame_.LeftEdge, x, y + HTopLeft, Height - HTopLeft - HBottomLeft, pModelViewProjectionMatrix, imageDrawBatch
    End If

    If Frame_.RightEdge >= 0 Then
        iV_DrawImageRepeatY IntImages, Frame_.RightEdge, x + Width - iV_GetImageWidth(IntImages, Frame_.RightEdge), y + HTopRight, Height - HTopRight - HBottomRight, pModelViewProjectionMatrix, imageDrawBatch
    End If

    imageDrawBatch.draw True
End Sub

' MARK: - IntListTabWidget

' IntListTabWidget(): ListTabWidget() {} (and the member's initialiser)
Public Sub IntListTabWidget_construct(ByVal self As WIDGET)
    ListTabWidget_construct self
    self.kind = K_IntListTabWidget ' (VB) virtual calls from here on reach IntListTabWidget's overrides
    IntListTabWidget_(self).tabAlignment_ = TabAlignment_LeftAligned
End Sub

Public Function IntListTabWidget_make(Optional ByVal tabAlign As TabAlignment = TabAlignment_LeftAligned) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_IntListTabWidget)
    IntListTabWidget_construct widget
    IntListTabWidget_(widget).tabAlignment_ = tabAlign
    IntListTabWidget_initialize widget ' (VB) widget->initialize(): its class is IntListTabWidget
    Set IntListTabWidget_make = widget
End Function

' (VB) override of ListTabWidget::initialize
Public Sub IntListTabWidget_initialize(ByVal self As WIDGET)
    Dim tabAlign As TabAlignment, nul As AtlasImage
    tabAlign = IntListTabWidget_(self).tabAlignment_
    ListTabWidget_initialize self
    TabSelectionWidget_setHeight ListTabWidget_tabWidget(self), 15
    TabSelectionWidget_addStyle ListTabWidget_tabWidget(self), TabSelectionStyle_(AtlasImage_(IntImages, IMAGE_TAB1), AtlasImage_(IntImages, IMAGE_TAB1DOWN), AtlasImage_(IntImages, IMAGE_TABHILIGHT), nul, nul, nul, nul, nul, nul, 2, tabAlign)
    TabSelectionWidget_addStyle ListTabWidget_tabWidget(self), TabSelectionStyle_(AtlasImage_(IntImages, IMAGE_TAB1_SM), AtlasImage_(IntImages, IMAGE_TAB1DOWN_SM), AtlasImage_(IntImages, IMAGE_TABHILIGHT_SM), nul, nul, nul, nul, nul, nul, 2, tabAlign)
    TabSelectionWidget_addStyle ListTabWidget_tabWidget(self), TabSelectionStyle_(AtlasImage_(IntImages, IMAGE_TAB1_SM), AtlasImage_(IntImages, IMAGE_TAB1DOWN_SM), AtlasImage_(IntImages, IMAGE_TABHILIGHT_SM), AtlasImage_(IntImages, IMAGE_LFTTAB), AtlasImage_(IntImages, IMAGE_LFTTABD), AtlasImage_(IntImages, IMAGE_LFTTABD), AtlasImage_(IntImages, IMAGE_RGTTAB), AtlasImage_(IntImages, IMAGE_RGTTABD), AtlasImage_(IntImages, IMAGE_RGTTABD), 2, tabAlign)
End Sub
