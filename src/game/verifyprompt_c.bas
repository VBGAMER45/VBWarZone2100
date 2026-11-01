Attribute VB_Name = "verifyprompt_c"
' lib/widget/verifyprompt.cpp, verifyprompt.h, whole (namespace WzVerifyPrompt): the captcha and slide prompts a
' lobby can ask a joining player to resolve (netlobby_c's GetLobbyErrorResolveWidget). The image providers
' (std::shared_ptr<IvImageProviderInterface>) are netlobby's B64ImageData, the interface's one implementation;
' VerifyPromptData's GetType is the data's type (CaptchaVerifyPromptData, SlideVerifyPromptData). The widget
' classes' fields are their data classes; the std::weak_ptr captures are widgweak_vb serials. The overrides that
' read a texture return early while it is Nothing: 4.7.0 keeps a prompt whose image failed to load (initialize's
' false is ignored by make) and then dereferences the null texture. Translations (_()) are not ported.
' (VB) widget class PromptActionButtonWidget : W_BUTTON
' (VB) widget class CaptchaPromptWidget : WIDGET
' (VB) widget class VerifySlideWidget : W_SLIDER
' (VB) widget class SlidePromptWidget : WIDGET
' (VB) widget class CaptchaWrapperWidget : WIDGET
Option Explicit

' class CaptchaVerifyPromptData : public VerifyPromptData
Public Type CaptchaVerifyPromptData
    image As B64ImageData                       ' std::shared_ptr<IvImageProviderInterface> image;
End Type

' class SlideVerifyPromptData : public VerifyPromptData
Public Type SlideVerifyPromptData
    image As B64ImageData
    tile As B64ImageData
    tile_width As Long                          ' uint32_t = 0
    tile_height As Long
    tile_x As Long
    tile_y As Long
End Type

' PromptActionButtonWidget's const int InternalPadding = 12;
Private Const PromptAction_InternalPadding As Long = 12
' CaptchaPromptWidget's and SlidePromptWidget's consts
Private Const textFontId As Long = font_regular_bold
Private Const innerPaddingY As Long = 10
Private Const minimumInputAreaWidth As Long = 200
Private Const sliderHeight As Long = 28
' CaptchaWrapperWidget's consts
Private Const Wrapper_innerPaddingX As Long = 20
Private Const Wrapper_innerPaddingY As Long = 20

Public Const FN_verifyprompt_captchaOnReturn As Long = 17240        ' [weakParent](W_EDITBOX& widg)
Public Const FN_verifyprompt_captchaSubmitTask As Long = 17241      ' [str, weakParent]()
Public Const FN_verifyprompt_captchaOkClick As Long = 17242         ' [weakParent](W_BUTTON&)
Public Const FN_verifyprompt_slideOnChange As Long = 17243          ' [weakParent](W_SLIDER &)
Public Const FN_verifyprompt_slideOnReleased As Long = 17244        ' [weakParent](VerifySlideWidget &)

' MARK: - Helpers

' (VB) std::max<int>, std::min<int>
Private Function jmax(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then jmax = a Else jmax = b
End Function

Private Function jmin(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then jmin = a Else jmin = b
End Function

' (VB) std::ceil
Private Function ceil_(ByVal v As Double) As Double
    ceil_ = -Int(-v)
End Function

' (VB) std::round (halves away from zero)
Private Function round_(ByVal v As Double) As Double
    If v < 0 Then round_ = -Int(-v + 0.5) Else round_ = Int(v + 0.5)
End Function

Private Function getTextureDimensionsScaleRatioForMax(ByVal dimensions As gfx_api_texture, ByVal maxWidth As Long, ByVal maxHeight As Long) As Double
    Dim dw As Long, dh As Long, scalingRatio As Double, widthRatio As Double, heightRatio As Double
    dimensions.getDimensions dw, dh
    ' (size_t maxWidth, maxHeight: a negative height wraps to a huge size_t)
    If (maxWidth = 0 Or UnsignedLE(dw, maxWidth)) And (maxHeight = 0 Or UnsignedLE(dh, maxHeight)) Then
        getTextureDimensionsScaleRatioForMax = 1#
        Exit Function
    End If

    widthRatio = SizeT(maxWidth) / CDbl(dw)
    heightRatio = SizeT(maxHeight) / CDbl(dh)
    If maxWidth <> 0 And maxHeight <> 0 Then
        If widthRatio < heightRatio Then scalingRatio = widthRatio Else scalingRatio = heightRatio
    Else
        If maxWidth <> 0 Then scalingRatio = widthRatio Else scalingRatio = heightRatio
    End If

    getTextureDimensionsScaleRatioForMax = scalingRatio
End Function

' (VB) the int passed as size_t (a negative int converts to 2^64 + it)
Private Function SizeT(ByVal v As Long) As Double
    If v < 0 Then SizeT = 18446744073709551616# + v Else SizeT = v
End Function

' (VB) dimensions.width <= (size_t)max
Private Function UnsignedLE(ByVal dimValue As Long, ByVal maxValue As Long) As Boolean
    UnsignedLE = CDbl(dimValue) <= SizeT(maxValue)
End Function

Private Function textureWidth(ByVal t As gfx_api_texture) As Long
    Dim w As Long, h As Long
    t.getDimensions w, h
    textureWidth = w
End Function

Private Function textureHeight(ByVal t As gfx_api_texture) As Long
    Dim w As Long, h As Long
    t.getDimensions w, h
    textureHeight = h
End Function

' MARK: - PromptActionButtonWidget

' PromptActionButtonWidget() {}
Public Sub PromptActionButtonWidget_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_PromptActionButtonWidget ' (VB) virtual calls from here on reach PromptActionButtonWidget's overrides
End Sub

Private Sub PromptActionButtonWidget_initialize(ByVal self As WIDGET, ByVal text As String)
    Dim b As W_BUTTON, minButtonWidthForText As Long
    self.setString text
    Set b = W_BUTTON_(self)
    b.FontID = font_regular_bold
    minButtonWidthForText = iV_GetTextWidth(b.pText, b.FontID)
    self.setGeometry 0, 0, minButtonWidthForText + PromptAction_InternalPadding, iV_GetTextLineSize(b.FontID) + PromptAction_InternalPadding
End Sub

' static std::shared_ptr<PromptActionButtonWidget> make(const WzString& text)
Public Function PromptActionButtonWidget_make(ByVal text As String) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_PromptActionButtonWidget)
    PromptActionButtonWidget_construct widget
    PromptActionButtonWidget_initialize widget, text
    Set PromptActionButtonWidget_make = widget
End Function

Public Sub PromptActionButtonWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean, isDown As Boolean, isDisabled As Boolean
    Dim isHighlight As Boolean, colour As Long, fillClr As Long, d As PromptActionButtonWidget, b As W_BUTTON
    Dim fw As Long, fx As Long, fy As Long
    Set b = W_BUTTON_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(b.pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = (Not isDisabled) And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    If isDisabled Then                  ' unavailable
        colour = psPalette(WZCOL_TEXT_DARK)
    Else                                ' available
        If isHighlight Then             ' hilight
            colour = psPalette(WZCOL_TEXT_BRIGHT)
        Else                            ' don't highlight
            colour = psPalette(WZCOL_TEXT_MEDIUM)
        End If
    End If

    If isHighlight Or isDown Then
        fillClr = pal_RGBA(255, 255, 255, 30)
        pie_UniTransBoxFill x0, y0, x1, y1, fillClr
    End If
    iV_Box x0, y0, x1, y1, colour

    If haveText Then
        Set d = PromptActionButtonWidget_(self)
        d.wzText.setText b.pText, b.FontID
        fw = d.wzText.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
        d.wzText.render fx, fy, colour
    End If
End Sub

' MARK: - CaptchaPromptWidget

' CaptchaPromptWidget()
Public Sub CaptchaPromptWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_CaptchaPromptWidget   ' (VB) virtual calls from here on reach CaptchaPromptWidget's overrides
    CaptchaPromptWidget_(self).weakSerial = WeakWidget_register(self)
End Sub

' ~CaptchaPromptWidget()
Public Sub CaptchaPromptWidget_destruct(ByVal self As WIDGET)
    Dim d As CaptchaPromptWidget
    Set d = CaptchaPromptWidget_(self)
    If Not d.pImageTexture Is Nothing Then
        Set d.pImageTexture = Nothing   ' delete pImageTexture
    End If
    WeakWidget_unregister d.weakSerial
End Sub

Public Function CaptchaPromptWidget_make(promptData As CaptchaVerifyPromptData, ByVal responseCallback As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_CaptchaPromptWidget)
    CaptchaPromptWidget_construct widget
    CaptchaPromptWidget_initialize widget, promptData, responseCallback
    Set CaptchaPromptWidget_make = widget
End Function

Private Function CaptchaPromptWidget_initialize(ByVal self As WIDGET, promptData As CaptchaVerifyPromptData, ByVal responseCallback As WzFunc) As Boolean
    Dim d As CaptchaPromptWidget, image As iV_Image, weakParent As Long
    Set d = CaptchaPromptWidget_(self)
    Set d.promptData_image = promptData.image
    Set d.responseCallback = responseCallback

    Set d.promptText = W_LABEL_new()
    W_LABEL_setFont d.promptText, textFontId, psPalette(WZCOL_FORM_LIGHT)
    d.promptText.setString "Please enter the text below"
    W_LABEL_setCanTruncate d.promptText, True
    d.promptText.setTransparentToClicks True
    self.attach d.promptText

    If Not (Not d.promptData_image Is Nothing) Then ASSERT_FAILURE "promptData.image != nullptr", "lib/widget/verifyprompt.cpp:216", "initialize", "No image to load": Exit Function
    If Not (d.promptData_image.GetImage(image)) Then ASSERT_FAILURE "image.has_value()", "lib/widget/verifyprompt.cpp:218", "initialize", "Failed to load image": Exit Function

    Set d.pImageTexture = gfx_api_context_loadTextureFromUncompressedImage(image, texture_type_user_interface, "slide_image")
    If Not (Not d.pImageTexture Is Nothing) Then ASSERT_FAILURE "pImageTexture != nullptr", "lib/widget/verifyprompt.cpp:221", "initialize", "Failed to load image texture": Exit Function

    Set d.inputBox = W_EDITBOX_new()
    self.attach d.inputBox
    d.inputBox.setGeometry 0, 0, 280, 24
    W_EDITBOX_setBoxColours d.inputBox, psPalette(WZCOL_WHITE), psPalette(WZCOL_WHITE), psPalette(WZCOL_MENU_BACKGROUND)
    W_EDITBOX_setMaxStringSize d.inputBox, 16
    W_EDITBOX_setPlaceholder d.inputBox, "Enter the text"

    weakParent = d.weakSerial
    W_EDITBOX_setOnReturnHandler d.inputBox, WzFunc_new(FN_verifyprompt_captchaOnReturn, weakParent)

    ' Add OK button
    Set d.okayButton = PromptActionButtonWidget_make("OK")
    self.attach d.okayButton
    W_BUTTON_addOnClickHandler d.okayButton, WzFunc_new(FN_verifyprompt_captchaOkClick, weakParent)

    CaptchaPromptWidget_initialize = True
End Function

Public Sub CaptchaPromptWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As CaptchaPromptWidget, x0 As Long, y0 As Long, scaledW As Single, scaledH As Single
    Dim imageAreaXPadding As Long, imageBoxX0 As Long, imageBoxY0 As Long
    Set d = CaptchaPromptWidget_(self)
    If d.pImageTexture Is Nothing Then Exit Sub        ' (VB) the prompt's image failed to load
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    scaledW = textureWidth(d.pImageTexture) * d.imageScalingRatio
    scaledH = textureHeight(d.pImageTexture) * d.imageScalingRatio

    imageAreaXPadding = (self.width() - CLng(ceil_(scaledW))) \ 2
    imageBoxX0 = x0 + imageAreaXPadding
    imageBoxY0 = y0 + d.promptText.y() + d.promptText.height() + innerPaddingY

    ' Draw white background behind the pImageTexture
    pie_UniTransBoxFill imageBoxX0, imageBoxY0, imageBoxX0 + scaledW, imageBoxY0 + scaledH, psPalette(WZCOL_WHITE)

    ' Draw the pImageTexture
    iV_DrawImageAnisotropic d.pImageTexture, imageBoxX0, imageBoxY0, 0!, 0!, scaledW, scaledH, 0!, psPalette(WZCOL_WHITE)

    If d.submittedEntry Then
        pie_UniTransBoxFill imageBoxX0, imageBoxY0, imageBoxX0 + scaledW, imageBoxY0 + scaledH, psPalette(WZCOL_TRANSPARENT_BOX)
    End If
End Sub

Public Sub CaptchaPromptWidget_geometryChanged(ByVal self As WIDGET)
    CaptchaPromptWidget_recalcLayout self
End Sub

Private Sub CaptchaPromptWidget_recalcLayout(ByVal self As WIDGET)
    Dim d As CaptchaPromptWidget, w As Long, h As Long, inputBoxAreaHeight As Long, neededHeightForNonImageWidgets As Long
    Dim availableHeightForImage As Long, scaledImageWidth As Long, imageBoxY0 As Long, imageBoxY1 As Long
    Dim inputBoxAreaY0 As Long, availableInputBoxAreaWidth As Long, inputBoxAreaXPadding As Long, buttonX0 As Long, inputBoxWidth As Long
    Set d = CaptchaPromptWidget_(self)
    If d.pImageTexture Is Nothing Or d.inputBox Is Nothing Then Exit Sub   ' (VB) not initialized, or no image
    w = self.width()
    h = self.height()

    d.promptText.setGeometry 0, 0, w, d.promptText.idealHeight()

    inputBoxAreaHeight = jmax(d.inputBox.height(), d.okayButton.height())
    neededHeightForNonImageWidgets = d.promptText.height() + innerPaddingY + innerPaddingY + inputBoxAreaHeight
    availableHeightForImage = h - neededHeightForNonImageWidgets

    d.imageScalingRatio = getTextureDimensionsScaleRatioForMax(d.pImageTexture, w, availableHeightForImage)

    scaledImageWidth = CLng(ceil_(textureWidth(d.pImageTexture) * d.imageScalingRatio))

    imageBoxY0 = d.promptText.y() + d.promptText.height() + innerPaddingY
    imageBoxY1 = imageBoxY0 + CLng(Fix(textureHeight(d.pImageTexture) * d.imageScalingRatio))

    inputBoxAreaY0 = imageBoxY1 + innerPaddingY
    availableInputBoxAreaWidth = jmin(jmax(scaledImageWidth, minimumInputAreaWidth), w)
    inputBoxAreaXPadding = (w - availableInputBoxAreaWidth) \ 2

    buttonX0 = w - inputBoxAreaXPadding - d.okayButton.width()
    d.okayButton.setGeometry buttonX0, inputBoxAreaY0, d.okayButton.width(), inputBoxAreaHeight

    inputBoxWidth = jmax(jmax(buttonX0 - 5, 0) - inputBoxAreaXPadding, 0)
    d.inputBox.setGeometry inputBoxAreaXPadding, inputBoxAreaY0, inputBoxWidth, inputBoxAreaHeight
End Sub

Public Function CaptchaPromptWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As CaptchaPromptWidget, result As Long
    Set d = CaptchaPromptWidget_(self)
    If d.pImageTexture Is Nothing Or d.promptText Is Nothing Then Exit Function    ' (VB) no image
    result = d.promptText.idealWidth()
    result = jmax(textureWidth(d.pImageTexture), result)
    result = jmax(minimumInputAreaWidth, result)
    CaptchaPromptWidget_idealWidth = result
End Function

Public Function CaptchaPromptWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As CaptchaPromptWidget, result As Long
    Set d = CaptchaPromptWidget_(self)
    If d.pImageTexture Is Nothing Or d.inputBox Is Nothing Then Exit Function      ' (VB) no image
    result = Abs(iV_GetTextAboveBase(textFontId))
    result = result + d.promptText.idealHeight() + innerPaddingY
    result = result + textureHeight(d.pImageTexture) + innerPaddingY
    result = result + jmax(d.inputBox.height(), d.okayButton.height())

    CaptchaPromptWidget_idealHeight = result
End Function

Private Sub CaptchaPromptWidget_submitEntry(ByVal self As WIDGET, ByVal value As String)
    Dim d As CaptchaPromptWidget
    Set d = CaptchaPromptWidget_(self)
    If Not (Not d.responseCallback Is Nothing) Then ASSERT_FAILURE "responseCallback != nullptr", "lib/widget/verifyprompt.cpp:331", "submitEntry", "No response callback??": Exit Sub

    If d.submittedEntry Then
        Exit Sub
    End If

    WzFunc_call d.responseCallback, value

    If Not d.promptText Is Nothing Then
        W_LABEL_setFont d.promptText, textFontId, psPalette(WZCOL_TEXT_DARK)
    End If
    If Not d.inputBox Is Nothing Then
        W_EDITBOX_setState d.inputBox, WEDBS_DISABLE
        W_EDITBOX_setBoxColours d.inputBox, psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BORDER), psPalette(WZCOL_MENU_BACKGROUND)
    End If
    If Not d.okayButton Is Nothing Then
        d.okayButton.setState WBUT_DISABLE
    End If
    d.submittedEntry = True
End Sub

' MARK: - VerifySlideWidget

' VerifySlideWidget(W_SLDINIT const *init) : W_SLIDER(init)
Public Sub VerifySlideWidget_construct(ByVal self As WIDGET, init As W_SLDINIT)
    W_SLIDER_construct_init self, init
    self.kind = K_VerifySlideWidget     ' (VB) virtual calls from here on reach VerifySlideWidget's overrides
End Sub

Public Function VerifySlideWidget_make() As WIDGET
    Dim sSldInit As W_SLDINIT, widget As WIDGET
    W_SLDINIT_construct sSldInit
    sSldInit.orientation = WSLD_LEFT

    Set widget = WIDGET_alloc(K_VerifySlideWidget)
    VerifySlideWidget_construct widget, sSldInit
    W_SLIDER_(widget).barSize = 48
    Set VerifySlideWidget_make = widget
End Function

Public Sub VerifySlideWidget_setOnReleasedCallback(ByVal self As WIDGET, ByVal newOnReleasedCallback As WzFunc)
    Set VerifySlideWidget_(self).onReleasedCallback = newOnReleasedCallback
End Sub

Public Sub VerifySlideWidget_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim f As WzFunc
    W_SLIDER_released self, psContext, key

    Set f = VerifySlideWidget_(self).onReleasedCallback
    If Not f Is Nothing Then
        WzFunc_call f, self
    End If
End Sub

Public Sub VerifySlideWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, w As Long, h As Long, barHeight As Long, barY0 As Long, s As W_SLIDER, d As VerifySlideWidget
    Dim sliderX As Long, sliderDrawX0 As Long, sliderDrawX1 As Long, arrowTextX0 As Single, arrowTextY0 As Single, colour As Long
    Set s = W_SLIDER_(self)
    Set d = VerifySlideWidget_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    w = self.width()
    h = self.height()

    ' Draw bar

    barHeight = jmax(h \ 2, 1)
    barY0 = y0 + ((h - barHeight) \ 2)
    pie_UniTransBoxFill x0, barY0, x0 + w, barY0 + barHeight, psPalette(WZCOL_DBLUE)

    ' Draw slider box

    If s.numStops > 0 Then sliderX = (w - s.barSize) * s.pos \ s.numStops Else sliderX = 0
    sliderDrawX0 = x0 + sliderX + 2
    sliderDrawX1 = x0 + sliderX + jmax(s.barSize, 5) - 2
    If sliderDrawX1 - sliderDrawX0 <= 2 Then
        sliderDrawX0 = sliderDrawX0 - 1
        sliderDrawX1 = sliderDrawX1 + 1
    End If
    If W_SLIDER_isEnabled(self) Then colour = psPalette(WZCOL_LBLUE) Else colour = psPalette(WZCOL_DBLUE)
    pie_UniTransBoxFill sliderDrawX0, y0, sliderDrawX1, y0 + h, colour

    ' Draw right-facing arrow centered in slider box // "→" = ->

    d.rightArrowText.setText ChrW$(&H2192), font_medium_bold
    arrowTextX0 = sliderDrawX0 + (((sliderDrawX1 - sliderDrawX0) - d.rightArrowText.width()) \ 2)
    arrowTextY0 = y0 + (h - d.rightArrowText.lineSize()) \ 2 - d.rightArrowText.aboveBase()
    If W_SLIDER_isEnabled(self) Then colour = psPalette(WZCOL_WHITE) Else colour = psPalette(WZCOL_TEXT_DARK)
    d.rightArrowText.render arrowTextX0, arrowTextY0, colour
End Sub

' MARK: - SlidePromptWidget

' SlidePromptWidget()
Public Sub SlidePromptWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_SlidePromptWidget     ' (VB) virtual calls from here on reach SlidePromptWidget's overrides
    SlidePromptWidget_(self).weakSerial = WeakWidget_register(self)
End Sub

' ~SlidePromptWidget()
Public Sub SlidePromptWidget_destruct(ByVal self As WIDGET)
    Dim d As SlidePromptWidget
    Set d = SlidePromptWidget_(self)
    If Not d.pImageTexture Is Nothing Then
        Set d.pImageTexture = Nothing   ' delete pImageTexture
    End If

    If Not d.pTileTexture Is Nothing Then
        Set d.pTileTexture = Nothing    ' delete pTileTexture
    End If
    WeakWidget_unregister d.weakSerial
End Sub

Public Function SlidePromptWidget_make(promptData As SlideVerifyPromptData, ByVal dragDrop As Boolean, ByVal responseCallback As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_SlidePromptWidget)
    SlidePromptWidget_construct widget
    SlidePromptWidget_initialize widget, promptData, dragDrop, responseCallback
    Set SlidePromptWidget_make = widget
End Function

Private Function SlidePromptWidget_initialize(ByVal self As WIDGET, promptData As SlideVerifyPromptData, ByVal dragDrop As Boolean, ByVal responseCallback As WzFunc) As Boolean
    Dim d As SlidePromptWidget, image As iV_Image, tile As iV_Image, weakParent As Long
    Set d = SlidePromptWidget_(self)
    Set d.promptData_image = promptData.image
    Set d.promptData_tile = promptData.tile
    d.promptData_tile_width = promptData.tile_width
    d.promptData_tile_height = promptData.tile_height
    d.promptData_tile_x = promptData.tile_x
    d.promptData_tile_y = promptData.tile_y
    d.dragDrop = dragDrop
    Set d.responseCallback = responseCallback

    Set d.promptText = W_LABEL_new()
    W_LABEL_setFont d.promptText, textFontId, psPalette(WZCOL_FORM_LIGHT)
    If Not dragDrop Then
        d.promptText.setString "Please drag the slider to verify"
    Else
        d.promptText.setString "Please drag the tile to verify"
    End If
    W_LABEL_setCanTruncate d.promptText, True
    d.promptText.setTransparentToClicks True
    self.attach d.promptText

    If Not (Not d.promptData_image Is Nothing) Then ASSERT_FAILURE "promptData.image != nullptr", "lib/widget/verifyprompt.cpp:545", "initialize", "No image to load": Exit Function
    If Not (d.promptData_image.GetImage(image)) Then ASSERT_FAILURE "image.has_value()", "lib/widget/verifyprompt.cpp:547", "initialize", "Failed to load image": Exit Function

    If Not (Not d.promptData_tile Is Nothing) Then ASSERT_FAILURE "promptData.tile != nullptr", "lib/widget/verifyprompt.cpp:549", "initialize", "No tile to load": Exit Function
    If Not (d.promptData_tile.GetImage(tile)) Then ASSERT_FAILURE "tile.has_value()", "lib/widget/verifyprompt.cpp:551", "initialize", "Failed to load tile": Exit Function

    d.imageWidth = image.width()
    d.imageHeight = image.height()

    Set d.pImageTexture = gfx_api_context_loadTextureFromUncompressedImage(image, texture_type_user_interface, "slide_image")
    If Not (Not d.pImageTexture Is Nothing) Then ASSERT_FAILURE "pImageTexture != nullptr", "lib/widget/verifyprompt.cpp:557", "initialize", "Failed to load image texture": Exit Function

    Set d.pTileTexture = gfx_api_context_loadTextureFromUncompressedImage(tile, texture_type_user_interface, "slide_tile")
    If Not (Not d.pTileTexture Is Nothing) Then ASSERT_FAILURE "pTileTexture != nullptr", "lib/widget/verifyprompt.cpp:560", "initialize", "Failed to load tile texture": Exit Function

    d.currTileX = promptData.tile_x
    d.currTileY = promptData.tile_y

    weakParent = d.weakSerial

    If Not dragDrop Then
        ' Make slider with custom drawing function that draws a slider bar + box with -> arrow in it
        Set d.slider = VerifySlideWidget_make()
        W_SLIDER_addOnChange d.slider, WzFunc_new(FN_verifyprompt_slideOnChange, weakParent)
        VerifySlideWidget_setOnReleasedCallback d.slider, WzFunc_new(FN_verifyprompt_slideOnReleased, weakParent)
        W_SLIDER_(d.slider).numStops = (textureWidth(d.pImageTexture) - (promptData.tile_x + promptData.tile_width)) And &HFFFF&
        W_SLIDER_(d.slider).pos = 0
        self.attach d.slider
    End If

    SlidePromptWidget_initialize = True
End Function

Private Sub SlidePromptWidget_handleSliderPosChanged(ByVal self As WIDGET)
    Dim d As SlidePromptWidget
    Set d = SlidePromptWidget_(self)
    d.currTileX = d.promptData_tile_x + W_SLIDER_(d.slider).pos
End Sub

Private Sub SlidePromptWidget_handleSliderInputReleased(ByVal self As WIDGET)
    Dim d As SlidePromptWidget
    Set d = SlidePromptWidget_(self)
    If Not (Not d.responseCallback Is Nothing) Then ASSERT_FAILURE "responseCallback != nullptr", "lib/widget/verifyprompt.cpp:596", "handleSliderInputReleased", "No response callback??": Exit Sub

    If d.submittedResponse Then
        Exit Sub
    End If

    WzFunc_call d.responseCallback, d.currTileX, d.currTileY

    If Not d.promptText Is Nothing Then
        W_LABEL_setFont d.promptText, textFontId, psPalette(WZCOL_TEXT_DARK)
    End If
    If Not d.slider Is Nothing Then
        W_SLIDER_disable d.slider
    End If

    d.submittedResponse = True

    SlidePromptWidget_recalcLayout self
End Sub

Public Sub SlidePromptWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As SlidePromptWidget, x0 As Long, y0 As Long, scaledW As Single, scaledH As Single
    Dim imageAreaXPadding As Long, imageBoxX0 As Long, imageBoxY0 As Long
    Set d = SlidePromptWidget_(self)
    If d.pImageTexture Is Nothing Or d.pTileTexture Is Nothing Then Exit Sub   ' (VB) the prompt's images failed to load
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    scaledW = textureWidth(d.pImageTexture) * d.imageScalingRatio
    scaledH = textureHeight(d.pImageTexture) * d.imageScalingRatio

    imageAreaXPadding = (self.width() - CLng(ceil_(scaledW))) \ 2
    imageBoxX0 = x0 + imageAreaXPadding
    imageBoxY0 = y0 + d.promptText.y() + d.promptText.height() + innerPaddingY

    ' Draw the pImageTexture
    iV_DrawImageAnisotropic d.pImageTexture, imageBoxX0, imageBoxY0, 0!, 0!, scaledW, scaledH, 0!, psPalette(WZCOL_WHITE)

    ' Draw the pTileTexture at (currTileX, currTileY)
    iV_DrawImageAnisotropic d.pTileTexture, imageBoxX0, imageBoxY0, CSng(d.currTileX * d.imageScalingRatio), CSng(d.currTileY * d.imageScalingRatio), _
        CSng(d.promptData_tile_width * d.imageScalingRatio), CSng(d.promptData_tile_height * d.imageScalingRatio), 0!, psPalette(WZCOL_WHITE)

    If d.submittedResponse Then
        pie_UniTransBoxFill imageBoxX0, imageBoxY0, imageBoxX0 + scaledW, imageBoxY0 + scaledH, psPalette(WZCOL_TRANSPARENT_BOX)
    End If
End Sub

Public Sub SlidePromptWidget_geometryChanged(ByVal self As WIDGET)
    SlidePromptWidget_recalcLayout self
End Sub

' WzRect calcTileRect() const: tileX0, tileY0, its width and height
Private Sub SlidePromptWidget_calcTileRect(ByVal self As WIDGET, tileX0 As Long, tileY0 As Long, tileDisplayWidth As Long, tileDisplayHeight As Long)
    Dim d As SlidePromptWidget, w As Long, scaledImageWidth As Long, imageAreaXPadding As Long, imageBoxX0 As Long, imageBoxY0 As Long
    Set d = SlidePromptWidget_(self)
    w = self.width()

    scaledImageWidth = CLng(ceil_(textureWidth(d.pImageTexture) * d.imageScalingRatio))

    imageAreaXPadding = (w - scaledImageWidth) \ 2
    imageBoxX0 = imageAreaXPadding
    imageBoxY0 = d.promptText.y() + d.promptText.height() + innerPaddingY

    tileX0 = imageBoxX0 + CLng(ceil_(d.currTileX * d.imageScalingRatio))
    tileY0 = imageBoxY0 + CLng(ceil_(d.currTileY * d.imageScalingRatio))

    tileDisplayWidth = CLng(ceil_(d.promptData_tile_width * d.imageScalingRatio))
    tileDisplayHeight = CLng(ceil_(d.promptData_tile_height * d.imageScalingRatio))
End Sub

Private Sub SlidePromptWidget_recalcLayout(ByVal self As WIDGET)
    Dim d As SlidePromptWidget, w As Long, h As Long, neededHeightForNonImageWidgets As Long, availableHeightForImage As Long
    Dim scaledImageWidth As Long, imageBoxY0 As Long, imageBoxY1 As Long, availableInputAreaWidth As Long, inputAreaXPadding As Long, sliderY0 As Long
    Set d = SlidePromptWidget_(self)
    If d.pImageTexture Is Nothing Then Exit Sub        ' (VB) not initialized, or no image
    w = self.width()
    h = self.height()

    d.promptText.setGeometry 0, 0, w, d.promptText.idealHeight()

    If d.dragDrop Then
        neededHeightForNonImageWidgets = d.promptText.height() + innerPaddingY
    Else
        neededHeightForNonImageWidgets = d.promptText.height() + innerPaddingY + (innerPaddingY + sliderHeight)
    End If
    availableHeightForImage = h - neededHeightForNonImageWidgets

    d.imageScalingRatio = getTextureDimensionsScaleRatioForMax(d.pImageTexture, w, availableHeightForImage)

    scaledImageWidth = CLng(ceil_(textureWidth(d.pImageTexture) * d.imageScalingRatio))

    imageBoxY0 = d.promptText.y() + d.promptText.height() + innerPaddingY
    imageBoxY1 = imageBoxY0 + CLng(Fix(textureHeight(d.pImageTexture) * d.imageScalingRatio))

    If Not d.slider Is Nothing Then
        availableInputAreaWidth = jmin(jmax(scaledImageWidth, minimumInputAreaWidth), w)
        inputAreaXPadding = (w - availableInputAreaWidth) \ 2

        sliderY0 = imageBoxY1 + innerPaddingY

        d.slider.setGeometry inputAreaXPadding, sliderY0, availableInputAreaWidth, sliderHeight
    End If
End Sub

Public Function SlidePromptWidget_idealWidth(ByVal self As WIDGET) As Long
    Dim d As SlidePromptWidget, result As Long
    Set d = SlidePromptWidget_(self)
    If d.pImageTexture Is Nothing Or d.promptText Is Nothing Then Exit Function    ' (VB) no image
    result = d.promptText.idealWidth()
    result = jmax(textureWidth(d.pImageTexture), result)
    result = jmax(minimumInputAreaWidth, result)
    SlidePromptWidget_idealWidth = result
End Function

Public Function SlidePromptWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As SlidePromptWidget, result As Long
    Set d = SlidePromptWidget_(self)
    If d.pImageTexture Is Nothing Or d.promptText Is Nothing Then Exit Function    ' (VB) no image
    result = Abs(iV_GetTextAboveBase(textFontId))
    result = result + d.promptText.idealHeight() + innerPaddingY
    result = result + textureHeight(d.pImageTexture)
    If Not d.dragDrop Then
        result = result + innerPaddingY + sliderHeight
    End If

    SlidePromptWidget_idealHeight = result
End Function

Public Function SlidePromptWidget_capturesMouseDrag(ByVal self As WIDGET, ByVal key As WIDGET_KEY) As Boolean
    Dim d As SlidePromptWidget
    Set d = SlidePromptWidget_(self)
    SlidePromptWidget_capturesMouseDrag = d.dragDrop And Not d.submittedResponse
End Function

Public Sub SlidePromptWidget_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As SlidePromptWidget
    WIDGET_released self, psContext, key

    Set d = SlidePromptWidget_(self)
    If d.has_handlingDragTileStartMouseOffset Then
        If d.dragDrop Then
            SlidePromptWidget_handleSliderInputReleased self
        End If
        d.has_handlingDragTileStartMouseOffset = False
    End If
End Sub

' handling mouse drag
Public Sub SlidePromptWidget_mouseDragged(ByVal self As WIDGET, ByVal key As WIDGET_KEY, start As W_CONTEXT, current As W_CONTEXT)
    Dim d As SlidePromptWidget, tileX0 As Long, tileY0 As Long, tileW As Long, tileH As Long, startMx As Long, startMy As Long
    Dim w As Long, scaledImageWidth As Long, imageAreaXPadding As Long, imageBoxX0 As Long, imageBoxY0 As Long
    Dim mx As Long, my As Long, tileDragPosAdjustedX As Long, tileDragPosAdjustedY As Long
    Dim proposedNewCurrTileX As Double, proposedNewCurrTileY As Double, maxX As Double, maxY As Double
    Set d = SlidePromptWidget_(self)
    If d.pImageTexture Is Nothing Then Exit Sub        ' (VB) no image
    If Not d.has_handlingDragTileStartMouseOffset Then
        SlidePromptWidget_calcTileRect self, tileX0, tileY0, tileW, tileH
        startMx = start.mx - self.x()
        startMy = start.my - self.y()
        ' tileRect.contains(startMx, startMy)
        If Not (Not (startMx < tileX0) And (startMx < tileX0 + tileW) And Not (startMy < tileY0) And (startMy < tileY0 + tileH)) Then
            Exit Sub
        End If

        ' click started inside the tile rect - process drag
        d.has_handlingDragTileStartMouseOffset = True
        d.handlingDragTileStartMouseOffsetX = startMx - tileX0
        d.handlingDragTileStartMouseOffsetY = startMy - tileY0
    End If

    w = self.width()

    scaledImageWidth = CLng(ceil_(textureWidth(d.pImageTexture) * d.imageScalingRatio))

    imageAreaXPadding = (w - scaledImageWidth) \ 2
    imageBoxX0 = imageAreaXPadding
    imageBoxY0 = d.promptText.y() + d.promptText.height() + innerPaddingY

    mx = current.mx - self.x()
    my = current.my - self.y()
    tileDragPosAdjustedX = mx - d.handlingDragTileStartMouseOffsetX - imageBoxX0
    tileDragPosAdjustedY = my - d.handlingDragTileStartMouseOffsetY - imageBoxY0

    ' convert to logical tile coordinates
    proposedNewCurrTileX = round_(tileDragPosAdjustedX / d.imageScalingRatio)
    proposedNewCurrTileY = round_(tileDragPosAdjustedY / d.imageScalingRatio)
    ' std::clamp<int64_t>(proposed, 0, imageWidth - tile_width) (uint32_t arithmetic: the difference wraps)
    maxX = CDbl(d.imageWidth) - CDbl(d.promptData_tile_width)
    If maxX < 0 Then maxX = maxX + 4294967296#
    maxY = CDbl(d.imageHeight) - CDbl(d.promptData_tile_height)
    If maxY < 0 Then maxY = maxY + 4294967296#
    If proposedNewCurrTileX < 0 Then proposedNewCurrTileX = 0
    If proposedNewCurrTileX > maxX Then proposedNewCurrTileX = maxX
    If proposedNewCurrTileY < 0 Then proposedNewCurrTileY = 0
    If proposedNewCurrTileY > maxY Then proposedNewCurrTileY = maxY
    d.currTileX = CLng(proposedNewCurrTileX)
    d.currTileY = CLng(proposedNewCurrTileY)
End Sub

' MARK: - Centered Wrapper Widget

' CaptchaWrapperWidget() {}
Public Sub CaptchaWrapperWidget_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_CaptchaWrapperWidget  ' (VB) virtual calls from here on reach CaptchaWrapperWidget's overrides
End Sub

Public Function CaptchaWrapperWidget_make(ByVal wrappedCaptchaWidget As WIDGET) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_CaptchaWrapperWidget)
    CaptchaWrapperWidget_construct widget
    CaptchaWrapperWidget_initialize widget, wrappedCaptchaWidget
    Set CaptchaWrapperWidget_make = widget
End Function

Private Sub CaptchaWrapperWidget_initialize(ByVal self As WIDGET, ByVal wrappedCaptchaWidget As WIDGET)
    Dim d As CaptchaWrapperWidget
    Set d = CaptchaWrapperWidget_(self)
    Set d.wrappedCaptchaWidget = wrappedCaptchaWidget
    self.attach d.wrappedCaptchaWidget

    d.cachedIdealWidth = d.wrappedCaptchaWidget.idealWidth()
    d.cachedIdealHeight = d.wrappedCaptchaWidget.idealHeight()
End Sub

Public Sub CaptchaWrapperWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    ' Draw background
    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(0, 0, 0, 35)
End Sub

Public Sub CaptchaWrapperWidget_geometryChanged(ByVal self As WIDGET)
    CaptchaWrapperWidget_recalcLayout self
End Sub

Public Function CaptchaWrapperWidget_idealWidth(ByVal self As WIDGET) As Long
    CaptchaWrapperWidget_idealWidth = CaptchaWrapperWidget_(self).cachedIdealWidth + (Wrapper_innerPaddingX * 2)
End Function

Public Function CaptchaWrapperWidget_idealHeight(ByVal self As WIDGET) As Long
    CaptchaWrapperWidget_idealHeight = CaptchaWrapperWidget_(self).cachedIdealHeight + (Wrapper_innerPaddingY * 2)
End Function

Private Sub CaptchaWrapperWidget_recalcLayout(ByVal self As WIDGET)
    Dim d As CaptchaWrapperWidget, w As Long, h As Long, wrappedWidth As Long, wrappedHeight As Long, wrappedX0 As Long, wrappedY0 As Long
    Set d = CaptchaWrapperWidget_(self)
    If d.wrappedCaptchaWidget Is Nothing Then Exit Sub ' (VB) not initialized yet
    w = self.width()
    h = self.height()

    ' position wrappedCaptchaWidget centered in the parent
    wrappedWidth = jmin(w - (Wrapper_innerPaddingX * 2), d.cachedIdealWidth)
    wrappedHeight = jmin(h - (Wrapper_innerPaddingY * 2), d.cachedIdealHeight)
    wrappedX0 = (w - wrappedWidth) \ 2
    wrappedY0 = (h - wrappedHeight) \ 2

    d.wrappedCaptchaWidget.setGeometry wrappedX0, wrappedY0, wrappedWidth, wrappedHeight
End Sub

' MARK: - Public API

' std::shared_ptr<WIDGET> makeCaptchaWidget(CaptchaVerifyPromptData&& promptData, const CaptchaVerifyResponseFunc& responseCallback)
Public Function WzVerifyPrompt_makeCaptchaWidget(promptData As CaptchaVerifyPromptData, ByVal responseCallback As WzFunc) As WIDGET
    Set WzVerifyPrompt_makeCaptchaWidget = CaptchaWrapperWidget_make(CaptchaPromptWidget_make(promptData, responseCallback))
End Function

' std::shared_ptr<WIDGET> makeSlideWidget(SlideVerifyPromptData&& promptData, bool dragDrop, const SlideVerifyResponseFunc& responseCallback)
Public Function WzVerifyPrompt_makeSlideWidget(promptData As SlideVerifyPromptData, ByVal dragDrop As Boolean, ByVal responseCallback As WzFunc) As WIDGET
    Set WzVerifyPrompt_makeSlideWidget = CaptchaWrapperWidget_make(SlidePromptWidget_make(promptData, dragDrop, responseCallback))
End Function

' The lambdas
Public Function verifyprompt_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongParent As WIDGET, widg As WIDGET, str_ As String
    Select Case f.id
    Case FN_verifyprompt_captchaOnReturn            ' [weakParent](W_EDITBOX& widg)
        Set widg = args(0)
        str_ = W_EDITBOX_getString(widg)
        widgScheduleTask WzFunc_new(FN_verifyprompt_captchaSubmitTask, str_, f.cap(0))

    Case FN_verifyprompt_captchaSubmitTask          ' [str, weakParent]()
        Set strongParent = WeakWidget_lock(f.cap(1))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "lib/widget/verifyprompt.cpp:235", "initialize", "No parent?": Exit Function
        CaptchaPromptWidget_submitEntry strongParent, f.cap(0)

    Case FN_verifyprompt_captchaOkClick             ' [weakParent](W_BUTTON&)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "lib/widget/verifyprompt.cpp:245", "initialize", "No parent?": Exit Function
        CaptchaPromptWidget_submitEntry strongParent, W_EDITBOX_getString(CaptchaPromptWidget_(strongParent).inputBox)

    Case FN_verifyprompt_slideOnChange              ' [weakParent](W_SLIDER &)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "lib/widget/verifyprompt.cpp:573", "initialize", "Null parent": Exit Function
        SlidePromptWidget_handleSliderPosChanged strongParent

    Case FN_verifyprompt_slideOnReleased            ' [weakParent](VerifySlideWidget &)
        Set strongParent = WeakWidget_lock(f.cap(0))
        If Not (Not strongParent Is Nothing) Then ASSERT_FAILURE "strongParent != nullptr", "lib/widget/verifyprompt.cpp:578", "initialize", "Null parent": Exit Function
        SlidePromptWidget_handleSliderInputReleased strongParent
    End Select
End Function
