Attribute VB_Name = "musicmanager_c"
' src/musicmanager.cpp, musicmanager.h, whole: the Music Manager screen (the album/track table with a checkbox per
' game mode, the "now playing" details box), on the title screens (WzMusicManagerTitleUI) and over the game
' (audiooptions' doOpenMusicManager). Its widget classes' fields are the data classes named as the classes; class
' MusicManager_CDAudioEventSink (a CDAudioEventSink) is MusicManager_CDAudioEventSink.cls, implementing WzSound's
' ICDAudioEventSink. A std::shared_ptr<const WZ_TRACK> is a WZ_TRACK (Nothing = nullptr); the form's
' std::vector<std::shared_ptr<const WZ_TRACK>> trackList is a Collection (trackIdx 0-based: trackList(trackIdx + 1));
' a std::vector<std::shared_ptr<WIDGET>> (TableRow::make's columns) is a Collection; std::vector<TableColumn> and
' std::vector<size_t> are arrays and their counts (table_c's convention). The form's std::weak_ptr (the track rows'
' click handlers, the event sink) is its widgweak_vb serial. A gfx_api::texture* is a gfx_api_texture (delete is
' Set Nothing). MusicGameMode is SoundApi's enum. Translations (_()) are not ported: the English text is used.
' (VB) widget class W_MusicModeCheckboxButton : W_BUTTON
' (VB) widget class TrackDetailsForm : W_FORM
' (VB) widget class W_MusicListHeaderColImage : W_BUTTON
' (VB) widget class MusicManagerForm : W_FORM
Option Explicit

Private Const W_TRACK_ROW_PADDING As Long = 5
Private Const W_TRACK_COL_PADDING As Long = 10
Private Const W_TRACK_CHECKBOX_SIZE As Long = 12

Private Const W_TRACK_COL_TITLE_X As Long = W_TRACK_COL_PADDING
Private Const W_TRACK_COL_TITLE_W As Long = 180
Private Const W_TRACK_COL_ALBUM_X As Long = (W_TRACK_COL_TITLE_X + W_TRACK_COL_TITLE_W + W_TRACK_COL_PADDING)
Private Const W_TRACK_COL_ALBUM_W As Long = 130
Private Const W_TRACK_CHECKBOX_STARTINGPOS As Long = (W_TRACK_COL_ALBUM_X + W_TRACK_COL_ALBUM_W + W_TRACK_COL_PADDING + W_TRACK_COL_PADDING)

Private Const W_TRACK_HEADER_Y As Long = 20
Private Const W_TRACK_HEADER_HEIGHT As Long = (20 + (W_TRACK_ROW_PADDING * 2))
Private Const W_TRACK_HEADER_COL_IMAGE_SIZE As Long = 16

Private Const TL_W As Long = FRONTEND_BOTFORMW
Private Const TL_H As Long = 400
Private Const TL_X As Long = FRONTEND_BOTFORMX
Private Const TL_SX As Long = FRONTEND_SIDEX

Private Const TL_ENTRYW As Long = (FRONTEND_BOTFORMW - 80)
Private Const TL_ENTRYH As Long = (25)

Private Const TL_PREVIEWBOX_Y_SPACING As Long = 45
Private Const TL_PREVIEWBOX_H As Long = 84

Private Const WZ_TRACKDETAILS_IMAGE_Y As Long = 12
Private Const WZ_TRACKDETAILS_IMAGE_SIZE As Long = 60

' cdaudio.h: #define NUM_MUSICGAMEMODES 4
Private Const NUM_MUSICGAMEMODES As Long = 4

' the lambdas (WzFunc: musicmanager_c_fn)
Public Const FN_musicModeCheckbox_onClick As Long = 15300   ' W_MusicModeCheckboxButton(...): addOnClickHandler([](W_BUTTON& button) {...})
Public Const FN_musicManager_trackRowClick As Long = 15301  ' addTrackList: trackRow->addOnClickHandler([weakSelf, trackIdx](W_BUTTON& clickedButton) {...})
Public Const FN_musicManager_setTrackMode As Long = 15302   ' createTrackRow: pCheckBox->addOnClickHandler([captureTrack](W_BUTTON& button) {...})

'--------------------------------------------------------------------------
' (VB) helpers

' a PIELIGHT with its byte.a replaced
Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' MARK: - W_MusicModeCheckboxButton

' W_MusicModeCheckboxButton(MusicGameMode mode, bool isChecked): W_BUTTON(), mode(mode), isChecked(isChecked)
Public Sub W_MusicModeCheckboxButton_construct(ByVal self As WIDGET, ByVal mode As MusicGameMode, ByVal isChecked As Boolean)
    Dim d As W_MusicModeCheckboxButton
    W_BUTTON_construct self
    self.kind = K_W_MusicModeCheckboxButton ' (VB) virtual calls from here on reach W_MusicModeCheckboxButton's overrides
    Set d = W_MusicModeCheckboxButton_(self)
    d.mode = mode
    d.isChecked = isChecked
    d.cbSize = 0
    W_BUTTON_addOnClickHandler self, WzFunc_new(FN_musicModeCheckbox_onClick)
End Sub

' std::make_shared<W_MusicModeCheckboxButton>(mode, isChecked)
Public Function W_MusicModeCheckboxButton_new(ByVal mode As MusicGameMode, ByVal isChecked As Boolean) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_MusicModeCheckboxButton)
    W_MusicModeCheckboxButton_construct w, mode, isChecked
    Set W_MusicModeCheckboxButton_new = w
End Function

Public Sub W_MusicModeCheckboxButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_MusicModeCheckboxButton, x0 As Long, y0 As Long, down As Boolean, isDisabled As Boolean
    Dim checkboxOffset As Vector2i, checkboxPos As Vector2i, notifyBoxAddColor As Long, checkBoxOutsideColor As Long
    Dim checkBoxInsideColor As Long, factor As Single
    Const CB_INNER_INSET As Long = 2
    Set d = W_MusicModeCheckboxButton_(self)
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()

    down = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0

    ' calculate checkbox dimensions
    checkboxOffset = Vector2i_(0, (self.height() - d.cbSize) \ 2) ' left-align, center vertically
    checkboxPos = Vector2i_(x0 + checkboxOffset.x, y0 + checkboxOffset.y)

    ' draw checkbox border
    notifyBoxAddColor = psPalette(WZCOL_NOTIFICATION_BOX)
    If Not isDisabled Then factor = 0.7! Else factor = 0.2!
    notifyBoxAddColor = pielightWithAlpha(notifyBoxAddColor, CLng(Fix(CSng(SHRU(notifyBoxAddColor, 24)) * factor)))
    pie_UniTransBoxFill checkboxPos.x, checkboxPos.y, checkboxPos.x + d.cbSize, checkboxPos.y + d.cbSize, notifyBoxAddColor
    checkBoxOutsideColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        checkBoxOutsideColor = pielightWithAlpha(checkBoxOutsideColor, 60)
    End If
    iV_Box2 checkboxPos.x, checkboxPos.y, checkboxPos.x + d.cbSize, checkboxPos.y + d.cbSize, checkBoxOutsideColor, checkBoxOutsideColor

    If down Or d.isChecked Then
        ' draw checkbox "checked" inside
        checkBoxInsideColor = psPalette(WZCOL_TEXT_MEDIUM)
        If Not isDisabled Then
            checkBoxInsideColor = pielightWithAlpha(checkBoxInsideColor, 200)
        Else
            checkBoxInsideColor = pielightWithAlpha(checkBoxInsideColor, 60)
        End If
        pie_UniTransBoxFill checkboxPos.x + CB_INNER_INSET, checkboxPos.y + CB_INNER_INSET, checkboxPos.x + d.cbSize - (CB_INNER_INSET), checkboxPos.y + d.cbSize - (CB_INNER_INSET), checkBoxInsideColor
    End If
End Sub

Public Sub W_MusicModeCheckboxButton_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    If Not W_MusicModeCheckboxButton_isEnabled(self) Then Exit Sub
    W_BUTTON_highlight self, psContext
End Sub

Public Sub W_MusicModeCheckboxButton_highlightLost(ByVal self As WIDGET)
    If Not W_MusicModeCheckboxButton_isEnabled(self) Then Exit Sub
    W_BUTTON_highlightLost self
End Sub

Public Function W_MusicModeCheckboxButton_getIsChecked(ByVal self As WIDGET) As Boolean
    W_MusicModeCheckboxButton_getIsChecked = W_MusicModeCheckboxButton_(self).isChecked
End Function

Public Sub W_MusicModeCheckboxButton_setCheckboxSize(ByVal self As WIDGET, ByVal size As Long)
    W_MusicModeCheckboxButton_(self).cbSize = size
End Sub

Public Function W_MusicModeCheckboxButton_checkboxSize(ByVal self As WIDGET) As Long
    W_MusicModeCheckboxButton_checkboxSize = W_MusicModeCheckboxButton_(self).cbSize
End Function

Public Function W_MusicModeCheckboxButton_getMusicMode(ByVal self As WIDGET) As MusicGameMode
    W_MusicModeCheckboxButton_getMusicMode = W_MusicModeCheckboxButton_(self).mode
End Function

Private Function W_MusicModeCheckboxButton_isEnabled(ByVal self As WIDGET) As Boolean
    W_MusicModeCheckboxButton_isEnabled = (self.getState() And WBUT_DISABLE) = 0
End Function

' MARK: - "Now Playing" Details Block

Private Function loadImageToTexture(ByVal imagePath As String) As gfx_api_texture
    Dim filename As String, dotPos As Long, extension As String
    If Len(imagePath) = 0 Then
        Set loadImageToTexture = Nothing
        Exit Function
    End If

    filename = imagePath
    dotPos = InStrRev(filename, ".")   ' determine the filetype (strrchr(filename, '.'))
    If dotPos > 0 Then extension = Mid$(filename, dotPos)

    If dotPos = 0 Or StrComp(extension, ".png", vbBinaryCompare) <> 0 Then
        debug_ 198, LOG_ERROR, "loadImageToTexture", "Bad image filename: %s", filename
        Set loadImageToTexture = Nothing
        Exit Function
    End If

    Set loadImageToTexture = gfx_api_context_loadTextureFromFile(imagePath, texture_type_user_interface)
End Function

' TrackDetailsForm(): W_FORM()
Public Sub TrackDetailsForm_construct(ByVal self As WIDGET)
    Dim d As TrackDetailsForm
    W_FORM_construct self
    self.kind = K_TrackDetailsForm ' (VB) virtual calls from here on reach TrackDetailsForm's overrides
    Set d = TrackDetailsForm_(self)
    Set d.pAlbumCoverTexture = Nothing
    d.outerHorizontalPadding = 20
End Sub

' virtual ~TrackDetailsForm()
Public Sub TrackDetailsForm_destruct(ByVal self As WIDGET)
    Dim d As TrackDetailsForm
    Set d = TrackDetailsForm_(self)
    If Not d.pAlbumCoverTexture Is Nothing Then
        Set d.pAlbumCoverTexture = Nothing  ' delete pAlbumCoverTexture
    End If
End Sub

Public Function TrackDetailsForm_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_TrackDetailsForm)
    TrackDetailsForm_construct widget
    TrackDetailsForm_initialize widget
    Set TrackDetailsForm_make = widget
End Function

Public Function TrackDetailsForm_loadImage(ByVal self As WIDGET, ByVal imagePath As String) As Boolean
    Dim d As TrackDetailsForm
    Set d = TrackDetailsForm_(self)
    If Len(imagePath) = 0 Then
        If Not d.pAlbumCoverTexture Is Nothing Then
            Set d.pAlbumCoverTexture = Nothing
        End If
        d.album_cover_path = ""
        TrackDetailsForm_loadImage = True
        Exit Function
    End If

    If imagePath = d.album_cover_path Then
        ' already loaded
        TrackDetailsForm_loadImage = True
        Exit Function
    End If

    If Not d.pAlbumCoverTexture Is Nothing Then
        Set d.pAlbumCoverTexture = Nothing
    End If
    Set d.pAlbumCoverTexture = loadImageToTexture(imagePath)
    If d.pAlbumCoverTexture Is Nothing Then
        d.album_cover_path = ""
        TrackDetailsForm_loadImage = False
        Exit Function
    End If
    d.album_cover_path = imagePath
    TrackDetailsForm_loadImage = True
End Function

Public Sub TrackDetailsForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As TrackDetailsForm, x0 As Long, y0 As Long, w As Long, h As Long, imageLeft As Long, imageTop As Long
    Set d = TrackDetailsForm_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()
    h = self.height()

    pie_UniTransBoxFill x0, y0, x0 + w, y0 + h, pal_RGBA(0, 0, 0, 100)

    If W_FORM_(self).disableChildren Then Exit Sub

    ' now draw the album cover, if present
    imageLeft = x0 + d.outerHorizontalPadding + d.psNowPlaying.width() + 10
    imageTop = y0 + WZ_TRACKDETAILS_IMAGE_Y

    If Not d.pAlbumCoverTexture Is Nothing Then
        iV_DrawImageAnisotropic d.pAlbumCoverTexture, imageLeft, imageTop, 0!, 0!, WZ_TRACKDETAILS_IMAGE_SIZE, WZ_TRACKDETAILS_IMAGE_SIZE, 0!, psPalette(WZCOL_WHITE)
    End If
End Sub

Public Sub TrackDetailsForm_geometryChanged(ByVal self As WIDGET)
    Dim d As TrackDetailsForm, w As Long, totalUsableTextWidth As Long, leftSideWidth As Long
    Dim albumInfoXPosStart As Long, maxWidthOfAlbumLabel As Long, descriptionContainerY0 As Long
    Dim descriptionContainerHeight As Long
    Set d = TrackDetailsForm_(self)
    w = self.width()
    totalUsableTextWidth = w - (d.outerHorizontalPadding * 2) - 10 - WZ_TRACKDETAILS_IMAGE_SIZE - 10
    leftSideWidth = (totalUsableTextWidth \ 2) - 10

    d.psNowPlaying.setGeometry d.outerHorizontalPadding, 14, leftSideWidth, 20
    d.psSelectedTrackName.setGeometry d.outerHorizontalPadding, 14 + 20, leftSideWidth, 20
    d.psSelectedTrackAuthorName.setGeometry d.outerHorizontalPadding, 14 + 20 + 17, leftSideWidth, 20

    albumInfoXPosStart = d.psNowPlaying.x() + d.psNowPlaying.width() + 10 + WZ_TRACKDETAILS_IMAGE_SIZE + 10
    maxWidthOfAlbumLabel = self.width() - albumInfoXPosStart - d.outerHorizontalPadding

    d.psSelectedTrackAlbumName.setGeometry albumInfoXPosStart, 14, maxWidthOfAlbumLabel, 20
    d.psSelectedTrackAlbumDate.setGeometry albumInfoXPosStart, 14 + 13, maxWidthOfAlbumLabel, 20

    descriptionContainerY0 = 12 + 13 + 18
    descriptionContainerHeight = self.height() - descriptionContainerY0 - 12
    d.descriptionContainer.setGeometry albumInfoXPosStart, descriptionContainerY0, maxWidthOfAlbumLabel, descriptionContainerHeight
End Sub

Private Sub TrackDetailsForm_initialize(ByVal self As WIDGET)
    Dim d As TrackDetailsForm, albumInfoXPosStart As Long, maxWidthOfAlbumLabel As Long
    Set d = TrackDetailsForm_(self)

    Set d.psNowPlaying = W_LABEL_new()
    self.attach d.psNowPlaying
    d.psNowPlaying.setGeometry d.outerHorizontalPadding, 14, 210, 20
    W_LABEL_setFont d.psNowPlaying, font_regular, psPalette(WZCOL_TEXT_MEDIUM)
    W_LABEL_setTextAlignment d.psNowPlaying, WLAB_ALIGNTOPLEFT
    d.psNowPlaying.setTransparentToMouse True

    Set d.psSelectedTrackName = W_LABEL_new()
    self.attach d.psSelectedTrackName
    d.psSelectedTrackName.setGeometry d.outerHorizontalPadding, 14 + 20, 210, 20
    W_LABEL_setFont d.psSelectedTrackName, font_regular_bold, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setTextAlignment d.psSelectedTrackName, WLAB_ALIGNTOPLEFT
    W_LABEL_setCanTruncate d.psSelectedTrackName, True
    d.psSelectedTrackName.setTransparentToMouse True

    ' Add Selected Track author details
    Set d.psSelectedTrackAuthorName = W_LABEL_new()
    self.attach d.psSelectedTrackAuthorName
    d.psSelectedTrackAuthorName.setGeometry d.outerHorizontalPadding, 14 + 20 + 17, 210, 20
    W_LABEL_setFont d.psSelectedTrackAuthorName, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setTextAlignment d.psSelectedTrackAuthorName, WLAB_ALIGNTOPLEFT
    W_LABEL_setCanTruncate d.psSelectedTrackAuthorName, True
    d.psSelectedTrackAuthorName.setTransparentToMouse True

    ' album info xPosStart
    albumInfoXPosStart = d.psNowPlaying.x() + d.psNowPlaying.width() + 10 + WZ_TRACKDETAILS_IMAGE_SIZE + 10
    maxWidthOfAlbumLabel = TL_W - albumInfoXPosStart - d.outerHorizontalPadding

    Set d.psSelectedTrackAlbumName = W_LABEL_new()
    self.attach d.psSelectedTrackAlbumName
    d.psSelectedTrackAlbumName.setGeometry albumInfoXPosStart, 14, maxWidthOfAlbumLabel, 20
    W_LABEL_setFont d.psSelectedTrackAlbumName, font_small, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setTextAlignment d.psSelectedTrackAlbumName, WLAB_ALIGNTOPLEFT
    W_LABEL_setCanTruncate d.psSelectedTrackAlbumName, True
    d.psSelectedTrackAlbumName.setTransparentToMouse True

    Set d.psSelectedTrackAlbumDate = W_LABEL_new()
    self.attach d.psSelectedTrackAlbumDate
    d.psSelectedTrackAlbumDate.setGeometry albumInfoXPosStart, 14 + 13, maxWidthOfAlbumLabel, 20
    W_LABEL_setFont d.psSelectedTrackAlbumDate, font_small, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setTextAlignment d.psSelectedTrackAlbumDate, WLAB_ALIGNTOPLEFT
    W_LABEL_setCanTruncate d.psSelectedTrackAlbumDate, True
    d.psSelectedTrackAlbumDate.setTransparentToMouse True

    Set d.descriptionContainer = ScrollableListWidget_make()
    self.attach d.descriptionContainer
    d.descriptionContainer.setGeometry albumInfoXPosStart, 14 + 13 + 18, maxWidthOfAlbumLabel, 20
End Sub

Public Sub TrackDetailsForm_displayTrack(ByVal self As WIDGET, ByVal track As WZ_TRACK)
    Dim d As TrackDetailsForm, pAlbum As WZ_ALBUM, psSelectedTrackAlbumDescription As WIDGET
    Set d = TrackDetailsForm_(self)
    If Not track Is Nothing Then
        d.psNowPlaying.setString "NOW PLAYING" & ":"
        d.psSelectedTrackName.setString track.title
        d.psSelectedTrackAuthorName.setString track.author
        Set pAlbum = track.album           ' track->album.lock()
    Else
        d.psNowPlaying.setString ""
        d.psSelectedTrackName.setString ""
        d.psSelectedTrackAuthorName.setString ""
        Set pAlbum = Nothing
    End If

    If Not pAlbum Is Nothing Then
        d.psSelectedTrackAlbumName.setString "Album" & ": " & pAlbum.title
        d.psSelectedTrackAlbumDate.setString "Date" & ": " & pAlbum.date_
    Else
        d.psSelectedTrackAlbumName.setString ""
        d.psSelectedTrackAlbumDate.setString ""
    End If

    ScrollableListWidget_clear d.descriptionContainer
    If Not pAlbum Is Nothing Then
        Set psSelectedTrackAlbumDescription = Paragraph_new()
        psSelectedTrackAlbumDescription.setGeometry 0, 0, d.descriptionContainer.width(), 20
        Paragraph_setFont psSelectedTrackAlbumDescription, font_small
        Paragraph_setFontColour psSelectedTrackAlbumDescription, psPalette(WZCOL_TEXT_MEDIUM)
        psSelectedTrackAlbumDescription.setTransparentToMouse True
        Paragraph_addText psSelectedTrackAlbumDescription, pAlbum.description
        ScrollableListWidget_addItem d.descriptionContainer, psSelectedTrackAlbumDescription
    End If

    If Not pAlbum Is Nothing Then
        TrackDetailsForm_loadImage self, pAlbum.album_cover_filename
    Else
        TrackDetailsForm_loadImage self, ""
    End If
End Sub

' MARK: - Track List

' W_MusicListHeaderColImage(): W_BUTTON()
Public Sub W_MusicListHeaderColImage_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_W_MusicListHeaderColImage ' (VB) virtual calls from here on reach W_MusicListHeaderColImage's overrides
    Set W_MusicListHeaderColImage_(self).pAlbumCoverTexture = Nothing
    Set W_BUTTON_(self).AudioCallback = Nothing
End Sub

' std::make_shared<W_MusicListHeaderColImage>()
Public Function W_MusicListHeaderColImage_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_MusicListHeaderColImage)
    W_MusicListHeaderColImage_construct w
    Set W_MusicListHeaderColImage_new = w
End Function

' virtual ~W_MusicListHeaderColImage()
Public Sub W_MusicListHeaderColImage_destruct(ByVal self As WIDGET)
    Dim d As W_MusicListHeaderColImage
    Set d = W_MusicListHeaderColImage_(self)
    If Not d.pAlbumCoverTexture Is Nothing Then
        Set d.pAlbumCoverTexture = Nothing  ' delete pAlbumCoverTexture
    End If
End Sub

Public Function W_MusicListHeaderColImage_loadImage(ByVal self As WIDGET, ByVal imagePath As String) As Boolean
    Dim d As W_MusicListHeaderColImage
    Set d = W_MusicListHeaderColImage_(self)
    If Not d.pAlbumCoverTexture Is Nothing Then
        Set d.pAlbumCoverTexture = Nothing
    End If
    Set d.pAlbumCoverTexture = loadImageToTexture(imagePath)
    W_MusicListHeaderColImage_loadImage = Not d.pAlbumCoverTexture Is Nothing
End Function

Public Sub W_MusicListHeaderColImage_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_MusicListHeaderColImage, imageLeft As Long, imageTop As Long
    Set d = W_MusicListHeaderColImage_(self)
    If Not d.pAlbumCoverTexture Is Nothing Then
        imageLeft = self.x() + xOffset
        imageTop = self.y() + yOffset

        iV_DrawImageAnisotropic d.pAlbumCoverTexture, imageLeft, imageTop, 0!, 0!, W_TRACK_HEADER_COL_IMAGE_SIZE, W_TRACK_HEADER_COL_IMAGE_SIZE, 0!, psPalette(WZCOL_WHITE)
    End If
End Sub

Public Function W_MusicListHeaderColImage_idealWidth(ByVal self As WIDGET) As Long
    W_MusicListHeaderColImage_idealWidth = 16
End Function

' static const std::string music_mode_col_header_images[]
Private Function music_mode_col_header_images(ByVal i As Long) As String
    Select Case i
    Case 0: music_mode_col_header_images = "images/frontend/image_music_campaign.png"
    Case 1: music_mode_col_header_images = "images/frontend/image_music_challenges.png"
    Case 2: music_mode_col_header_images = "images/frontend/image_music_skirmish.png"
    Case 3: music_mode_col_header_images = "images/frontend/image_music_multiplayer.png"
    End Select
End Function

' MARK: - MusicManagerForm

' explicit MusicManagerForm(bool ingame): W_FORM(), ingame(ingame)
Public Sub MusicManagerForm_construct(ByVal self As WIDGET, ByVal isInGame As Boolean)
    Dim d As MusicManagerForm
    W_FORM_construct self
    self.kind = K_MusicManagerForm ' (VB) virtual calls from here on reach MusicManagerForm's overrides
    Set d = MusicManagerForm_(self)
    Set d.trackList = New Collection
    d.minDesiredTableWidth = 0
    d.ingame_ = isInGame
    d.tablePaddingX = 6
    d.weakSerial = WeakWidget_register(self)    ' (VB) its std::weak_ptrs
End Sub

Private Sub MusicManagerForm_initialize(ByVal self As WIDGET)
    Dim d As MusicManagerForm, i As Long, sink As MusicManager_CDAudioEventSink
    Set d = MusicManagerForm_(self)
    ' get track list
    Set d.trackList = New Collection
    For i = 0 To PlayList_GetFullTrackList_size() - 1
        d.trackList.Add PlayList_GetFullTrackList_at(i)
    Next
    Set d.selectedTrack = cdAudio_GetCurrentTrack()

    MusicManagerForm_addTrackList self, self
    MusicManagerForm_addTrackDetailsBox self

    MusicManagerForm_updateActiveTrackImpl self, d.selectedTrack

    ' register for cd audio events
    Set sink = New MusicManager_CDAudioEventSink
    sink.construct d.weakSerial
    Set d.musicManagerAudioEventSink = sink
    cdAudio_RegisterForEvents d.musicManagerAudioEventSink
End Sub

Public Function MusicManagerForm_make(ByVal isInGame As Boolean) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_MusicManagerForm)
    MusicManagerForm_construct widget, isInGame
    MusicManagerForm_initialize widget
    Set MusicManagerForm_make = widget
End Function

' MusicManagerForm::~MusicManagerForm()
Public Sub MusicManagerForm_destruct(ByVal self As WIDGET)
    Dim d As MusicManagerForm
    Set d = MusicManagerForm_(self)
    WeakWidget_unregister d.weakSerial
    If Not d.musicManagerAudioEventSink Is Nothing Then
        d.musicManagerAudioEventSink.setUnregisterEventSink
        Set d.musicManagerAudioEventSink = Nothing
    End If
    If Not d.ingame_ Then
        cdAudio_PlayTrack cdAudio_CurrentSongContext()
    End If
End Sub

Public Sub MusicManagerForm_geometryChanged(ByVal self As WIDGET)
    Dim d As MusicManagerForm, w As Long, h As Long, trackDetailsBoxHeight As Long, trackDetailsBoxY0 As Long
    Dim tracksTableX0 As Long, tracksTableY0 As Long, tracksTableHeight As Long, tracksTableWidth As Long
    Dim minimumColumnWidths() As Long, minimumColumnWidthsCount As Long
    Set d = MusicManagerForm_(self)
    w = self.width()
    h = self.height()

    If w = 0 Or h = 0 Then
        Exit Sub
    End If

    trackDetailsBoxHeight = TL_PREVIEWBOX_H
    trackDetailsBoxY0 = h - trackDetailsBoxHeight
    d.trackDetailsBox.setGeometry 0, trackDetailsBoxY0, w, trackDetailsBoxHeight

    tracksTableX0 = d.tablePaddingX
    tracksTableY0 = 0
    tracksTableHeight = trackDetailsBoxY0 - tracksTableY0 - 5
    tracksTableWidth = w - (d.tablePaddingX * 2)
    d.table.setGeometry tracksTableX0, tracksTableY0, tracksTableWidth, tracksTableHeight

    minimumColumnWidthsCount = ScrollableTableWidget_getMinimumColumnWidths(d.table, minimumColumnWidths)
    ScrollableTableWidget_setMinimumColumnWidths d.table, minimumColumnWidths, minimumColumnWidthsCount
End Sub

Public Function MusicManagerForm_idealWidth(ByVal self As WIDGET) As Long
    Dim d As MusicManagerForm
    Set d = MusicManagerForm_(self)
    MusicManagerForm_idealWidth = MAX_(d.table.idealWidth(), (d.minDesiredTableWidth + (d.tablePaddingX * 2) + 100))
End Function

Public Function MusicManagerForm_idealHeight(ByVal self As WIDGET) As Long
    MusicManagerForm_idealHeight = MusicManagerForm_(self).trackDetailsBox.idealHeight() + 400
End Function

Public Sub MusicManagerForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    ' currently, no-op
End Sub

Public Sub MusicManagerForm_updateActiveTrack(ByVal self As WIDGET, ByVal track As WZ_TRACK)
    If track Is MusicManagerForm_(self).selectedTrack Then
        Exit Sub
    End If
    MusicManagerForm_updateActiveTrackImpl self, track
End Sub

Private Sub MusicManagerForm_updateActiveTrackImpl(ByVal self As WIDGET, ByVal track As WZ_TRACK)
    Dim d As MusicManagerForm, trackRowIndex As Long, i As Long, tableRows As Collection, row As WIDGET
    Set d = MusicManagerForm_(self)
    Set d.selectedTrack = track

    TrackDetailsForm_displayTrack d.trackDetailsBox, track

    ' auto it = std::find(trackList.begin(), trackList.end(), track)
    trackRowIndex = -1
    For i = 1 To d.trackList.Count
        If d.trackList(i) Is track Then
            trackRowIndex = i - 1
            Exit For
        End If
    Next
    If trackRowIndex = -1 Then
        Exit Sub
    End If
    Set tableRows = ScrollableTableWidget_getRows(d.table)
    If Not (trackRowIndex < tableRows.Count) Then ASSERT_FAILURE "trackRowIndex < tableRows.size()", "src/musicmanager.cpp:601", "updateActiveTrackImpl", "Invalid row index: %zu", trackRowIndex
    For i = 0 To tableRows.Count - 1
        Set row = tableRows(i + 1)
        If i = trackRowIndex Then
            TableRow_setBackgroundColor row, True, psPalette(WZCOL_KEYMAP_ACTIVE)
            TableRow_setHighlightsOnMouseOver row, False
            GoTo continue_
        End If
        TableRow_setBackgroundColor row, False     ' nullopt
        TableRow_setHighlightsOnMouseOver row, True
continue_:
    Next
End Sub

Private Function MusicManagerForm_createColumnHeaderLabel(ByVal self As WIDGET, ByVal text As String, ByVal minLeftPadding As Long) As WIDGET
    Dim label As WIDGET
    Set label = W_LABEL_new()
    W_LABEL_setTextAlignment label, WLAB_ALIGNTOPLEFT
    W_LABEL_setFontColour label, psPalette(WZCOL_TEXT_MEDIUM)
    label.setString text
    label.setGeometry 0, 0, W_LABEL_getMaxLineWidth(label), 0
    W_LABEL_setCacheNeverExpires label, True
    W_LABEL_setCanTruncate label, True
    Set MusicManagerForm_createColumnHeaderLabel = Margin_wrap(W_TRACK_ROW_PADDING + (W_TRACK_ROW_PADDING \ 2), 0, 0, minLeftPadding, label)
End Function

Private Sub MusicManagerForm_addTrackList(ByVal self As WIDGET, ByVal parent As WIDGET)
    Dim d As MusicManagerForm, columns() As TableColumn, columnsCount As Long, musicModeIdx As Long
    Dim headerImage As String, musicModeHeader As WIDGET, musicGameModeStr As String, maxIdealColumnHeight As Long
    Dim minimumColumnWidths() As Long, minimumColumnWidthsCount As Long, i As Long, weakSelf As Long
    Dim tracksCount As Long, trackIdx As Long, trackRow As WIDGET
    Set d = MusicManagerForm_(self)

    ' Create column headers for tracks listing
    ' (std::vector<TableColumn> columns: two, then one per music mode)
    ReDim columns(0 To 1 + NUM_MUSICGAMEMODES)
    columns(0) = TableColumn_(MusicManagerForm_createColumnHeaderLabel(self, "Title", 0), TableColumn_ResizeBehavior_RESIZABLE_AUTOEXPAND)
    columns(1) = TableColumn_(MusicManagerForm_createColumnHeaderLabel(self, "Album", 0), TableColumn_ResizeBehavior_RESIZABLE_AUTOEXPAND)
    columnsCount = 2

    For musicModeIdx = 0 To NUM_MUSICGAMEMODES - 1
        headerImage = music_mode_col_header_images(musicModeIdx)
        Set musicModeHeader = W_MusicListHeaderColImage_new()
        W_MusicListHeaderColImage_loadImage musicModeHeader, headerImage
        musicModeHeader.setGeometry 0, 0, W_TRACK_HEADER_COL_IMAGE_SIZE, W_TRACK_HEADER_HEIGHT
        musicGameModeStr = MusicGameMode_to_string(musicModeIdx)
        musicModeHeader.setTip musicGameModeStr
        columns(columnsCount) = TableColumn_(Margin_wrap(W_TRACK_ROW_PADDING + (W_TRACK_ROW_PADDING \ 2), 0, 0, 0, musicModeHeader), TableColumn_ResizeBehavior_FIXED_WIDTH)
        columnsCount = columnsCount + 1
    Next

    maxIdealColumnHeight = W_TRACK_HEADER_HEIGHT
    ' std::vector<size_t> minimumColumnWidths = { W_TRACK_COL_TITLE_W, W_TRACK_COL_ALBUM_W }
    ReDim minimumColumnWidths(0 To columnsCount - 1)
    minimumColumnWidths(0) = W_TRACK_COL_TITLE_W
    minimumColumnWidths(1) = W_TRACK_COL_ALBUM_W
    minimumColumnWidthsCount = 2
    For i = 2 To columnsCount - 1
        minimumColumnWidths(minimumColumnWidthsCount) = MAX_(columns(i).columnWidget.idealWidth(), 0)
        minimumColumnWidthsCount = minimumColumnWidthsCount + 1
    Next
    ' std::accumulate(minimumColumnWidths.begin(), minimumColumnWidths.end(), 0)
    d.minDesiredTableWidth = 0
    For i = 0 To minimumColumnWidthsCount - 1
        d.minDesiredTableWidth = d.minDesiredTableWidth + minimumColumnWidths(i)
    Next
    d.minDesiredTableWidth = d.minDesiredTableWidth + (W_TRACK_COL_PADDING * 2) + (W_TRACK_COL_PADDING * 2 * minimumColumnWidthsCount)

    ' Create table
    Set d.table = ScrollableTableWidget_make(columns, columnsCount, maxIdealColumnHeight)
    ScrollableTableWidget_setColumnPadding d.table, Vector2i_(W_TRACK_COL_PADDING, 0)
    parent.attach d.table
    ScrollableTableWidget_setBackgroundColor d.table, pal_RGBA(0, 0, 0, 0)
    ScrollableTableWidget_setMinimumColumnWidths d.table, minimumColumnWidths, minimumColumnWidthsCount
    ScrollableTableWidget_setDrawColumnLines d.table, False

    ' Add tracks
    weakSelf = d.weakSerial
    tracksCount = d.trackList.Count
    For trackIdx = 0 To tracksCount - 1
        Set trackRow = MusicManagerForm_createTrackRow(self, d.trackList(trackIdx + 1))
        TableRow_setHighlightsOnMouseOver trackRow, True
        trackRow.setGeometry 0, 0, TL_ENTRYW, TL_ENTRYH
        W_BUTTON_addOnClickHandler trackRow, WzFunc_new(FN_musicManager_trackRowClick, weakSelf, trackIdx)
        ScrollableTableWidget_addRow d.table, trackRow
    Next
End Sub

Private Function MusicManagerForm_createTrackRow(ByVal self As WIDGET, ByVal track As WZ_TRACK) As WIDGET
    Dim d As MusicManagerForm, columnWidgets As Collection, album As WZ_ALBUM, album_name As String
    Dim trackTitleLabel As WIDGET, albumTitleLabel As WIDGET, musicMode As MusicGameMode, musicModeIdx As Long
    Dim pCheckBox As WIDGET, captureTrack As WZ_TRACK, rowHeight As Long, row As WIDGET
    Set d = MusicManagerForm_(self)
    Set columnWidgets = New Collection  ' std::vector<std::shared_ptr<WIDGET>> columnWidgets;

    Set album = track.album             ' track->album.lock()
    album_name = album.title

    ' Track Title
    Set trackTitleLabel = W_LABEL_new()
    W_LABEL_setFont trackTitleLabel, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    trackTitleLabel.setString track.title
    trackTitleLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(trackTitleLabel), W_LABEL_requiredHeight(trackTitleLabel)
    trackTitleLabel.setTransparentToMouse True
    W_LABEL_setCanTruncate trackTitleLabel, True
    columnWidgets.Add trackTitleLabel

    ' Album Title
    Set albumTitleLabel = W_LABEL_new()
    W_LABEL_setFont albumTitleLabel, font_regular, psPalette(WZCOL_TEXT_BRIGHT)
    albumTitleLabel.setString album_name
    albumTitleLabel.setGeometry 0, 0, W_LABEL_getMaxLineWidth(albumTitleLabel), W_LABEL_requiredHeight(albumTitleLabel)
    albumTitleLabel.setTransparentToMouse True
    W_LABEL_setCanTruncate albumTitleLabel, True
    columnWidgets.Add albumTitleLabel

    ' add music mode checkboxes
    musicMode = MusicGameMode_MENUS
    If d.ingame_ Then
        musicMode = PlayList_GetCurrentMusicMode()
    End If

    For musicModeIdx = 0 To NUM_MUSICGAMEMODES - 1
        Set pCheckBox = W_MusicModeCheckboxButton_new(musicModeIdx, PlayList_IsTrackEnabledForMusicMode(track, musicModeIdx))
        Set captureTrack = track
        W_BUTTON_addOnClickHandler pCheckBox, WzFunc_new(FN_musicManager_setTrackMode, captureTrack)
        pCheckBox.setGeometry 0, 0, W_TRACK_CHECKBOX_SIZE, W_TRACK_CHECKBOX_SIZE
        W_MusicModeCheckboxButton_setCheckboxSize pCheckBox, W_TRACK_CHECKBOX_SIZE
        If musicMode <> MusicGameMode_MENUS Then
            If musicModeIdx <> musicMode Then
                pCheckBox.setState WBUT_DISABLE
            End If
        End If
        columnWidgets.Add pCheckBox
    Next

    rowHeight = TL_ENTRYH

    Set row = TableRow_make(columnWidgets, rowHeight)
    TableRow_setDisabledColor row, psPalette(WZCOL_FORM_DARK)
    Set MusicManagerForm_createTrackRow = row
End Function

Private Sub MusicManagerForm_addTrackDetailsBox(ByVal self As WIDGET)
    Dim d As MusicManagerForm
    Set d = MusicManagerForm_(self)
    Set d.trackDetailsBox = TrackDetailsForm_make()
    self.attach d.trackDetailsBox
End Sub

Public Function makeMusicManagerForm(ByVal isInGame As Boolean) As WIDGET
    Set makeMusicManagerForm = MusicManagerForm_make(isInGame)
End Function

'--------------------------------------------------------------------------
' the lambdas

Public Function musicmanager_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim button As WIDGET, d As W_MusicModeCheckboxButton, strongSelf As WIDGET, trackIdx As Long
    Dim captureTrack As WZ_TRACK
    Select Case f.id
    Case FN_musicModeCheckbox_onClick           ' [](W_BUTTON& button)
        Set button = args(0)
        Set d = W_MusicModeCheckboxButton_(button)   ' dynamic_cast<W_MusicModeCheckboxButton&>(button)
        If Not W_MusicModeCheckboxButton_isEnabled(button) Then Exit Function
        d.isChecked = Not d.isChecked

    Case FN_musicManager_trackRowClick          ' [weakSelf, trackIdx](W_BUTTON& clickedButton)
        Set strongSelf = WeakWidget_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/musicmanager.cpp:673", "addTrackList", "Widget gone?": Exit Function
        trackIdx = f.cap(1)
        If MusicManagerForm_(strongSelf).selectedTrack Is MusicManagerForm_(strongSelf).trackList(trackIdx + 1) Then Exit Function
        cdAudio_PlaySpecificTrack MusicManagerForm_(strongSelf).trackList(trackIdx + 1)

    Case FN_musicManager_setTrackMode           ' [captureTrack](W_BUTTON& button)
        Set button = args(0)
        Set captureTrack = f.cap(0)
        PlayList_SetTrackMusicMode captureTrack, W_MusicModeCheckboxButton_getMusicMode(button), W_MusicModeCheckboxButton_getIsChecked(button)
    End Select
End Function
