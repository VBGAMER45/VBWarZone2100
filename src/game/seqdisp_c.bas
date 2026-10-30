Attribute VB_Name = "seqdisp_c"
' src/seqdisp.cpp, seqdisp.h: the full screen videos' list, their text overlays and subtitles, the
' intelligence screen's videos. The video player (lib/sequence: seq_Play, seq_Update, seq_Shutdown,
' seq_GetFrameTime, seq_SetDisplaySize, makeVideoProvider) is not ported: its functions are stand-ins
' (stubs_ui) that play nothing, as an install without the sequences, whose paths run. A
' std::shared_ptr<VideoProvider> is a Long (0 = nullptr). The on-demand downloader's URL requests (curl,
' urlrequest.cpp) are not ported either: urlRequestData (stubs_ui) returns no handle, so a request fails as
' when it cannot be made. std::vector<WzText> wzCachedSeqText is a Collection of WzText.
'
' Functions for the display of the Escape Sequences (FMV).
Option Explicit

Public Const SEQUENCE_KILL As Long = 3  ' stop
Public Const SEQUENCE_HOLD As Long = 4  ' play once and hold last frame

Public Enum SEQ_TEXT_POSITIONING
    ' Position text.
    SEQ_TEXT_POSITION
    ' Justify if less than 520/600 length.
    SEQ_TEXT_JUSTIFY
End Enum

'**************************************************************************
'
'	local Definitions
'
'**************************************************************************
Private Const MAX_TEXT_OVERLAYS As Long = 32
Private Const MAX_SEQ_LIST As Long = 6
Private Const SUBTITLE_BOX_MIN As Long = 430
Private Const SUBTITLE_BOX_MAX As Long = 480

' NOTE: The original game never had a true fullscreen mode for FMVs on >640x480 screens.
' They would just use double sized videos, and move the text to that area.
' Since we *do* offer fullscreen FMVs, this isn't really needed anymore, depending
' on how we want to handle the text.
Private D_W2 As Long    ' = 0: Text width offset
Private D_H2 As Long    ' = 0: Text height offset

Private Type SEQTEXT
    pText As String                     ' char pText[MAX_STR_LENGTH] ("" = '\0' first)
    x As Long                           ' UDWORD
    y As Long
    startTime As Double
    endTime As Double
    bSubtitle As Boolean
End Type

Private Type SEQLIST
    pSeq As String                      ' WzString: name of the sequence to play
    pAudio As String                    ' name of the wav to play
    bSeqLoop As Boolean                 ' loop this sequence
    currentText As Long                 ' current number of text messages for this seq
    aText(0 To MAX_TEXT_OVERLAYS - 1) As SEQTEXT   ' text data to display for this sequence
End Type

' enum VIDEO_RESOLUTION
Private Const VIDEO_PRESELECTED_RESOLUTION As Long = 0
Private Const VIDEO_USER_CHOSEN_RESOLUTION As Long = 1

' OnDemandVideoDownloader::RequestDetails::RequestStatus
Private Const RequestStatus_Pending As Long = 0
Private Const RequestStatus_Success As Long = 1
Private Const RequestStatus_Failure As Long = 2

'**************************************************************************
'
'	local Variables
'
'**************************************************************************

Private bAudioPlaying As Boolean        ' = false
Private bHoldSeqForAudio As Boolean     ' = false
Private bSeqSubtitles As Boolean        ' = true (seqdisp_c_init)
Private bSeqPlaying As Boolean          ' = false
Private currVideoName As String
Private currAudioName As String
Private aVideoProvider As Long          ' std::shared_ptr<VideoProvider> (0 = nullptr)
Private aSeqList(0 To MAX_SEQ_LIST - 1) As SEQLIST
Private currentSeq As Long              ' = -1 (seqdisp_c_init)
Private currentPlaySeq As Long          ' = -1 (seqdisp_c_init)

' local rendered text cache
Private wzCachedSeqText As New Collection   ' std::vector<WzText>

' static OnDemandVideoDownloader onDemandVideoProvider: its priorRequests (an unordered_map of the videos'
' names to their RequestDetails: the arrays below, rd_count of them) and its optional<WzString> baseURLPath
Private rd_name() As String
Private rd_status() As Long             ' RequestDetails::status = Pending
Private rd_progressPercentage() As Long ' uint32_t = 0
Private rd_requestHandle() As Long      ' AsyncURLRequestHandle (0: none)
Private rd_count As Long
Private baseURLPath As String
Private has_baseURLPath As Boolean

' seq_RenderVideoToBuffer's statics
Private Const VIDEO_NOT_PLAYING As Long = 0
Private Const VIDEO_PLAYING As Long = 1
Private Const VIDEO_FINISHED As Long = 2
Private Const VIDEO_LOOP As Long = 0
Private Const VIDEO_HOLD_LAST_FRAME As Long = 1
Private videoPlaying As Long            ' = VIDEO_NOT_PLAYING
Private frameHold As Long               ' = VIDEO_LOOP

' seq_AddTextForVideo's static SDWORD lastX
Private lastX As Long

' update_user_resolution's static glm::uvec2 previousScreenResolution = {0, 0}
Private previousScreenResolution_x As Long
Private previousScreenResolution_y As Long

' (VB) the statics' initialisers (exemain_vb calls it)
Public Sub seqdisp_c_init()
    bSeqSubtitles = True
    currentSeq = -1
    currentPlaySeq = -1
End Sub

' SEQLIST::reset()
Private Sub SEQLIST_reset(l As SEQLIST)
    Dim i As Long, z As SEQTEXT
    l.bSeqLoop = False
    l.currentText = 0
    For i = 0 To MAX_TEXT_OVERLAYS - 1          ' memset(aText, 0, sizeof(aText));
        l.aText(i) = z
    Next
    l.pSeq = ""
    l.pAudio = ""
End Sub

' (VB) priorRequests.find(videoName): its index, -1 when not found (end())
Private Function priorRequests_find(ByVal videoName As String) As Long
    Dim i As Long
    For i = 0 To rd_count - 1
        If rd_name(i) = videoName Then
            priorRequests_find = i
            Exit Function
        End If
    Next
    priorRequests_find = -1
End Function

Private Function urlEncodeVideoPathComponents(ByVal partialPath As String) As String
    Dim pathComponents() As String, i As Long, result As String
    ' Divide up partialPath into path components (by /)
    pathComponents = Split(partialPath, "/")
    ' Then urlEncode each path component, and recombine
    result = ""
    For i = 0 To UBound(pathComponents)
        If i > 0 Then
            result = result & "/"
        End If
        result = result & urlEncode(pathComponents(i))
    Next
    urlEncodeVideoPathComponents = result
End Function

' bool OnDemandVideoDownloader::requestVideoData(const WzString& videoName)
Private Function OnDemandVideoDownloader_requestVideoData(ByVal videoName As String) As Boolean
    Dim it As Long, requestDetails As Long, url As String, requestHandle As Long
    it = priorRequests_find(videoName)
    If it <> -1 Then
        ' already requested
        OnDemandVideoDownloader_requestVideoData = rd_status(it) <> RequestStatus_Failure
        Exit Function
    End If

    ' need to request

    ' auto requestDetails = std::make_shared<RequestDetails>(); priorRequests[videoName] = requestDetails;
    If rd_count = 0 Then
        ReDim rd_name(0 To 0): ReDim rd_status(0 To 0): ReDim rd_progressPercentage(0 To 0): ReDim rd_requestHandle(0 To 0)
    Else
        ReDim Preserve rd_name(0 To rd_count): ReDim Preserve rd_status(0 To rd_count)
        ReDim Preserve rd_progressPercentage(0 To rd_count): ReDim Preserve rd_requestHandle(0 To rd_count)
    End If
    requestDetails = rd_count
    rd_name(requestDetails) = videoName
    rd_status(requestDetails) = RequestStatus_Pending
    rd_progressPercentage(requestDetails) = 0
    rd_requestHandle(requestDetails) = 0
    rd_count = rd_count + 1

    If Not has_baseURLPath Then
        rd_status(requestDetails) = RequestStatus_Failure
        OnDemandVideoDownloader_requestVideoData = False
        Exit Function
    End If

    url = baseURLPath & urlEncodeVideoPathComponents(videoName)
    ' (urlRequest.onResponse, onFailure, progressCallback and maxDownloadSizeLimit = 200 MB: the URL requests
    ' are not ported, urlRequestData returns no handle)
    requestHandle = urlRequestData(url)
    If requestHandle = 0 Then
        debug_ 278, LOG_ERROR, "requestVideoData", "Failed to request data: %s", url
        rd_status(requestDetails) = RequestStatus_Failure
        OnDemandVideoDownloader_requestVideoData = False
        Exit Function
    End If

    rd_requestHandle(requestDetails) = requestHandle

    OnDemandVideoDownloader_requestVideoData = True
End Function

' std::shared_ptr<const std::vector<char>> OnDemandVideoDownloader::getVideoData(const WzString& videoName):
' (VB) False for nullptr (no request succeeds before the URL requests are ported: no data to return)
Private Function OnDemandVideoDownloader_getVideoData(ByVal videoName As String) As Boolean
    Dim it As Long
    it = priorRequests_find(videoName)
    If it <> -1 Then
        ' already requested
        If rd_status(it) = RequestStatus_Success Then
            OnDemandVideoDownloader_getVideoData = True     ' return it->second->responseData;
            Exit Function
        Else
            OnDemandVideoDownloader_getVideoData = False
            Exit Function
        End If
    End If

    OnDemandVideoDownloader_requestVideoData videoName
    OnDemandVideoDownloader_getVideoData = False
End Function

Private Function OnDemandVideoDownloader_getVideoDataRequestProgress(ByVal videoName As String) As Long
    Dim it As Long
    it = priorRequests_find(videoName)
    If it <> -1 Then
        ' already requested
        Select Case rd_status(it)
        Case RequestStatus_Pending
            OnDemandVideoDownloader_getVideoDataRequestProgress = rd_progressPercentage(it)
        Case RequestStatus_Success
            OnDemandVideoDownloader_getVideoDataRequestProgress = 100
        Case Else
            OnDemandVideoDownloader_getVideoDataRequestProgress = 0
        End Select
        Exit Function
    End If
    OnDemandVideoDownloader_getVideoDataRequestProgress = 0
End Function

Private Function OnDemandVideoDownloader_getVideoDataError(ByVal videoName As String) As Boolean
    Dim it As Long
    it = priorRequests_find(videoName)
    If it <> -1 Then
        OnDemandVideoDownloader_getVideoDataError = rd_status(it) = RequestStatus_Failure
        Exit Function
    End If
    OnDemandVideoDownloader_getVideoDataError = False
End Function

' Cancel video download request
Private Function OnDemandVideoDownloader_cancelDownloadRequest(ByVal videoName As String) As Boolean
    Dim it As Long, i As Long
    it = priorRequests_find(videoName)
    If it = -1 Then
        OnDemandVideoDownloader_cancelDownloadRequest = False
        Exit Function
    End If
    If rd_requestHandle(it) <> 0 Then
        urlRequestSetCancelFlag rd_requestHandle(it)
    End If
    ' priorRequests.erase(it);
    For i = it To rd_count - 2
        rd_name(i) = rd_name(i + 1)
        rd_status(i) = rd_status(i + 1)
        rd_progressPercentage(i) = rd_progressPercentage(i + 1)
        rd_requestHandle(i) = rd_requestHandle(i + 1)
    Next
    rd_count = rd_count - 1
    OnDemandVideoDownloader_cancelDownloadRequest = True
End Function

Private Sub OnDemandVideoDownloader_clear()
    rd_count = 0                        ' priorRequests.clear();
    Erase rd_name, rd_status, rd_progressPercentage, rd_requestHandle
End Sub

Private Sub OnDemandVideoDownloader_setBaseURLPath(ByVal basePath As String)
    baseURLPath = basePath
    has_baseURLPath = True
End Sub

Private Function OnDemandVideoDownloader_hasBaseURLPath() As Boolean
    OnDemandVideoDownloader_hasBaseURLPath = has_baseURLPath
End Function

'**************************************************************************
'
'	Source
'
'**************************************************************************

Public Function seq_hasVideos() As Boolean
    seq_hasVideos = PHYSFS_exists("sequences/devastation.ogg") <> 0 Or OnDemandVideoDownloader_hasBaseURLPath()
End Function

Public Sub seq_setOnDemandVideoURL(ByVal videoBaseURL As String)
    OnDemandVideoDownloader_setBaseURLPath videoBaseURL
End Sub

' Renders a video sequence specified by filename to a buffer
Public Function seq_RenderVideoToBuffer(ByVal sequenceName As String, ByVal seqCommand As Long) As Boolean
    If seqCommand = SEQUENCE_KILL Then
        'stop the movie
        seq_Shutdown
        bSeqPlaying = False
        frameHold = VIDEO_LOOP
        videoPlaying = VIDEO_NOT_PLAYING
        seq_RenderVideoToBuffer = True
        Exit Function
    End If

    If Not bSeqPlaying And frameHold = VIDEO_LOOP Then
        'start the ball rolling

        iV_SetTextColour psPalette(WZCOL_TEXT_BRIGHT)

        ' We do *NOT* want to use the user-choosen resolution when we
        ' are doing intelligence videos.
        If seq_StartFullScreenVideo(sequenceName, "", VIDEO_PRESELECTED_RESOLUTION) Then videoPlaying = VIDEO_PLAYING Else videoPlaying = VIDEO_FINISHED
        bSeqPlaying = True
    End If

    If videoPlaying <> VIDEO_FINISHED Then
        If seq_Update() Then videoPlaying = VIDEO_PLAYING Else videoPlaying = VIDEO_FINISHED
    End If

    If videoPlaying = VIDEO_FINISHED Then
        seq_Shutdown
        bSeqPlaying = False
        frameHold = VIDEO_HOLD_LAST_FRAME
        videoPlaying = VIDEO_NOT_PLAYING
        seq_RenderVideoToBuffer = False
        Exit Function
    End If

    seq_RenderVideoToBuffer = True
End Function


Private Sub seq_SetUserResolution()
    Dim video_size_x As Long, video_size_y As Long, scaled_width As Single, sw As Single, sh As Single, x As Long, y As Long

    Select Case war_GetFMVmode()
    Case FMV_1X
        ' Native (1x)
        video_size_x = 320: video_size_y = 240
    Case FMV_2X
        ' Double (2x)
        video_size_x = 640: video_size_y = 480
    Case FMV_FULLSCREEN
        sw = CSng(screenWidth) / 4!
        sh = CSng(screenHeight) / 3!
        If sw < sh Then scaled_width = sw * 4! Else scaled_width = sh * 4!
        video_size_x = Fix(scaled_width): video_size_y = Fix(CSng(scaled_width * 3! / 4!))
    Case Else
        ASSERT_FAILURE "!""invalid FMV mode""", "src/seqdisp.cpp:461", "seq_SetUserResolution", "Invalid FMV mode: %u", war_GetFMVmode()
        Exit Sub
    End Select

    x = (screenWidth - video_size_x) \ 2
    y = (screenHeight - video_size_y) \ 2
    seq_SetDisplaySize video_size_x, video_size_y, x, y
End Sub

Private Function seqPlayOrQueueFetch(ByVal videoName As String, ByVal audioName As String) As Boolean
    Dim aVideoName As String, fpInfile As Long, isNoVideo As Boolean, log_part As code_part, aAudioName As String
    Const maxVolume As Single = 1!
    aVideoProvider = 0                  ' aVideoProvider.reset();
    currVideoName = videoName
    currAudioName = audioName

    ' Try to find local sequences
    aVideoName = "sequences/" & videoName
    fpInfile = PHYSFS_openRead(aVideoName)

    If fpInfile <> 0 Then
        aVideoProvider = makeVideoProvider_file(fpInfile, videoName)
    Else
        isNoVideo = Left$(videoName, 7) = "novideo"

        If Not OnDemandVideoDownloader_hasBaseURLPath() Or isNoVideo Then
            ' no on-demand fallback available - log the failure to open local file
            log_part = LOG_INFO
            If isNoVideo Then
                ' in these special cases, don't clutter the logs with LOG_INFO level events
                log_part = LOG_VIDEO
            End If
            debug_ 497, log_part, "seqPlayOrQueueFetch", "unable to open '%s' for playback", aVideoName

            seq_Shutdown
            seqPlayOrQueueFetch = False
            Exit Function
        End If

        ' Try to download video from on-demand provider
        If Not OnDemandVideoDownloader_getVideoData(videoName) Then
            If OnDemandVideoDownloader_getVideoDataError(videoName) Then
                seq_Shutdown
                seqPlayOrQueueFetch = False
                Exit Function
            End If
            ' Otherwise, request is queued
            ' Return true and handle waiting inside seq_UpdateFullScreenVideo
            ' (Each later call to seq_UpdateFullScreenVideo will check if still waiting, and then dispkay the download progress before itself kicking off seq_Play once the buffer is ready)
            seqPlayOrQueueFetch = True
            Exit Function
        End If

        aVideoProvider = makeVideoProvider_data(videoName)
    End If

    If Not seq_Play(aVideoProvider) Then
        seq_Shutdown
        seqPlayOrQueueFetch = False
        Exit Function
    End If

    If Len(audioName) = 0 Then
        bAudioPlaying = False
    Else
        If Not (Len(audioName) = 0) Then ASSERT_FAILURE "audioName.isEmpty()", "src/seqdisp.cpp:533", "seqPlayOrQueueFetch", "Unexpected audio stream specified: %s (for video: %s)", audioName, videoName

        ' NOT controlled by sliders for now?
        ' (static const float maxVolume = 1.f;)

        aAudioName = "sequenceaudio/" & audioName
        bAudioPlaying = Not audio_PlayStream(aAudioName, maxVolume, Nothing, Empty) Is Nothing
        If Not (bAudioPlaying = True) Then ASSERT_FAILURE "bAudioPlaying == true", "src/seqdisp.cpp:540", "seqPlayOrQueueFetch", "unable to initialise sound %s", aAudioName
    End If

    seqPlayOrQueueFetch = True
End Function

Private Function seqPlayOrQueueFetch_Retry() As Boolean
    seqPlayOrQueueFetch_Retry = seqPlayOrQueueFetch(currVideoName, currAudioName)
End Function

'full screenvideo functions
' "audioName" doesn't appear to be used? (seems to always be set to empty string?) // past-due (2022)
Private Function seq_StartFullScreenVideo(ByVal videoName As String, ByVal audioName As String, ByVal resolution As Long) As Boolean
    bHoldSeqForAudio = False
    iV_SetTextColour psPalette(WZCOL_TEXT_BRIGHT)

    If resolution = VIDEO_USER_CHOSEN_RESOLUTION Then
        ' set the dimensions to show full screen or native or ...
        seq_SetUserResolution
    End If

    If Not seqPlayOrQueueFetch(videoName, audioName) Then
        seq_StartFullScreenVideo = False
        Exit Function
    End If

    ' We do not want to enter loop_SetVideoPlaybackMode() when we are
    ' doing intelligence videos.
    If resolution = VIDEO_USER_CHOSEN_RESOLUTION Then
        'start video mode
        If loop_GetVideoMode() = 0 Then
            ' check to see if we need to pause, and set font each time
            cdAudio_Pause
            loop_SetVideoPlaybackMode
            iV_SetTextColour psPalette(WZCOL_TEXT_BRIGHT)
        End If
    End If

    seq_StartFullScreenVideo = True
End Function

Public Sub update_user_resolution()
    If previousScreenResolution_x <> screenWidth Or previousScreenResolution_y <> screenHeight Then
        seq_SetUserResolution
        previousScreenResolution_x = screenWidth
        previousScreenResolution_y = screenHeight
    End If
End Sub

' (VB) wzCachedSeqText[i], resized to hold it
Private Function cachedSeqText(ByVal i As Long) As WzText
    Do While i >= wzCachedSeqText.Count
        wzCachedSeqText.Add New WzText
    Loop
    Set cachedSeqText = wzCachedSeqText(i + 1)
End Function

Public Function seq_UpdateFullScreenVideo() As Boolean
    Dim i As Long, bMoreThanOneSequenceLine As Boolean, stillPlaying As Boolean, subMin As Long, subMax As Long
    Dim frameTime As Double, seqtext_ As SEQTEXT, currentText As SEQTEXT, t As WzText
    update_user_resolution

    If aVideoProvider = 0 Then
        ' request has been queued
        ' try again and see if it's available
        If Not seqPlayOrQueueFetch_Retry() Then
            seq_UpdateFullScreenVideo = False
            Exit Function
        End If

        If aVideoProvider = 0 Then
            ' still don't have the video data yet, but it's being fetched
            ' display loading message
            Set t = cachedSeqText(1)        ' if (wzCachedSeqText.size() < 2) wzCachedSeqText.resize(2);
            Set t = cachedSeqText(0)
            t.setText va("%s (%u%%)...", "Loading video", OnDemandVideoDownloader_getVideoDataRequestProgress(currVideoName)), font_scaled
            t.render (pie_GetVideoBufferWidth() - t.width()) \ 2, (pie_GetVideoBufferHeight() - t.height()) \ 2, psPalette(WZCOL_WHITE)
            seq_UpdateFullScreenVideo = True
            Exit Function
        End If

        ' otherwise, we have the data - continue
    End If

    bMoreThanOneSequenceLine = False

    subMin = SUBTITLE_BOX_MAX + D_H2
    subMax = SUBTITLE_BOX_MIN + D_H2

    'get any text lines over bottom of the video
    frameTime = seq_GetFrameTime()
    For i = 0 To MAX_TEXT_OVERLAYS - 1
        seqtext_ = aSeqList(currentPlaySeq).aText(i)
        If Len(seqtext_.pText) <> 0 Then
            If seqtext_.bSubtitle Then
                If ((frameTime >= seqtext_.startTime) And (frameTime <= seqtext_.endTime)) Or _
                   aSeqList(currentPlaySeq).bSeqLoop Then 'if its a looped video always draw the text
                    If U32(subMin) > U32(seqtext_.y) And U32(seqtext_.y) > SUBTITLE_BOX_MIN Then
                        subMin = seqtext_.y
                    End If
                    If U32(subMax) < U32(seqtext_.y) Then
                        subMax = seqtext_.y
                    End If
                End If
            End If

            ' (if (frameTime >= seqtext.endTime && frameTime < seqtext.endTime): its body is commented out)
        End If
    Next

    subMin = subMin - D_H2 'adjust video window here because text is already ofset for big screens
    subMax = subMax - D_H2

    If U32(subMin) < SUBTITLE_BOX_MIN Then
        subMin = SUBTITLE_BOX_MIN
    End If
    If U32(subMax) > SUBTITLE_BOX_MAX Then
        subMax = SUBTITLE_BOX_MAX
    End If

    If U32(subMax) > U32(subMin) Then
        bMoreThanOneSequenceLine = True
    End If

    'call sequence player to download last frame
    stillPlaying = seq_Update()
    'print any text over the video
    frameTime = seq_GetFrameTime()

    For i = 0 To MAX_TEXT_OVERLAYS - 1
        currentText = aSeqList(currentPlaySeq).aText(i)
        If Len(currentText.pText) <> 0 Then
            If ((frameTime >= currentText.startTime) And (frameTime <= currentText.endTime)) Or _
               (aSeqList(currentPlaySeq).bSeqLoop) Then 'if its a looped video always draw the text
                Set t = cachedSeqText(i)
                If bMoreThanOneSequenceLine Then
                    currentText.x = 20 + D_W2
                End If
                t.setText currentText.pText, font_scaled
                t.render currentText.x - 1, currentText.y - 1, psPalette(WZCOL_GREY)
                t.render currentText.x - 1, currentText.y + 1, psPalette(WZCOL_GREY)
                t.render currentText.x + 1, currentText.y - 1, psPalette(WZCOL_GREY)
                t.render currentText.x + 1, currentText.y + 1, psPalette(WZCOL_GREY)
                t.render currentText.x, currentText.y, psPalette(WZCOL_WHITE)
            End If
        End If
    Next
    If Not stillPlaying Or bHoldSeqForAudio Then
        If bAudioPlaying Then
            If aSeqList(currentPlaySeq).bSeqLoop Then
                seq_Shutdown

                If Not seq_Play(aVideoProvider) Then
                    bHoldSeqForAudio = True
                End If
            Else
                bHoldSeqForAudio = True
            End If
            seq_UpdateFullScreenVideo = True 'should hold the video
            Exit Function
        Else
            seq_UpdateFullScreenVideo = False 'should terminate the video
            Exit Function
        End If
    End If

    seq_UpdateFullScreenVideo = True
End Function

Public Sub seqReleaseAll()
    seq_Shutdown
    aVideoProvider = 0
    Set wzCachedSeqText = New Collection
End Sub

Public Function seq_StopFullScreenVideo() As Boolean
    If Not seq_AnySeqLeft() Then
        loop_ClearVideoPlaybackMode
    End If

    OnDemandVideoDownloader_cancelDownloadRequest currVideoName

    seq_Shutdown

    aVideoProvider = 0
    Set wzCachedSeqText = New Collection

    seq_StopFullScreenVideo = True
End Function

' add a string at x,y or add string below last line if x and y are 0
Public Function seq_AddTextForVideo(ByVal pText As String, ByVal xOffset As Long, ByVal yOffset As Long, ByVal startTime As Double, ByVal endTime As Double, ByVal textJustification As SEQ_TEXT_POSITIONING) As Boolean
    Dim sourceLength As Long, currentLength As Long, currentText As String, buffer_width As Long, ct As Long
    Const MIN_JUSTIFICATION As Long = 40
    Dim justification As Long
    ' make sure we take xOffset into account, we don't always start at 0
    buffer_width = pie_GetVideoBufferWidth() - xOffset

    If Not (aSeqList(currentSeq).currentText < MAX_TEXT_OVERLAYS) Then ASSERT_FAILURE "aSeqList[currentSeq].currentText < MAX_TEXT_OVERLAYS", "src/seqdisp.cpp:774", "seq_AddTextForVideo", "too many text lines": seq_AddTextForVideo = False: Exit Function

    sourceLength = Len(pText)
    currentLength = sourceLength
    ct = aSeqList(currentSeq).currentText

    'if the string is bigger than the buffer get the last end of the last fullword in the buffer
    If currentLength >= MAX_STR_LENGTH Then
        currentLength = MAX_STR_LENGTH - 1
        'get end of the last word
        Do While (seqCharAt(pText, currentLength) <> " ") And (currentLength > 0)
            currentLength = currentLength - 1
        Loop
    End If

    currentText = Left$(pText, currentLength)    'terminate the string what ever

    'check the string is shortenough to print
    'if not take a word of the end and try again
    Do While (currentLength > 0) And U32(iV_GetTextWidth(currentText, font_scaled)) > U32(buffer_width)
        currentLength = currentLength - 1
        Do While (seqCharAt(pText, currentLength) <> " ") And (currentLength > 0)
            currentLength = currentLength - 1
        Loop
        currentText = Left$(currentText, currentLength)  'terminate the string what ever
    Loop
    currentText = Left$(currentText, currentLength)    'terminate the string what ever
    aSeqList(currentSeq).aText(ct).pText = currentText

    'check if x and y are 0 and put text on next line
    If ((xOffset = 0) And (yOffset = 0)) And (currentLength > 0) Then
        aSeqList(currentSeq).aText(ct).x = lastX
        aSeqList(currentSeq).aText(ct).y = _
            aSeqList(currentSeq).aText(ct - 1).y + iV_GetTextLineSize(font_scaled)
    Else
        aSeqList(currentSeq).aText(ct).x = xOffset + D_W2
        aSeqList(currentSeq).aText(ct).y = yOffset + D_H2
    End If
    lastX = aSeqList(currentSeq).aText(ct).x

    justification = buffer_width - iV_GetTextWidth(currentText, font_scaled)
    If textJustification = SEQ_TEXT_JUSTIFY And currentLength = sourceLength And justification > MIN_JUSTIFICATION Then
        aSeqList(currentSeq).aText(ct).x = aSeqList(currentSeq).aText(ct).x + (justification \ 2)
    End If

    'set start and finish times for the objects
    aSeqList(currentSeq).aText(ct).startTime = startTime
    aSeqList(currentSeq).aText(ct).endTime = endTime
    aSeqList(currentSeq).aText(ct).bSubtitle = (textJustification <> 0)

    aSeqList(currentSeq).currentText = aSeqList(currentSeq).currentText + 1
    If aSeqList(currentSeq).currentText >= MAX_TEXT_OVERLAYS Then
        aSeqList(currentSeq).currentText = 0
    End If

    'check text is okay on the screen
    If currentLength < sourceLength Then
        'RECURSE x= 0 y = 0 for nextLine
        If textJustification = SEQ_TEXT_JUSTIFY Then
            textJustification = SEQ_TEXT_POSITION
        End If
        seq_AddTextForVideo Mid$(pText, currentLength + 2), 0, 0, startTime, endTime, textJustification
    End If
    seq_AddTextForVideo = True
End Function

' (VB) pText[i] (0-based; the end is '\0', "")
Private Function seqCharAt(ByVal s As String, ByVal i As Long) As String
    seqCharAt = Mid$(s, i + 1, 1)
End Function

' (VB) lineStream >> xOffset >> yOffset >> startTime >> endTime: the four numbers at the start of the line (whitespace
' between them; std::istream's classic locale), False when one cannot be read
Private Function seqReadLineNumbers(ByVal s As String, xOffset As Long, yOffset As Long, startTime As Double, endTime As Double) As Boolean
    Dim p As Long
    p = 1
    If Not seqReadNumber(s, p, False, xOffset, 0#) Then Exit Function
    If Not seqReadNumber(s, p, False, yOffset, 0#) Then Exit Function
    If Not seqReadNumber(s, p, True, 0, startTime) Then Exit Function
    If Not seqReadNumber(s, p, True, 0, endTime) Then Exit Function
    seqReadLineNumbers = True
End Function

' (VB) one operator>> of an int (isDouble False: into i) or a double (into d) from s at p (1-based), p moved past it
Private Function seqReadNumber(ByVal s As String, p As Long, ByVal isDouble As Boolean, i As Long, d As Double) As Boolean
    Dim start As Long, c As String, digits As Long, t As String
    Do While p <= Len(s)                ' skipws
        c = Mid$(s, p, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
        p = p + 1
    Loop
    start = p
    If p <= Len(s) Then
        c = Mid$(s, p, 1)
        If c = "-" Or c = "+" Then p = p + 1
    End If
    Do While p <= Len(s)
        c = Mid$(s, p, 1)
        If c < "0" Or c > "9" Then Exit Do
        p = p + 1
        digits = digits + 1
    Loop
    If isDouble Then
        If p <= Len(s) Then
            If Mid$(s, p, 1) = "." Then
                p = p + 1
                Do While p <= Len(s)
                    c = Mid$(s, p, 1)
                    If c < "0" Or c > "9" Then Exit Do
                    p = p + 1
                    digits = digits + 1
                Loop
            End If
        End If
    End If
    If digits = 0 Then Exit Function
    t = Mid$(s, start, p - start)
    If isDouble Then
        d = Val(t)                      ' (Val: the classic locale's '.')
    Else
        i = CLng(Val(t))
    End If
    seqReadNumber = True
End Function

Private Function seq_AddTextFromFile(ByVal pTextName As String, ByVal textJustification As SEQ_TEXT_POSITIONING) As Boolean
    Dim aTextName As String, fileSize As Long, xOffset As Long, yOffset As Long, startTime As Double, endTime As Double
    Dim content As String, lines() As String, k As Long, pCurrentLine As String, firstQuote As Long, lastQuote As Long, text As String

    ' NOTE: The original game never had a fullscreen mode for FMVs on >640x480 screens.
    ' They would just use double sized videos, and move the text to that area.
    ' We just use the full screen for text right now, instead of using offsets.
    ' However, depending on reaction, we may use the old style again.
    D_H2 = 0                '( pie_GetVideoBufferHeight()- 480)/2;
    D_W2 = 0                '( pie_GetVideoBufferWidth() - 640)/2;
    aTextName = Left$("sequenceaudio/" & pTextName, MAX_STR_LENGTH - 1)

    If loadFileToBufferNoError(aTextName, fileLoadBuffer, FILE_LOAD_BUFFER_SIZE, fileSize) = False Then 'Did I mention this is lame? -Q
        seq_AddTextFromFile = False
        Exit Function
    End If

    ' (pTextBuffer = fileLoadBuffer; std::istringstream stream(pTextBuffer): the text up to its '\0')
    content = utf8_decode(fileLoadBuffer, 0, fileSize)
    k = InStr(1, content, vbNullChar, vbBinaryCompare)
    If k <> 0 Then content = Left$(content, k - 1)
    lines = Split(content, vbLf)        ' std::getline(stream, pCurrentLine, '\n')
    For k = 0 To UBound(lines)
        pCurrentLine = lines(k)
        If k = UBound(lines) And Len(pCurrentLine) = 0 Then Exit For    ' (no line after the last '\n')
        If Len(pCurrentLine) <> 0 Then
            If Left$(pCurrentLine, 1) <> "/" Then
                If seqReadLineNumbers(pCurrentLine, xOffset, yOffset, startTime, endTime) Then
                    ' Since all the positioning was hardcoded to specific values, we now calculate the
                    ' ratio of our screen, compared to what the game expects and multiply that to x, y.
                    ' This makes the text always take up the full screen, instead of original style.
                    xOffset = Fix(CDbl(pie_GetVideoBufferWidth()) / 640# * CDbl(xOffset))
                    yOffset = Fix(CDbl(pie_GetVideoBufferHeight()) / 480# * CDbl(yOffset))
                    'get the text
                    firstQuote = InStr(1, pCurrentLine, """", vbBinaryCompare)
                    lastQuote = InStrRev(pCurrentLine, """", -1, vbBinaryCompare)

                    If Not (firstQuote <> 0 And lastQuote <> 0 And firstQuote < lastQuote) Then ASSERT_FAILURE "firstQuote != std::string::npos && lastQuote != std::string::npos && firstQuote < lastQuote", "src/seqdisp.cpp:895", "seq_AddTextFromFile", "error parsing text file"
                    If firstQuote <> 0 And lastQuote <> 0 And firstQuote < lastQuote Then
                        text = Mid$(pCurrentLine, firstQuote + 1, lastQuote - firstQuote - 1)
                        seq_AddTextForVideo text, xOffset, yOffset, startTime, endTime, textJustification
                    End If
                End If
            End If
        End If
    Next
    seq_AddTextFromFile = True
End Function

'clear the sequence list
Public Sub seq_ClearSeqList()
    Dim i As Long
    currentSeq = -1
    currentPlaySeq = -1
    For i = 0 To MAX_SEQ_LIST - 1
        SEQLIST_reset aSeqList(i)
    Next
    OnDemandVideoDownloader_clear
End Sub

'add a sequence to the list to be played
' (pTextName: vbNullString = nullptr)
Public Sub seq_AddSeqToList(ByVal pSeqName As String, ByVal audioName As String, ByVal pTextName As String, ByVal bLoop As Boolean)
    Dim aSubtitleName As String, extension As Long
    currentSeq = currentSeq + 1

    If Not (currentSeq < MAX_SEQ_LIST) Then ASSERT_FAILURE "currentSeq < MAX_SEQ_LIST", "src/seqdisp.cpp:924", "seq_AddSeqToList", "too many sequences": Exit Sub

    If OnDemandVideoDownloader_hasBaseURLPath() Then
        OnDemandVideoDownloader_requestVideoData pSeqName
    End If

    'OK so add it to the list
    aSeqList(currentSeq).pSeq = pSeqName
    aSeqList(currentSeq).pAudio = audioName
    aSeqList(currentSeq).bSeqLoop = bLoop
    If StrPtr(pTextName) <> 0 Then
        ' Ordinary text shouldn't be justified
        seq_AddTextFromFile pTextName, SEQ_TEXT_POSITION
    End If

    If bSeqSubtitles Then
        aSubtitleName = Left$(pSeqName, MAX_STR_LENGTH - 1)

        ' check for a subtitle file
        extension = InStrRev(aSubtitleName, ".", -1, vbBinaryCompare)
        If extension <> 0 Then
            aSubtitleName = Left$(aSubtitleName, extension - 1)
        End If
        aSubtitleName = Left$(aSubtitleName & ".txt", MAX_STR_LENGTH - 1)

        ' Subtitles should be center justified
        seq_AddTextFromFile aSubtitleName, SEQ_TEXT_JUSTIFY
    End If
End Sub

'checks to see if there are any sequences left in the list to play
Public Function seq_AnySeqLeft() As Boolean
    Dim nextSeq As Long
    nextSeq = currentPlaySeq + 1

    'check haven't reached end
    If nextSeq >= MAX_SEQ_LIST Then
        seq_AnySeqLeft = False
        Exit Function
    End If
    seq_AnySeqLeft = Len(aSeqList(nextSeq).pSeq) <> 0
End Function

Public Function seq_StartNextFullScreenVideo() As Boolean
    Dim bPlayedOK As Boolean

    currentPlaySeq = currentPlaySeq + 1
    If currentPlaySeq >= MAX_SEQ_LIST Then
        bPlayedOK = False
    Else
        bPlayedOK = seq_StartFullScreenVideo(aSeqList(currentPlaySeq).pSeq, aSeqList(currentPlaySeq).pAudio, VIDEO_USER_CHOSEN_RESOLUTION)
    End If

    If bPlayedOK = False Then
        'don't do the callback if we're playing the win/lose video
        If getScriptWinLoseVideo() <> 0 Then
            displayGameOver getScriptWinLoseVideo() = PLAY_WIN, False
        End If
    End If

    seq_StartNextFullScreenVideo = bPlayedOK
End Function


Public Sub seq_SetSubtitles(ByVal bNewState As Boolean)
    bSeqSubtitles = bNewState
End Sub

Public Function seq_GetSubtitles() As Boolean
    seq_GetSubtitles = bSeqSubtitles
End Function
