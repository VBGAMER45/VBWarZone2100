Attribute VB_Name = "sequence_c"
' lib/sequence/sequence.cpp, sequence.h: the video player (Ogg Theora video, Ogg Vorbis audio), whole. The
' calls into its libraries:
' * libogg: ogg_framing_c (src/shared/ogg, compiled in); its structs are ogg_h's UDTs (copies, not pointers);
' * libtheora's legacy API (theora.h): theora_h and its modules (this project's);
' * libvorbis's synthesis API: WzSound's, through SoundApi (over its stb_vorbis port): vorbis_info,
'   vorbis_comment, vorbis_dsp_state and vorbis_block are Long handles, an ogg_packet is passed as its fields;
'   videodata.vi's channels and rate are read once the headers are in (vi_channels, vi_rate);
' * OpenAL, sound_GetUIVolume, sound_GetError, audio_Disabled: SoundApi's (al_win's names; al.h's constants
'   are this module's: an AL_ member of SoundApi's enums would clash with al_win's constants, measured); a
'   pointer to ALuint names is a one-element array; al_win mixes only when polled, so seq_Update polls it
'   (al_vbProcessCurrent: OpenAL Soft's mixer thread);
' * the renderer: GfxApi's. std::unique_ptr<GFX> videoGfx is a GFX handle of GfxApi's GFX_vb* methods (WzGfx's
'   GFX class is private to it), 0 = nullptr; iV_Image is WzGfx's class; the texture's RGBA pixels are made
'   in frameRGBA (a Long per pixel) and copied into VideoFrameBitmap whole;
' * PhysFS: FrameworkApi's.
' A std::shared_ptr<VideoProvider> is a VideoProvider object (Nothing = nullptr): VideoProvider.cls holds the
' derived PhysFSVideoProvider or MemoryBufferVideoProvider. The statics' initialisers (basetime and
' videobuf_granulepos = -1) are set at the first seq_Play (sequence_statics). timer.cpp: timer_c.
'
' This file is derived from the SDL player example as found in the OggTheora
' software codec source code. In particular this is examples/player_example.c
' as found in OggTheora 1.0beta3.
'
' The copyright to this file was originally owned by and licensed as follows.
' Please note, however, that *this* file, i.e. the one you are currently
' reading is not licensed as such anymore.
'
' Copyright (C) 2002-2007 Xiph.org Foundation
' (the BSD-style licence text of sequence.cpp's header)
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' AL/al.h (the values sequence.cpp uses)
Private Const AL_GAIN As Long = &H100A
Private Const AL_SOURCE_STATE As Long = &H1010
Private Const AL_PLAYING As Long = &H1012
Private Const AL_BUFFERS_QUEUED As Long = &H1015
Private Const AL_BUFFERS_PROCESSED As Long = &H1016
Private Const AL_FORMAT_MONO16 As Long = &H1101
Private Const AL_FORMAT_STEREO16 As Long = &H1103
Private Const AL_SIZE As Long = &H2004

' stick this in sequence.h perhaps?
' (VB) struct AudioData: AudioData_ (VB names ignore case: the variable audiodata keeps the name)
Private Type AudioData_
    buffer1 As Long                         ' ALuint: buffer 1
    buffer2 As Long                         ' buffer 2
    source As Long                          ' source
    totbufstarted As Long                   ' number of buffers started
    audiofd_fragsize As Long                ' audio fragment size, used to calculate how big audiobuf is
    audiobuf_fill As Long                   ' how full our audio buffer is
End Type

' (VB) struct VideoData: VideoData_
Private Type VideoData_
    oy As ogg_sync_state                    ' ogg sync state
    og As ogg_page                          ' ogg page
    vo As ogg_stream_state                  ' ogg stream state
    to_ As ogg_stream_state                 ' ogg stream state (to: a VB keyword)
    ti As theora_info                       ' theora info
    tc As theora_comment                    ' theora comment
    td As theora_state                      ' theora state
    vi As Long                              ' vorbis info (WzSound's handle)
    vd As Long                              ' vorbis display state
    vb As Long                              ' vorbis block
    vc As Long                              ' vorbis comment
End Type
' stick that in sequence.h perhaps?

' for our audio structure
Private audiodata As AudioData_
Private sourcestate As Long                 ' ALint = 0: Source state information

' for our video structure
Private videodata As VideoData_
Private vi_channels As Long                 ' (VB) videodata.vi.channels
Private vi_rate As Long                     ' (VB) videodata.vi.rate

'/ are we playing a theora stream?
Private theora_p As Long

'/ are we playing an ogg vorbis stream?
Private vorbis_p As Long

Private stateflag As Boolean
Private videoplaying As Boolean
Private videobuf_ready As Boolean           ' single frame video buffer ready for processing
Private audiobuf_ready As Boolean           ' single 'frame' audio buffer ready for processing

' file handle
Private inVideoProvider As VideoProvider    ' static std::shared_ptr<VideoProvider>

Private VideoFrameBitmap As New iV_Image    ' texture buffer
Private frameRGBA() As Long                 ' (VB) the bitmap's pixels as they are made (frame_width x the bitmap's height)
Private scanMode As SCANLINE_MODE           ' = SCANLINES_OFF
Private audiobuf() As Integer               ' ogg_int16_t *audiobuf = nullptr: audio buffer
Private audiobufBytes() As Byte             ' (VB) audiobuf's bytes for alBufferData
Private audiobufAllocated As Boolean        ' (VB) audiobuf != nullptr

' For timing
Private audioTime As Double

Private videobuf_time As Double
Private sampletimeOffset As Double
Private basetime As Double                  ' = -1 (sequence_statics)
Private last_time As Double
Private timer_expire As Double
Private timer_started As Boolean

Private audiobuf_granulepos As Double       ' ogg_int64_t = 0: time position of last sample
Private videobuf_granulepos As Double       ' ogg_int64_t = -1 (sequence_statics): time position of last video frame

' frame & dropped frame counter
Private frames As Long
Private dropped As Long

' Screen dimensions
Private Const NUM_VERTICES As Long = 4
Private videoGfx As Long                    ' std::unique_ptr<GFX> videoGfx = nullptr: GfxApi's GFX handle
Private vertices(0 To NUM_VERTICES * 2 - 1) As Single   ' gfx_api::gfxFloat vertices[NUM_VERTICES][2]
Private Scrnvidpos(0 To 2) As Single        ' gfx_api::gfxFloat

Private use_scanlines As SCANLINE_MODE
Private scanlinesDisabled As Boolean

Private Const texture_width As Long = 1024  ' size_t
Private Const texture_height As Long = 1024

' (not __BIG_ENDIAN__)
Private Const Rshift As Long = 0
Private Const Gshift As Long = 8
Private Const Bshift As Long = 16
Private Const Ashift As Long = 24
' RGBmask is used only after right-shifting, so ignore the leftmost bit of each byte
Private Const RGBmask As Long = &H7F7F7F
Private Const Amask As Long = &HFF000000

Private staticsReady As Boolean             ' (VB) sequence_statics has run

' (VB) the statics' initialisers that are not 0
Private Sub sequence_statics()
    basetime = -1
    videobuf_granulepos = -1
    staticsReady = True
End Sub

'//// Abstraction for handling loading videos from files, or from memory buffer
' (VideoProvider.cls, PhysFSVideoProvider.cls, MemoryBufferVideoProvider.cls)

' std::shared_ptr<VideoProvider> makeVideoProvider(PHYSFS_file *in, const WzString& filename)
Public Function makeVideoProvider_file(ByVal in_ As Long, ByVal filename As String) As VideoProvider
    Dim derived As PhysFSVideoProvider, v As VideoProvider
    Set derived = New PhysFSVideoProvider   ' std::make_shared<PhysFSVideoProvider>(in, filename)
    derived.construct in_
    Set v = New VideoProvider
    v.construct_physfs derived, filename
    Set makeVideoProvider_file = v
End Function

' std::shared_ptr<VideoProvider> makeVideoProvider(std::shared_ptr<const std::vector<char>> memoryBuffer, const
' WzString& filename): memoryBuffer is data (all of it; an array without elements is an empty vector)
Public Function makeVideoProvider_data(data() As Byte, ByVal filename As String) As VideoProvider
    Dim derived As MemoryBufferVideoProvider, v As VideoProvider, first As Long, n As Long
    first = 0: n = 0
    On Error Resume Next                    ' (an array without elements: UBound fails)
    first = LBound(data)
    n = UBound(data) - first + 1
    On Error GoTo 0
    Set derived = New MemoryBufferVideoProvider     ' std::make_shared<MemoryBufferVideoProvider>(memoryBuffer, filename)
    derived.construct data, first, n
    Set v = New VideoProvider
    v.construct_memory derived, filename
    Set makeVideoProvider_data = v
End Function

'* helper: push a page into the appropriate stream
'    this can be done blindly; a stream won't accept a page
'    that doesn't belong to it
Private Function queue_page(page As ogg_page) As Long
    If theora_p <> 0 Then
        ogg_stream_pagein videodata.to_, page
    End If

    If vorbis_p <> 0 Then
        ogg_stream_pagein videodata.vo, page
    End If

    queue_page = 0
End Function

' sets the frames number we are on
Private Sub seq_SetFrameNumber(ByVal frame As Long)
    frames = frame
End Sub

'/ @TODO FIXME:  This routine can & will fail when sources are used up!
Private Sub open_audio()
    Dim volume As Single, names(0 To 0) As Long
    volume = 1!

    audiodata.audiofd_fragsize = (((vi_channels * 16) \ 8) * vi_rate)
    ReDim audiobuf(0 To audiodata.audiofd_fragsize \ 2 - 1)     ' (ogg_int16_t *)malloc(audiodata.audiofd_fragsize)
    ReDim audiobufBytes(0 To audiodata.audiofd_fragsize - 1)
    audiobufAllocated = True

    ' FIX ME:  This call will fail, since we have, most likely, already
    ' used up all available sources late in the game!
    ' openal
    names(0) = audiodata.source
    alGenSources 1, names                   ' &audiodata.source
    audiodata.source = names(0)
    sound_GetError "lib/sequence/sequence.cpp:324"     ' on error, we *need* to free another source!

    ' Create an OpenAL buffer and fill it with the decoded data
    names(0) = audiodata.buffer1
    alGenBuffers 1, names
    audiodata.buffer1 = names(0)
    sound_GetError "lib/sequence/sequence.cpp:328"
    names(0) = audiodata.buffer2
    alGenBuffers 1, names
    audiodata.buffer2 = names(0)
    sound_GetError "lib/sequence/sequence.cpp:330"     ' Log and clear error codes

    audiodata.totbufstarted = 0

    ' set the volume of the FMV based on the user's preferences
    volume = sound_GetUIVolume()
    alSourcef audiodata.source, AL_GAIN, volume
End Sub

'* Cleans up audio sources & buffers
Private Sub audio_close()
    ' NOTE: sources & buffers deleted in seq_Shutdown()
    audiobuf_granulepos = 0
    '	clear struct
    '	memset(&audiodata,0x0,sizeof(audiodata));
    audiodata.audiobuf_fill = 0
    audiodata.audiofd_fragsize = 0
    audiodata.source = 0
    audiodata.buffer1 = 0: audiodata.buffer2 = 0
    If audiobufAllocated Then
        Erase audiobuf, audiobufBytes       ' free(audiobuf); audiobuf = nullptr;
        audiobufAllocated = False
    End If
End Sub

' Retrieves the current time with millisecond accuracy
Private Function getTimeNow() As Double
    getTimeNow = Timer_getElapsedMilliSecs()
End Function

' get relative time since beginning playback, compensating for A/V drift
Private Function getRelativeTime() As Double
    If basetime = -1 Then                   ' check to see if this is first time run
        basetime = getTimeNow()
        timer_expire = Timer_getElapsedMicroSecs()
        timer_expire = timer_expire + ToInt((videobuf_time - getTimeNow()) * 1000000#)
        timer_started = True
    End If
    getRelativeTime = ((getTimeNow() - basetime) * 0.001)
End Function

'* Allocates memory to hold the decoded video frame
Private Sub Allocate_videoFrame()
    Dim height_factor As Long, expected_size As Long
    If seq_getScanlinesDisabled() Then scanMode = SCANLINES_OFF Else scanMode = seq_getScanlineMode()
    If scanMode <> SCANLINES_OFF Then height_factor = 2 Else height_factor = 1
    expected_size = videodata.ti.frame_width * videodata.ti.frame_height * 4 * height_factor

    VideoFrameBitmap.allocate videodata.ti.frame_width, videodata.ti.frame_height * height_factor, 4, True
    If Not (VideoFrameBitmap.data_size() = expected_size) Then ASSERT_FAILURE "VideoFrameBitmap.data_size() == expected_size", "lib/sequence/sequence.cpp:389", "Allocate_videoFrame", "Allocated size does not match expected size!"
    ReDim frameRGBA(0 To videodata.ti.frame_width * videodata.ti.frame_height * height_factor - 1)
End Sub

Private Sub deallocateVideoFrame()
    VideoFrameBitmap.clear
    Erase frameRGBA
End Sub

' #define Vclip( x )	( (x > 0) ? ((x < 255) ? x : 255) : 0 )
' main routine to display video on screen.
Private Sub video_write(ByVal update As Boolean)
    Static yuv As yuv_buffer                ' (VB) Static: theora_decode_YUVout keeps its planes' arrays
    Dim video_width As Long, video_height As Long, half_width As Long
    Dim ortho As mat4, translation As mat4, pos As Vector3f, modelViewProjectionMatrix As mat4
    Dim cb As constant_buffer_type_SHADER_GFX_TEXT, helper As pipeline_state_helper
    video_width = videodata.ti.frame_width
    video_height = videodata.ti.frame_height

    ' (auto setRGBAFramePixel = [pRGBABitmapData](int pixelOffset, uint32_t rgbaValue): frameRGBA(pixelOffset))

    If update Then
        half_width = video_width \ 2

        theora_decode_YUVout videodata.td, yuv

        ' fill the RGBA buffer
        video_write_rgba yuv.y, yuv.u, yuv.v, yuv.y_stride, yuv.uv_stride, video_width, video_height, half_width, frameRGBA

        ' (VB) the pixels into the bitmap (pRGBABitmapData = VideoFrameBitmap.bmp_w())
        CopyMemory ByVal VideoFrameBitmap.bmp_w(), frameRGBA(0), VideoFrameBitmap.data_size()
        GFX_updateTexture videoGfx, VideoFrameBitmap
    End If

    ' const auto& modelViewProjectionMatrix = glm::ortho(0.f, static_cast<float>(pie_GetVideoBufferWidth()), static_cast<float>(pie_GetVideoBufferHeight()), 0.f) *
    '     glm::translate(glm::vec3(Scrnvidpos[0], Scrnvidpos[1], Scrnvidpos[2]));
    ortho = glm_ortho(0!, CSng(pie_GetVideoBufferWidth()), CSng(pie_GetVideoBufferHeight()), 0!)
    pos = glm_vec3(Scrnvidpos(0), Scrnvidpos(1), Scrnvidpos(2))
    translation = glm_translate(pos)
    modelViewProjectionMatrix = glm_mat4_mul(ortho, translation)

    Set helper = gfx_api_pso_get(VideoPSO)
    helper.bind
    cb.transform_matrix = modelViewProjectionMatrix     ' { modelViewProjectionMatrix, glm::vec2(0), glm::vec2(0), glm::vec4(1), 0 }
    cb.offset = glm_vec2(0!, 0!)
    cb.size = glm_vec2(0!, 0!)
    cb.color = glm_vec4(1!, 1!, 1!, 1!)
    cb.texture = 0
    helper.bind_constants VarPtr(cb), LenB(cb)
    GFX_draw videoGfx, VideoPSO, VarPtr(modelViewProjectionMatrix)     ' videoGfx->draw<gfx_api::VideoPSO>(...)
End Sub

' (VB) video_write's loops over the planes: yp, up, vp are yuv.y, yuv.u, yuv.v; the pixels into px
' (setRGBAFramePixel(offset, rgba) is px(offset) = rgba)
Private Sub video_write_rgba(yp() As Byte, up() As Byte, vp() As Byte, ByVal y_stride As Long, ByVal uv_stride As Long, _
        ByVal video_width As Long, ByVal video_height As Long, ByVal half_width As Long, px() As Long)
    Dim x As Long, y As Long, rgb_offset As Long, y_offset As Long, uv_offset As Long
    Dim Y_ As Long, U As Long, V As Long, A As Long, C As Long, R As Long, G As Long, B As Long, rgba As Long
    rgb_offset = 0
    y_offset = 0
    uv_offset = 0

    For y = 0 To video_height - 1
        y_offset = y * y_stride
        uv_offset = (y \ 2) * uv_stride     ' (y >> 1)

        For x = 0 To half_width - 1
            Y_ = CLng(yp(y_offset)) - 16      ' int Y = yuv.y[y_offset++] - 16;
            y_offset = y_offset + 1
            U = CLng(up(uv_offset)) - 128
            V = CLng(vp(uv_offset)) - 128
            uv_offset = uv_offset + 1

            A = 298 * Y_
            C = 409 * V

            ' R = Vclip((A + C + 128) >> 8) (a negative x is 0 either way: \ for >>)
            R = (A + C + 128) \ 256
            If R > 0 Then
                If R >= 255 Then R = 255
            Else
                R = 0
            End If
            G = (A - 100 * U - SAR(C, 1) + 128) \ 256
            If G > 0 Then
                If G >= 255 Then G = 255
            Else
                G = 0
            End If
            B = (A + 516 * U + 128) \ 256
            If B > 0 Then
                If B >= 255 Then B = 255
            Else
                B = 0
            End If

            ' (R << Rshift) | (G << Gshift) | (B << Bshift) | (0xFF << Ashift)
            rgba = R Or (G * &H100&) Or (B * &H10000) Or Amask

            px(rgb_offset) = rgba
            If scanMode = SCANLINES_50 Then
                ' halve the rgb values for a dimmed scanline ((rgba >> 1 & RGBmask) | Amask)
                px(rgb_offset + video_width) = (R \ 2) Or ((G \ 2) * &H100&) Or ((B \ 2) * &H10000) Or Amask
            ElseIf scanMode = SCANLINES_BLACK Then
                px(rgb_offset + video_width) = Amask
            End If
            rgb_offset = rgb_offset + 1

            ' second pixel, U and V (and thus C) are the same as before.
            Y_ = CLng(yp(y_offset)) - 16
            y_offset = y_offset + 1
            A = 298 * Y_

            R = (A + C + 128) \ 256
            If R > 0 Then
                If R >= 255 Then R = 255
            Else
                R = 0
            End If
            G = (A - 100 * U - SAR(C, 1) + 128) \ 256
            If G > 0 Then
                If G >= 255 Then G = 255
            Else
                G = 0
            End If
            B = (A + 516 * U + 128) \ 256
            If B > 0 Then
                If B >= 255 Then B = 255
            Else
                B = 0
            End If

            rgba = R Or (G * &H100&) Or (B * &H10000) Or Amask
            px(rgb_offset) = rgba
            If scanMode = SCANLINES_50 Then
                ' halve the rgb values for a dimmed scanline
                px(rgb_offset + video_width) = (R \ 2) Or ((G \ 2) * &H100&) Or ((B \ 2) * &H10000) Or Amask
            ElseIf scanMode = SCANLINES_BLACK Then
                px(rgb_offset + video_width) = Amask
            End If
            rgb_offset = rgb_offset + 1
        Next
        If scanMode <> SCANLINES_OFF Then
            rgb_offset = rgb_offset + video_width
        End If
    Next
End Sub

' FIXME: perhaps we should use wz's routine for audio?
' loads up the audio buffers, and calculates audio sync time.
Private Sub audio_write()
    Dim processed As Long, queued As Long, oldbuffer As Long, buffer_size As Long, current_sample As Double
    Dim names(0 To 0) As Long, fmt As Long
    processed = 0
    queued = 0

    alGetSourcei audiodata.source, AL_BUFFERS_PROCESSED, processed
    alGetSourcei audiodata.source, AL_BUFFERS_QUEUED, queued
    If (audiodata.totbufstarted < 2 Or processed <> 0) And audiodata.audiobuf_fill > 0 Then
        ' we have audiobuf_fill bytes of data
        oldbuffer = 0

        If audiodata.totbufstarted = 0 Then
            oldbuffer = audiodata.buffer1
        ElseIf audiodata.totbufstarted = 1 Then
            oldbuffer = audiodata.buffer2
        Else
            buffer_size = 0
            current_sample = 0

            names(0) = oldbuffer
            alSourceUnqueueBuffers audiodata.source, 1, names   ' &oldbuffer
            oldbuffer = names(0)
            alGetBufferi oldbuffer, AL_SIZE, buffer_size
            ' audio time sync
            audioTime = audioTime + CDbl(buffer_size) / (vi_rate * vi_channels)
            debug_ 538, LOG_VIDEO, "audio_write", "Audio sync"
            current_sample = audiobuf_granulepos - audiodata.audiobuf_fill \ 2 \ vi_channels
            sampletimeOffset = sampletimeOffset - (getTimeNow() - Fix(1000# * current_sample / vi_rate))  ' (int64 division)
        End If

        If vi_channels = 1 Then fmt = AL_FORMAT_MONO16 Else fmt = AL_FORMAT_STEREO16
        CopyMemory audiobufBytes(0), audiobuf(0), audiodata.audiobuf_fill     ' (VB) audiobuf as bytes
        alBufferData oldbuffer, fmt, audiobufBytes, audiodata.audiobuf_fill, vi_rate

        names(0) = oldbuffer
        alSourceQueueBuffers audiodata.source, 1, names         ' &oldbuffer
        audiodata.totbufstarted = audiodata.totbufstarted + 1
        If audiodata.totbufstarted > 2 Then
            audiodata.totbufstarted = 2
        End If

        If sourcestate <> AL_PLAYING Then
            debug_ 555, LOG_VIDEO, "audio_write", "starting source\n"
            alSourcePlay audiodata.source
        End If

        audiobuf_ready = False
        audiodata.audiobuf_fill = 0
    End If
End Sub

Private Sub seq_InitOgg()
    debug_ 566, LOG_VIDEO, "seq_InitOgg", "seq_InitOgg"

    stateflag = False
    theora_p = 0
    vorbis_p = 0

    videoplaying = False
    seq_setScanlinesDisabled False

    ' single frame video buffering
    videobuf_ready = False
    videobuf_granulepos = -1
    videobuf_time = 0
    frames = 0
    dropped = 0

    ' single audio fragment audio buffering
    audiodata.audiobuf_fill = 0
    audiobuf_ready = False

    audiobuf_granulepos = 0                 ' time position of last sample
    audioTime = 0

    sampletimeOffset = 0

    ' start up Ogg stream synchronization layer
    ogg_sync_init videodata.oy

    ' init supporting Vorbis structures needed in header parsing
    vorbis_info_init videodata.vi
    vorbis_comment_init videodata.vc

    ' init supporting Theora structuretotbufstarteds needed in header parsing
    theora_comment_init videodata.tc
    theora_info_init videodata.ti
    Timer_Init
    Timer_start
End Sub

' void update_buffers()
Private Sub update_buffers()
    Dim height_factor As Long, vtwidth As Single, vtheight As Single
    Dim texcoords(0 To NUM_VERTICES * 2 - 1) As Single
    If videoGfx = 0 Then
        Exit Sub
    End If

    ' when using scanlines we need to double the height
    If Not seq_getScanlinesDisabled() And seq_getScanlineMode() <> SCANLINES_OFF Then height_factor = 2 Else height_factor = 1
    vtwidth = CSng(U32(videodata.ti.frame_width)) / CSng(texture_width)
    vtheight = CSng(U32(videodata.ti.frame_height)) * height_factor / CSng(texture_height)
    ' gfx_api::gfxFloat texcoords[NUM_VERTICES * 2] = {0.0f, 0.0f, vtwidth, 0.0f, 0.0f, vtheight, vtwidth, vtheight};
    texcoords(0) = 0!: texcoords(1) = 0!: texcoords(2) = vtwidth: texcoords(3) = 0!
    texcoords(4) = 0!: texcoords(5) = vtheight: texcoords(6) = vtwidth: texcoords(7) = vtheight
    GFX_buffers videoGfx, NUM_VERTICES, VarPtr(vertices(0)), VarPtr(texcoords(0))
End Sub

Public Function seq_Play(ByVal video As VideoProvider) As Boolean
    Dim pp_level_max As Long, pp_level As Long, op As ogg_packet, ret As Long, test As ogg_stream_state
    Dim blankStream As ogg_stream_state, blackFrame As iV_Image, videoName As String, isCodec As Boolean
    pp_level_max = 0
    pp_level = 0
    If Not staticsReady Then sequence_statics

    If Not video Is Nothing Then videoName = video.filename() Else videoName = ""
    debug_ 626, LOG_VIDEO, "seq_Play", "starting playback of: %s", videoName

    If videoplaying Then
        debug_ 630, LOG_VIDEO, "seq_Play", "previous movie is not yet finished"
        seq_Shutdown
    End If

    seq_InitOgg

    If video Is Nothing Then
        seq_Play = False
        Exit Function
    End If

    video.seek_begin

    theora_p = 0
    vorbis_p = 0

    ' Ogg file open; parse the headers
    ' Only interested in Vorbis/Theora streams
    Do While Not stateflag
        ret = video.buffer_data(videodata.oy)

        If ret = 0 Then
            Exit Do
        End If

        Do While ogg_sync_pageout(videodata.oy, videodata.og) > 0
            test = blankStream              ' ogg_stream_state test;

            ' is this a mandated initial header? If not, stop parsing
            If ogg_page_bos(videodata.og) = 0 Then
                ' don't leak the page; get it into the appropriate stream
                queue_page videodata.og
                stateflag = True
                Exit Do
            End If

            ogg_stream_init test, ogg_page_serialno(videodata.og)
            ogg_stream_pagein test, videodata.og
            ogg_stream_packetout test, op

            ' identify the codec: try theora
            isCodec = False
            If theora_p = 0 Then
                If theora_decode_header(videodata.ti, videodata.tc, op) >= 0 Then isCodec = True
            End If
            If isCodec Then
                ' it is theora
                videodata.to_ = test        ' memcpy(&videodata.to, &test, sizeof(test));
                theora_p = 1
            Else
                If vorbis_p = 0 Then
                    If vorbis_synthesis_headerin(videodata.vi, videodata.vc, op.packet, op.bytes, op.b_o_s, op.e_o_s, op.granulepos, op.packetno) >= 0 Then isCodec = True
                End If
                If isCodec Then
                    ' it is vorbis
                    videodata.vo = test     ' memcpy(&videodata.vo, &test, sizeof(test));
                    vorbis_p = 1
                Else
                    ' whatever it is, we don't care about it
                    ogg_stream_clear test
                End If
            End If
        Loop
        ' fall through to non-bos page parsing
    Loop

    ' we're expecting more header packets.
    Do While (theora_p <> 0 And theora_p < 3) Or (vorbis_p <> 0 And vorbis_p < 3)
        ' look for further theora headers
        Do While theora_p <> 0 And theora_p < 3
            ret = ogg_stream_packetout(videodata.to_, op)
            If ret = 0 Then Exit Do
            If ret < 0 Then
                debug_ 706, LOG_ERROR, "seq_Play", "Error parsing Theora stream headers; corrupt stream?\n"
                seq_Play = False
                Exit Function
            End If

            If theora_decode_header(videodata.ti, videodata.tc, op) <> 0 Then
                debug_ 712, LOG_ERROR, "seq_Play", "Error parsing Theora stream headers; corrupt stream?\n"
                seq_Play = False
                Exit Function
            End If

            theora_p = theora_p + 1
        Loop

        ' look for more vorbis header packets
        Do While vorbis_p <> 0 And vorbis_p < 3
            ret = ogg_stream_packetout(videodata.vo, op)
            If ret = 0 Then Exit Do
            If ret < 0 Then
                debug_ 724, LOG_ERROR, "seq_Play", "Error parsing Vorbis stream headers; corrupt stream?\n"
                seq_Play = False
                Exit Function
            End If

            If vorbis_synthesis_headerin(videodata.vi, videodata.vc, op.packet, op.bytes, op.b_o_s, op.e_o_s, op.granulepos, op.packetno) <> 0 Then
                debug_ 730, LOG_ERROR, "seq_Play", "Error parsing Vorbis stream headers; corrupt stream?\n"
                seq_Play = False
                Exit Function
            End If

            vorbis_p = vorbis_p + 1
        Loop

        ' The header pages/packets will arrive before anything else we
        '        care about, or the stream is not obeying spec
        If ogg_sync_pageout(videodata.oy, videodata.og) > 0 Then
            queue_page videodata.og         ' demux into the appropriate stream
        Else
            ret = video.buffer_data(videodata.oy)  ' someone needs more data

            If ret = 0 Then
                debug_ 749, LOG_ERROR, "seq_Play", "End of file while searching for codec headers.\n"
                seq_Play = False
                Exit Function
            End If
        End If
    Loop

    ' and now we have it all.  initialize decoders
    If theora_p <> 0 Then
        theora_decode_init videodata.td, videodata.ti
        debug_ 759, LOG_VIDEO, "seq_Play", "Ogg logical stream %x is Theora %dx%d %.02f fps video", _
              videodata.to_.serialno, videodata.ti.width, videodata.ti.height, _
              U32(videodata.ti.fps_numerator) / U32(videodata.ti.fps_denominator)
        If videodata.ti.width <> videodata.ti.frame_width Or videodata.ti.height <> videodata.ti.frame_height Then
            debug_ 764, LOG_VIDEO, "seq_Play", "  Frame content is %dx%d with offset (%d,%d)", videodata.ti.frame_width, _
                  videodata.ti.frame_height, videodata.ti.offset_x, videodata.ti.offset_y
        End If

        ' hmm
        theora_control videodata.td, TH_DECCTL_GET_PPLEVEL_MAX, pp_level_max, 4     ' sizeof(pp_level_max)
        pp_level = pp_level_max
        theora_control videodata.td, TH_DECCTL_SET_PPLEVEL, pp_level, 4             ' sizeof(pp_level)
    Else
        ' tear down the partial theora setup
        theora_info_clear videodata.ti
        theora_comment_clear videodata.tc
    End If

    If vorbis_p <> 0 Then
        vorbis_synthesis_init videodata.vd, videodata.vi
        vorbis_block_init videodata.vd, videodata.vb
        vi_channels = vorbis_info_channels(videodata.vi)
        vi_rate = vorbis_info_rate(videodata.vi)
        debug_ 784, LOG_VIDEO, "seq_Play", "Ogg logical stream %x is Vorbis %d channel %d Hz audio", _
              videodata.vo.serialno, vi_channels, vi_rate
    Else
        ' tear down the partial vorbis setup
        vorbis_info_clear videodata.vi
        vorbis_comment_clear videodata.vc
    End If

    ' open audio
    If vorbis_p <> 0 Then
        open_audio
    End If

    ' open video
    If videoGfx <> 0 Then GFX_vbDelete videoGfx     ' (the unique_ptr's assignment deletes the old one)
    videoGfx = GFX_vbNew(GFX_TEXTURE, 2)            ' std::make_unique<GFX>(GFX_TEXTURE, 2)
    If theora_p <> 0 Then
        If U32(videodata.ti.frame_width) > texture_width Or U32(videodata.ti.frame_height) > texture_height Then
            debug_ 806, LOG_ERROR, "seq_Play", "Video size too large, must be below %zu x %zu!", _
                  texture_width, texture_height
            GFX_vbDelete videoGfx           ' videoGfx = nullptr;
            videoGfx = 0
            seq_Play = False
            Exit Function
        End If
        If videodata.ti.pixelformat <> OC_PF_420 Then
            debug_ 813, LOG_ERROR, "seq_Play", "Video not in YUV420 format!"
            GFX_vbDelete videoGfx           ' videoGfx = nullptr;
            videoGfx = 0
            seq_Play = False
            Exit Function
        End If
        Set blackFrame = New iV_Image
        blackFrame.allocate texture_width, texture_height, 4, True

        ' disable scanlines temporarily if the video is too large for the texture or shown too small
        If U32(videodata.ti.frame_height) * 2 > texture_height Or vertices(3 * 2 + 1) < U32(videodata.ti.frame_height) * 2 Then
            seq_setScanlinesDisabled True
        End If

        Allocate_videoFrame
        GFX_makeCompatibleTexture videoGfx, blackFrame, "mem::blackframe"
        GFX_updateTexture videoGfx, blackFrame
        blackFrame.clear

        update_buffers
    End If

    ' on to the main decode loop.  We assume in this example that audio
    '    and video start roughly together, and don't begin playback until
    '    we have a start frame for both.  This is not necessarily a valid
    '    assumption in Ogg A/V streams! It will always be true of the
    '    example_encoder (and most streams) though.
    sampletimeOffset = getTimeNow()
    videoplaying = True
    Set inVideoProvider = video             ' only set this here (after videoplaying is set), so that it's always freed in seq_Shutdown()
    seq_Play = True
End Function

Public Function seq_Playing() As Boolean
    seq_Playing = videoplaying
End Function

'*
' * Display the next frame and play the sound.
' * \return false if the end of the video is reached.
Public Function seq_Update() As Boolean
    Dim op As ogg_packet, ret As Long, i As Long, j As Long, pcm() As Single, count As Long, maxsamples As Long
    Dim val_ As Long, now_time As Double, delay As Double, atFrameTime As Boolean

    ' we want a video and audio frame ready to go at all times.  If
    '       we have to buffer incoming, buffer the compressed data (ie, let
    '       ogg do the buffering)
    If Not videoplaying Then
        debug_ 867, LOG_VIDEO, "seq_Update", "no movie playing"
        seq_Update = False
        Exit Function
    End If

    Do While vorbis_p <> 0 And Not audiobuf_ready
        ' if there's pending, decoded audio, grab it
        ret = vorbis_synthesis_pcmout(videodata.vd, pcm)
        If ret > 0 Then
            ' we now have float pcm data in pcm
            ' going to convert that to int pcm in audiobuf
            count = audiodata.audiobuf_fill \ 2
            maxsamples = (audiodata.audiofd_fragsize - audiodata.audiobuf_fill) \ 2 \ vi_channels

            i = 0
            Do While i < ret And i < maxsamples
                For j = 0 To vi_channels - 1
                    ' static_cast<int>(nearbyint(pcm[j][i] * 32767.f)) (the product in the x87's precision)
                    val_ = CLng(pcm(j, i) * 32767!)

                    If val_ > 32767 Then
                        val_ = 32767
                    ElseIf val_ < -32768 Then
                        val_ = -32768
                    End If
                    audiobuf(count) = val_
                    count = count + 1
                Next
                i = i + 1
            Loop

            vorbis_synthesis_read videodata.vd, i
            audiodata.audiobuf_fill = audiodata.audiobuf_fill + i * vi_channels * 2

            If audiodata.audiobuf_fill = audiodata.audiofd_fragsize Then
                audiobuf_ready = True
            End If

            If vorbis_dsp_granulepos(videodata.vd) >= 0 Then
                audiobuf_granulepos = vorbis_dsp_granulepos(videodata.vd) - ret + i
            Else
                audiobuf_granulepos = audiobuf_granulepos + i
            End If
        Else
            ' no pending audio; is there a pending packet to decode?
            If ogg_stream_packetout(videodata.vo, op) > 0 Then
                If vorbis_synthesis(videodata.vb, op.packet, op.bytes, op.b_o_s, op.e_o_s, op.granulepos, op.packetno) = 0 Then
                    ' test for success!
                    vorbis_synthesis_blockin videodata.vd, videodata.vb
                End If
            Else
                ' we need more data; break out to suck in another page
                Exit Do
            End If
        End If
    Loop

    Do While theora_p <> 0 And Not videobuf_ready
        ' theora is one in, one out...
        If ogg_stream_packetout(videodata.to_, op) > 0 Then
            now_time = 0
            delay = 0

            theora_decode_packetin videodata.td, op
            videobuf_granulepos = videodata.td.granulepos
            videobuf_time = theora_granule_time(videodata.td, videobuf_granulepos)

            now_time = getRelativeTime()
            delay = videobuf_time - getRelativeTime()

            If (delay >= 0!) Or (now_time - last_time >= 1!) Then
                videobuf_ready = True
                seq_SetFrameNumber seq_GetFrameNumber() + 1
            Else
                ' running slow, so we skip this frame
                dropped = dropped + 1
            End If
        Else
            Exit Do
        End If
    Loop

    ' (VB) OpenAL Soft's mixer runs in a thread of its own; WzSound's al_win mixes when polled, and the full
    ' screen videos run outside the loops that poll it (main.cpp's videoLoop): the source plays on here
    If vorbis_p <> 0 Then al_vbProcessCurrent
    alGetSourcei audiodata.source, AL_SOURCE_STATE, sourcestate

    If inVideoProvider.end_of_data() _
        And Not videobuf_ready _
        And ((Not audiobuf_ready And (audiodata.audiobuf_fill = 0)) Or audio_Disabled()) _
        And sourcestate <> AL_PLAYING Then
        video_write False
        seq_Shutdown
        debug_ 976, LOG_VIDEO, "seq_Update", "video finished"
        seq_Update = False
        Exit Function
    End If

    If Not videobuf_ready Or Not audiobuf_ready Then
        ' no data yet for somebody.  Grab another page
        ret = inVideoProvider.buffer_data(videodata.oy)
        Do While ogg_sync_pageout(videodata.oy, videodata.og) > 0
            queue_page videodata.og
        Loop
    End If

    ' If playback has begun, top audio buffer off immediately.
    If vorbis_p <> 0 _
        And stateflag _
        And Not audio_Disabled() Then
        ' FIXME : it is possible to crash if people are playing with no sound.
        ' play the data in pcm
        audio_write
    End If

    ' are we at or past time for this video frame?
    ' (VB) getRelativeTime() only when the C's && reaches it: its first call starts the clock
    atFrameTime = False
    If stateflag And videobuf_ready Then
        If videobuf_time <= getRelativeTime() Then atFrameTime = True
    End If
    If atFrameTime Then
        video_write True
        last_time = getRelativeTime()
        videobuf_ready = False
    ElseIf stateflag Then
        video_write False
    End If

    ' if our buffers either don't exist or are ready to go,
    '       we can begin playback
    If (theora_p = 0 Or videobuf_ready) And (vorbis_p = 0 Or audiobuf_ready) And Not stateflag Then
        debug_ 1016, LOG_VIDEO, "seq_Update", "all buffers ready"
        stateflag = True
    End If

    ' same if we've run out of input
    If inVideoProvider.end_of_data() Then
        stateflag = True
    End If

    seq_Update = True
End Function

Public Sub seq_Shutdown()
    Dim names(0 To 0) As Long
    ' tear it all down
    debug_ 1032, LOG_VIDEO, "seq_Shutdown", "seq_Shutdown"

    If Not videoplaying Then
        debug_ 1036, LOG_VIDEO, "seq_Shutdown", "movie is not playing"
        Exit Sub
    End If
    If videoGfx <> 0 Then GFX_vbDelete videoGfx
    videoGfx = 0                            ' videoGfx = nullptr;

    If vorbis_p <> 0 Then
        ogg_stream_clear videodata.vo
        vorbis_block_clear videodata.vb
        vorbis_dsp_clear videodata.vd
        vorbis_comment_clear videodata.vc
        vorbis_info_clear videodata.vi

        names(0) = audiodata.source
        alDeleteSources 1, names
        names(0) = audiodata.buffer1
        alDeleteBuffers 1, names
        names(0) = audiodata.buffer2
        alDeleteBuffers 1, names

        audio_close
    End If

    If theora_p <> 0 Then
        ogg_stream_clear videodata.to_
        theora_clear videodata.td
        theora_comment_clear videodata.tc
        theora_info_clear videodata.ti
        deallocateVideoFrame
    End If

    ogg_sync_clear videodata.oy

    Set inVideoProvider = Nothing           ' inVideoProvider.reset();

    videoplaying = False
    seq_setScanlinesDisabled False
    Timer_stop

    audioTime = 0
    ' sampletimeOffset = last_time = timer_expire = timer_started = 0;
    timer_started = False
    timer_expire = 0
    last_time = 0
    sampletimeOffset = 0
    basetime = -1

    debug_ 1077, LOG_VIDEO, "seq_Shutdown", " **** frames = %d dropped = %d ****", frames, dropped
End Sub

Public Function seq_GetFrameNumber() As Long
    seq_GetFrameNumber = frames
End Function

Public Function seq_GetFrameTime() As Double
    seq_GetFrameTime = videobuf_time
End Function

' this controls the size of the video to display on screen
Public Sub seq_SetDisplaySize(ByVal sizeX As Long, ByVal sizeY As Long, ByVal posX As Long, ByVal posY As Long)
    ' (vertices[i][k] is vertices(i * 2 + k))
    vertices(0 * 2 + 0) = 0!
    vertices(0 * 2 + 1) = 0!
    vertices(1 * 2 + 0) = sizeX
    vertices(1 * 2 + 1) = 0!
    vertices(2 * 2 + 0) = 0!
    vertices(2 * 2 + 1) = sizeY
    vertices(3 * 2 + 0) = sizeX
    vertices(3 * 2 + 1) = sizeY

    Scrnvidpos(0) = posX
    Scrnvidpos(1) = posY
    Scrnvidpos(2) = 0!

    update_buffers
End Sub

Public Sub seq_setScanlinesDisabled(ByVal flag As Boolean)
    scanlinesDisabled = flag
End Sub

Public Function seq_getScanlinesDisabled() As Boolean
    seq_getScanlinesDisabled = scanlinesDisabled
End Function

Public Sub seq_setScanlineMode(ByVal mode As SCANLINE_MODE)
    use_scanlines = mode
End Sub

Public Function seq_getScanlineMode() As SCANLINE_MODE
    seq_getScanlineMode = use_scanlines
End Function
