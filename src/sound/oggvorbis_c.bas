Attribute VB_Name = "oggvorbis_c"
' lib/sound/oggvorbis.cpp: WZVorbisDecoder::fromFilename (a VB class has no static function: the class is
' WZVorbisDecoder.cls). The C++ reads the file through libvorbisfile's callbacks (wz_ogg_read, wz_ogg_seek,
' wz_ogg_tell on the PHYSFS_file); the port decodes with stb_vorbis (stb_vorbis_c) from memory, so the
' file is read whole and closed at once, and the decoder keeps its bytes. libvorbisfile's errors become
' stb_vorbis's (the same messages for the nearest ones); ov_seekable (a file in memory always is) and
' ov_bitrate (stb_vorbis has none) are not asked.
Option Explicit

Private Const OV_EINVAL As Long = -131     ' vorbis/codec.h

Public Function WZVorbisDecoder_fromFilename(ByVal fileName As String) As WZVorbisDecoder
    Dim fileHandle As Long, length As Double, data() As Byte, got As Double, f As Long, error_ As Long
    Dim info As stb_vorbis_info, t As Double, decoder As WZVorbisDecoder

    fileHandle = PHYSFS_openRead(fileName)
    debug_ 142, LOG_WZ, "fromFilename", "Reading...[directory: %s] %s", PHYSFS_getRealDir(fileName) & "", fileName
    If fileHandle = 0 Then
        debug_ 145, LOG_ERROR, "fromFilename", "sound_LoadTrackFromFile: PHYSFS_openRead(""%s"") failed with error: %s" & vbLf, fileName, WZ_PHYSFS_getLastError()
        Set WZVorbisDecoder_fromFilename = Nothing
        Exit Function
    End If
    ' (VB) ov_open_callbacks(fileHandle, ovf, nullptr, 0, wz_oggVorbis_callbacks): the file in memory,
    ' stb_vorbis_open_memory
    length = PHYSFS_fileLength(fileHandle)
    If length > 0# Then ReDim data(0 To CLng(length) - 1) Else ReDim data(0 To 0)
    got = 0#
    If length > 0# Then got = WZ_PHYSFS_readBytes(fileHandle, data, 0, CLng(length))
    PHYSFS_close fileHandle
    If got <> length Or length <= 0# Then
        error_ = -1                         ' OV_EREAD
        f = 0
    Else
        f = stb_vorbis_open_memory(data, CLng(length), error_)
    End If
    If f = 0 Then
        Select Case error_
        Case -1
            debug_ 156, LOG_ERROR, "fromFilename", "A read from media returned an error."
        Case VORBIS_missing_capture_pattern, VORBIS_invalid_first_page, VORBIS_bad_packet_type, VORBIS_ogg_skeleton_not_supported
            debug_ 159, LOG_ERROR, "fromFilename", "Bitstream does not contain any Vorbis data."
        Case VORBIS_invalid_stream_structure_version
            debug_ 162, LOG_ERROR, "fromFilename", "Vorbis version mismatch."
        Case VORBIS_invalid_setup, VORBIS_invalid_stream, VORBIS_unexpected_eof, VORBIS_feature_not_supported, VORBIS_too_many_channels, _
             VORBIS_continued_packet_flag_invalid, VORBIS_incorrect_stream_serial_number
            debug_ 165, LOG_ERROR, "fromFilename", "Invalid Vorbis bitstream header."
        Case VORBIS_outofmem, VORBIS_invalid_api_mixing
            debug_ 168, LOG_ERROR, "fromFilename", "Internal logic fault; indicates a bug or heap/stack corruption."
        Case Else
        End Select
        ' (PHYSFS_close(fileHandle): closed above)
        Set WZVorbisDecoder_fromFilename = Nothing
        Exit Function
    End If
    If Not (f <> 0) Then ASSERT_FAILURE "ovf", "lib/sound/oggvorbis.cpp:176", "fromFilename", "doesn't make sense"
    info = stb_vorbis_get_info(f)
    If info.channels = 0 Then
        debug_ 180, LOG_ERROR, "fromFilename", "ov_info failed on %s", fileName
        stb_vorbis_close f
        Set WZVorbisDecoder_fromFilename = Nothing
        Exit Function
    End If

    ' (ov_seekable: a file in memory is)
    ' static_cast<int64_t>(ov_time_total(ovf.get(), -1)): the whole seconds
    t = Fix(CDbl(stb_vorbis_stream_length_in_seconds(f)))
    If t = OV_EINVAL Then
        debug_ 197, LOG_ERROR, "fromFilename", "%s: the argument was invalid.The requested bitstream did not exist or the bitstream is nonseekable.", fileName
        stb_vorbis_close f
        Set WZVorbisDecoder_fromFilename = Nothing
        Exit Function
    End If
    ' (ov_bitrate: stb_vorbis has none)
    Set decoder = New WZVorbisDecoder
    decoder.construct t, f, info.channels, info.sample_rate, stb_vorbis_stream_length_in_samples(f), data
    ' (VB: New raises an error where new returns nullptr: no "Failed to allocate" here)

    Set WZVorbisDecoder_fromFilename = decoder
End Function
