VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZVorbisDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sound/oggvorbis.cpp, oggvorbis.h: class WZVorbisDecoder (codecs.h's WZDecoder, the interface
' WZDecoder.cls), over stb_vorbis (stb_vorbis_c) in place of libvorbisfile.
' WZVorbisDecoder::fromFilename is oggvorbis_c's WZVorbisDecoder_fromFilename. decode's
' optional<size_t>: decode returns false for nullopt, the bytes in written.
Option Explicit

Implements WZDecoder

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Private Const OV_EINVAL As Long = -131     ' vorbis/codec.h

Private m_total_time As Double              ' int64_t = 0 (whole seconds)
Private m_bufferSize As Double              ' int64_t = 0
Private m_channels As Long                  ' m_info->channels
Private m_rate As Long                      ' m_info->rate
Private m_totalSamples As Long              ' ov_pcm_total(m_ovfile, -1)
Private m_file As Long                      ' (VB) the stb_vorbis handle (C++: m_file, the PHYSFS_file, and m_ovfile)
Private m_data() As Byte                    ' (VB) the file, which stb_vorbis decodes from
Private pcm() As Integer                    ' (VB) the samples stb_vorbis decodes into

' WZVorbisDecoder(int64_t totalTime, PHYSFS_file *f, vorbis_info* info, std::unique_ptr<OggVorbis_File> ovf)
Friend Sub construct(ByVal totalTime As Double, ByVal f As Long, ByVal channels As Long, ByVal rate As Long, _
                     ByVal totalSamples As Long, data() As Byte)
    m_total_time = totalTime
    m_file = f
    m_channels = channels
    m_rate = rate
    m_totalSamples = totalSamples
    m_data = data
End Sub

' Decode all the data into a buffer.
' \param [out] buffer: preallocated buffer to hold decoded data, little-endian. Will *not* be freed, even on error.
' \param [in] bufferSize: takes buffer size in *bytes* as input, to be consistent with oggvorbis,
' \returns how much bytes have been placed into buffer, or nullopt  if error
Public Function decode(buffer() As Byte, ByVal bufferSize As Long, written As Long) As Boolean
    Dim bufferOffset As Long, result As Long, spaceLeft As Long, num_shorts As Long, got As Long
    'unsigned int sampleCount = m_total_time;
    If m_total_time = OV_EINVAL Then
        debug_ 232, LOG_ERROR, "decode", "--sampleCount was OV_EINVAL. The requested bitstream did not exist or the bitstream is nonseekable"
        written = -1                        ' return -1: (size_t)-1
        decode = True
        Exit Function
    End If

    bufferOffset = 0
    Do
        ' (int section = 0)
        spaceLeft = bufferSize - bufferOffset
        ' (VB) ov_read(m_ovfile, buffer + bufferOffset, spaceLeft, OGG_ENDIAN, 2, 1, &section): stb_vorbis's
        ' 16-bit samples, whole frames
        num_shorts = (spaceLeft \ (2 * m_channels)) * m_channels
        result = 0
        If num_shorts > 0 Then
            If UBound(pcm) < num_shorts - 1 Then ReDim pcm(0 To num_shorts - 1)
            got = stb_vorbis_get_samples_short_interleaved(m_file, m_channels, pcm, 0, num_shorts)
            result = got * m_channels * 2
            If result > 0 Then CopyMemory buffer(bufferOffset), pcm(0), result
        End If
        ' (OV_HOLE, OV_EBADLINK, OV_EINVAL: libvorbisfile's errors; stb_vorbis stops at an error as at the end)
        If Not (result >= 0) Then ASSERT_FAILURE "result >= 0", "lib/sound/oggvorbis.cpp:256", "decode", "unexpected error while decoding: %i", result: decode = False: Exit Function
        bufferOffset = bufferOffset + result
    Loop While (result <> 0 And bufferOffset < bufferSize)
    m_bufferSize = bufferSize
    written = bufferOffset
    decode = True
End Function

Public Function totalTime() As Double
    totalTime = m_total_time
End Function

Public Function channels() As Long
    channels = m_channels
End Function

' Sampling rate of the bitstream
Public Function frequency() As Long
    frequency = m_rate
End Function

' Returns the total pcm samples of the physical bitstream.
Public Function totalSamples() As Long
    totalSamples = m_totalSamples
End Function

Private Sub Class_Initialize()
    ReDim pcm(0 To 0)
End Sub

' virtual ~WZVorbisDecoder(): ov_clear, PHYSFS_close (the file is closed already)
Private Sub Class_Terminate()
    If m_file <> 0 Then
        stb_vorbis_close m_file
        m_file = 0
    End If
End Sub

' (VB) codecs.h's WZDecoder
Private Function WZDecoder_decode(buffer() As Byte, ByVal bufferSize As Long, written As Long) As Boolean
    WZDecoder_decode = decode(buffer, bufferSize, written)
End Function

Private Function WZDecoder_totalTime() As Double
    WZDecoder_totalTime = totalTime()
End Function

Private Function WZDecoder_channels() As Long
    WZDecoder_channels = channels()
End Function

Private Function WZDecoder_frequency() As Long
    WZDecoder_frequency = frequency()
End Function
