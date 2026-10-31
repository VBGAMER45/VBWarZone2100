VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZOpusDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/sound/oggopus.cpp, oggopus.h: class WZOpusDecoder (codecs.h's WZDecoder) over opusfile (opusfile_c).
' WZOpusDecoder::fromFilename, fromMemoryBuffer and openOpusInternal are oggopus_c's WZOpusDecoder_fromFilename,
' WZOpusDecoder_fromMemoryBuffer and openOpusInternal. decode's optional<size_t>: decode returns false for
' nullopt, the bytes in written. m_duration is never set (4.7.0's totalTime() is 0 for an Opus stream).
Option Explicit

Implements WZDecoder

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' we need a small buffer to convert from big to little endian
' sample_rate (48000/sec) * depth (2 bytes) * 120 milliseconds
Private Const TMP_BUF As Long = 48& * 2 * 120

Private m_of As Long                            ' OggOpusFile* = nullptr (opusfile_c's handle)

Private m_data As WZAudioDataResourceInterface  ' std::unique_ptr<WZAudioDataResourceInterface>

Private m_bufferSize As Double                  ' int64_t = 0
' Because timestamps in Opus are fixed at 48 kHz, there is no need for a separate function to convert this to seconds
' nb samples = nb seconds
Private m_duration As Double                    ' int64_t = 0
' nb links. We do not supported chained files (which have > 1 link), because I am lazy
'	const int m_nlinks = 1;
'	const OpusHead *m_head = nullptr;
Private Const m_nchannels As Long = 2

Private pcm() As Integer                        ' (VB) decode's opus_int16 pcm[TMP_BUF]

' WZOpusDecoder(std::unique_ptr<WZAudioDataResourceInterface> data, OggOpusFile* ovf, const OpusHead *head)
Friend Sub construct(ByVal data As WZAudioDataResourceInterface, ByVal ovf As Long)
    m_of = ovf
    Set m_data = data
    ' (m_head(head) is commented out)
End Sub

Public Function decode(buffer() As Byte, ByVal bufferSize As Long, written As Long) As Boolean
    Dim bufferOffset As Long, samples_per_chan As Long, spaceLeft As Long, toRead As Long, minSize As Long, n As Long

    bufferOffset = 0
    samples_per_chan = 0
    m_bufferSize = bufferSize
    ' Decode PCM data into the buffer until there is nothing to decode left OR untill we hit bufferSize limit
    Do
        spaceLeft = bufferSize - bufferOffset
        If Not (spaceLeft >= 0) Then ASSERT_FAILURE "spaceLeft <= static_cast<size_t>(std::numeric_limits<int>::max())", "lib/sound/oggopus.cpp:338", "decode", "spaceLeft (%zu) exceeds int::max", spaceLeft
        If TMP_BUF < spaceLeft Then toRead = TMP_BUF Else toRead = spaceLeft

        ' Note: the return value is the number of *samples per channel*, not *bytes* !!
        samples_per_chan = op_read_stereo(m_of, pcm, 0, toRead)
        Select Case samples_per_chan
        Case OP_HOLE
            debug_ 346, LOG_ERROR, "decode", "There was a hole in the data, and some samples may have been skipped. Call this function again to continue decoding past the hole."
            decode = False
            Exit Function
        Case OP_EREAD
            debug_ 349, LOG_ERROR, "decode", "An underlying read operation failed. This may signal a truncation attack from an <https:> source."
            decode = False
            Exit Function
        Case OP_EFAULT
            debug_ 352, LOG_ERROR, "decode", "An internal memory allocation failed."
            decode = False
            Exit Function
        Case OP_EIMPL
            debug_ 355, LOG_ERROR, "decode", "An unseekable stream encountered a new link that used a feature that is not implemented, such as an unsupported channel family."
            decode = False
            Exit Function
        Case OP_EINVAL
            debug_ 358, LOG_ERROR, "decode", "The stream was only partially open."
            decode = False
            Exit Function
        Case OP_ENOTFORMAT
            debug_ 361, LOG_ERROR, "decode", "An unseekable stream encountered a new link that did not have any logical Opus streams in it. "
            decode = False
            Exit Function
        Case OP_EBADHEADER
            debug_ 364, LOG_ERROR, "decode", "An unseekable stream encountered a new link with a required header packet that was not properly formatted, contained illegal values, or was missing altogether."
            decode = False
            Exit Function
        Case OP_EVERSION
            debug_ 367, LOG_ERROR, "decode", "An unseekable stream encountered a new link with an ID header that contained an unrecognized version number."
            decode = False
            Exit Function
        Case OP_EBADPACKET
            debug_ 370, LOG_ERROR, "decode", "Failed to properly decode the next packet."
            decode = False
            Exit Function
        Case OP_EBADLINK
            debug_ 373, LOG_ERROR, "decode", "We failed to find data we had seen before."
            decode = False
            Exit Function
        Case OP_EBADTIMESTAMP
            debug_ 376, LOG_ERROR, "decode", "An unseekable stream encountered a new link with a starting timestamp that failed basic validity checks."
            decode = False
            Exit Function
        Case Else
        End Select
        If Not (samples_per_chan >= 0) Then ASSERT_FAILURE "samples_per_chan >= 0", "lib/sound/oggopus.cpp:381", "decode", "Unexpected Opus error %i", samples_per_chan
        ' convert to little endian
        ' for (int si = 0; si < m_nchannels * samples_per_chan; si++) { buffer[bufferOffset + 0] = (uint8_t)(pcm[si] & 0xFF);
        '   buffer[bufferOffset + 1] = (uint8_t)(pcm[si] >> 8 & 0xFF); bufferOffset +=2; }: (VB) the same bytes
        n = m_nchannels * samples_per_chan * 2
        If n > 0 Then
            ' (VB) the C writes past the buffer's end when the decoded samples are more than the space left (toRead
            ' counts shorts where spaceLeft counts bytes: with 60 ms packets in openal_track's 16 KiB buffers, the
            ' second packet goes 5408 bytes past it), and returns the whole count, which alBufferData then reads:
            ' the port grows the buffer instead, so the samples are the same and nothing else is overwritten
            If bufferOffset + n > UBound(buffer) + 1 Then ReDim Preserve buffer(0 To bufferOffset + n - 1)
            CopyMemory buffer(bufferOffset), pcm(0), n
            bufferOffset = bufferOffset + n
        End If
        If bufferSize < TMP_BUF Then minSize = bufferSize Else minSize = TMP_BUF
    Loop While (samples_per_chan <> 0 And bufferOffset < minSize)
    written = bufferOffset
    decode = True
End Function

Public Function totalTime() As Double
    totalTime = m_duration
End Function

Public Function channels() As Long
    channels = 2
End Function

' Sampling rate of the bitstream
Public Function frequency() As Long
    frequency = 48000
End Function

Private Sub Class_Initialize()
    ReDim pcm(0 To TMP_BUF - 1)
End Sub

' virtual ~WZOpusDecoder()
Private Sub Class_Terminate()
    op_free m_of
    m_of = 0
    Set m_data = Nothing                        ' m_data.reset()
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
