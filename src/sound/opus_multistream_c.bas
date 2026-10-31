Attribute VB_Name = "opus_multistream_c"
' libopus 1.5.2 (ref/opus/src): opus_multistream.c whole (validate_layout, get_left_channel,
' get_right_channel, get_mono_channel), with opus_multistream.h's request constants.
Option Explicit

' opus_multistream.h
Public Const OPUS_MULTISTREAM_GET_ENCODER_STATE_REQUEST As Long = 5120
Public Const OPUS_MULTISTREAM_GET_DECODER_STATE_REQUEST As Long = 5122

Public Function validate_layout(layout As ChannelLayout) As Long
    Dim i As Long, max_channel As Long

    max_channel = layout.nb_streams + layout.nb_coupled_streams
    If max_channel > 255 Then
        Exit Function
    End If
    For i = 0 To layout.nb_channels - 1
        If layout.mapping(i) >= max_channel And layout.mapping(i) <> 255 Then
            Exit Function
        End If
    Next
    validate_layout = 1
End Function

Public Function get_left_channel(layout As ChannelLayout, ByVal stream_id As Long, ByVal prev As Long) As Long
    Dim i As Long
    If prev < 0 Then i = 0 Else i = prev + 1
    Do While i < layout.nb_channels
        If layout.mapping(i) = stream_id * 2 Then
            get_left_channel = i
            Exit Function
        End If
        i = i + 1
    Loop
    get_left_channel = -1
End Function

Public Function get_right_channel(layout As ChannelLayout, ByVal stream_id As Long, ByVal prev As Long) As Long
    Dim i As Long
    If prev < 0 Then i = 0 Else i = prev + 1
    Do While i < layout.nb_channels
        If layout.mapping(i) = stream_id * 2 + 1 Then
            get_right_channel = i
            Exit Function
        End If
        i = i + 1
    Loop
    get_right_channel = -1
End Function

Public Function get_mono_channel(layout As ChannelLayout, ByVal stream_id As Long, ByVal prev As Long) As Long
    Dim i As Long
    If prev < 0 Then i = 0 Else i = prev + 1
    Do While i < layout.nb_channels
        If layout.mapping(i) = stream_id + layout.nb_coupled_streams Then
            get_mono_channel = i
            Exit Function
        End If
        i = i + 1
    Loop
    get_mono_channel = -1
End Function
