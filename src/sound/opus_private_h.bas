Attribute VB_Name = "opus_private_h"
' libopus 1.5.2 (ref/opus/src): opus_private.h's decoder parts: ChannelLayout, the MODE_* constants and
' align() (the repacketizer's, the extensions' and the encoders' declarations are not ported: nothing in the
' decoders reaches them; opus_packet_parse_impl is in opus_c, validate_layout and get_*_channel in
' opus_multistream_c).
Option Explicit

Public Type ChannelLayout
    nb_channels As Long
    nb_streams As Long
    nb_coupled_streams As Long
    mapping(0 To 255) As Byte
End Type

' typedef enum { MAPPING_TYPE_NONE, MAPPING_TYPE_SURROUND, MAPPING_TYPE_AMBISONICS } MappingType; (the encoder's)

Public Const MODE_SILK_ONLY As Long = 1000
Public Const MODE_HYBRID As Long = 1001
Public Const MODE_CELT_ONLY As Long = 1002

Public Const OPUS_SET_VOICE_RATIO_REQUEST As Long = 11018
Public Const OPUS_GET_VOICE_RATIO_REQUEST As Long = 11019

Public Const OPUS_SET_FORCE_MODE_REQUEST As Long = 11002

'/* Make sure everything is properly aligned. */
' (VB) struct foo {char c; union { void* p; opus_int32 i; opus_val32 v; } u;}: 4 on a 32-bit build
Public Function align(ByVal i As Long) As Long
    Const alignment As Long = 4
    '/* Optimizing compilers should optimize div and multiply into and
    '   for all sensible alignment values. */
    align = ((i + alignment - 1) \ alignment) * alignment
End Function
