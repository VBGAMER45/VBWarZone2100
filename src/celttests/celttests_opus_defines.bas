Attribute VB_Name = "celttests_opus_defines"
' (VB) CeltTests.exe only (cluster K, fork O2): the opus_defines.h constants the CELT decoder uses, with the
' names of O3's opus_defines_h.bas (ref/opus/include/opus_defines.h), so the CELT modules build and are tested
' before O3's top level exists. WzSound uses O3's module; this stand-in is not part of it.
Option Explicit

Public Const OPUS_OK As Long = 0
Public Const OPUS_BAD_ARG As Long = -1
Public Const OPUS_BUFFER_TOO_SMALL As Long = -2
Public Const OPUS_INTERNAL_ERROR As Long = -3
Public Const OPUS_INVALID_PACKET As Long = -4
Public Const OPUS_UNIMPLEMENTED As Long = -5
Public Const OPUS_INVALID_STATE As Long = -6
Public Const OPUS_ALLOC_FAIL As Long = -7

Public Const OPUS_SET_COMPLEXITY_REQUEST As Long = 4010
Public Const OPUS_GET_COMPLEXITY_REQUEST As Long = 4011
Public Const OPUS_GET_LOOKAHEAD_REQUEST As Long = 4027
Public Const OPUS_RESET_STATE As Long = 4028
Public Const OPUS_GET_FINAL_RANGE_REQUEST As Long = 4031
Public Const OPUS_GET_PITCH_REQUEST As Long = 4033
Public Const OPUS_SET_PHASE_INVERSION_DISABLED_REQUEST As Long = 4046
Public Const OPUS_GET_PHASE_INVERSION_DISABLED_REQUEST As Long = 4047
