Attribute VB_Name = "plum_protocol_h"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/protocol.h: the protocols' shared types and error codes. A protocol_t (the table of a protocol's
' functions) is a protocol number, PROTOCOL_PCP .. PROTOCOL_NOPROTOCOL (plum_client_c calls its functions).
' protocol_state_t's void *impl is a flag: the protocol's state is its module's (one protocol at a time).
Option Explicit

' typedef struct protocol_state_t
Public Type protocol_state_t
    gateway As addr_record_t
    impl As Boolean                             ' void *impl: (VB) allocated (the state is the protocol module's)
End Type

Public Const PROTOCOL_ERR_SUCCESS As Long = 0
Public Const PROTOCOL_ERR_UNKNOWN As Long = -1
Public Const PROTOCOL_ERR_INTERRUPTED As Long = -2
Public Const PROTOCOL_ERR_TIMEOUT As Long = -3
Public Const PROTOCOL_ERR_RESET As Long = -4
Public Const PROTOCOL_ERR_PROTOCOL_FAILED As Long = -5
Public Const PROTOCOL_ERR_NETWORK_FAILED As Long = -6
Public Const PROTOCOL_ERR_INSUFF_RESOURCES As Long = -7
Public Const PROTOCOL_ERR_UNSUPP_PROTOCOL As Long = -8
Public Const PROTOCOL_ERR_UNSUPP_VERSION As Long = -9
Public Const PROTOCOL_ERR_SKIPPED As Long = -10
Public Const PROTOCOL_ERR_NEED_DOWNGRADE As Long = -11

' typedef enum protocol_map_state_t
Public Const PROTOCOL_MAP_STATE_SUCCESS As Long = 0
Public Const PROTOCOL_MAP_STATE_FAILURE As Long = 1

' typedef struct protocol_map_output_t
Public Type protocol_map_output_t
    state As Long
    external_addr As addr_record_t
    refresh_timestamp As Double
    impl_record As Variant                      ' void *impl_record: Empty (NULL) or a Byte array (malloc'd)
End Type
