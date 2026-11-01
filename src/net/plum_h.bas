Attribute VB_Name = "plum_h"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum include/plum/plum.h: the API's types and constants (plum_c has its functions). A callback
' (plum_log_callback_t, plum_mapping_callback_t) is a function number (0: NULL).
Option Explicit

Public Const PLUM_ERR_SUCCESS As Long = 0
Public Const PLUM_ERR_INVALID As Long = -1      ' invalid argument
Public Const PLUM_ERR_FAILED As Long = -2       ' runtime error
Public Const PLUM_ERR_NOT_AVAIL As Long = -3    ' element not available

' typedef enum plum_log_level_t
Public Const PLUM_LOG_LEVEL_VERBOSE As Long = 0
Public Const PLUM_LOG_LEVEL_DEBUG As Long = 1
Public Const PLUM_LOG_LEVEL_INFO As Long = 2
Public Const PLUM_LOG_LEVEL_WARN As Long = 3
Public Const PLUM_LOG_LEVEL_ERROR As Long = 4
Public Const PLUM_LOG_LEVEL_FATAL As Long = 5
Public Const PLUM_LOG_LEVEL_NONE As Long = 6

' typedef struct plum_config_t
Public Type plum_config_t
    log_level As Long
    log_callback As Long                        ' plum_log_callback_t: 0 (NULL) means stdout
    dummytls_domain As String                   ' vbNullString (NULL) means disabled
End Type

' typedef enum plum_ip_protocol_t
Public Const PLUM_IP_PROTOCOL_TCP As Long = 0
Public Const PLUM_IP_PROTOCOL_UDP As Long = 1

' typedef enum plum_state_t
Public Const PLUM_STATE_DESTROYED As Long = 0
Public Const PLUM_STATE_PENDING As Long = 1
Public Const PLUM_STATE_SUCCESS As Long = 2
Public Const PLUM_STATE_FAILURE As Long = 3
Public Const PLUM_STATE_DESTROYING As Long = 4

Public Const PLUM_MAX_HOST_LEN As Long = 256
Public Const PLUM_MAX_ADDRESS_LEN As Long = 64

' typedef struct plum_mapping_t
Public Type plum_mapping_t
    protocol As Long                            ' plum_ip_protocol_t
    internal_port As Long                       ' uint16_t
    external_port As Long                       ' uint16_t
    external_host As String                     ' char external_host[PLUM_MAX_HOST_LEN]
    user_ptr As Long                            ' void *user_ptr
End Type
