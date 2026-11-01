Attribute VB_Name = "plum_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum (github.com/paullouisageneau/libplum at 47be21b, the commit 4.7.0 pins; Mozilla Public License 2.0, its
' source in ref/libplum) src/plum.c: the library's API (plum_h has include/plum/plum.h's types and constants). A
' callback (plum_log_callback_t, plum_mapping_callback_t) is a function number (0: NULL;
' port_mapping_manager_impl_libplum_c's FN_ constants).
' plum_get_dummytls_certificate and plum_get_dummytls_host (used by no caller; DummyTLS is off: plum_dummytls_c) are
' left out.
' (VB) No threads: plum_poll is the client thread's pass (plum_client_c), run each frame by PortMappingManager_poll.
Option Explicit

Private Const AF_INET As Long = 2

Private client As client_t                      ' static client_t *client = NULL;
Private client_set As Boolean                   ' (VB) client != NULL

' PLUM_EXPORT int plum_init(const plum_config_t *config)
Public Function plum_init(config As plum_config_t) As Long
    If client_set Then
        plum_init = PLUM_ERR_FAILED
        Exit Function
    End If

    plum_log_init
    plum_set_log_level config.log_level
    plum_set_log_handler config.log_callback

    plum_random_init

    dummytls_init
    ' if (config->dummytls_domain) dummytls_set_domain(config->dummytls_domain): (VB) left out (plum_dummytls_c);
    ' 4.7.0 passes NULL

    If Not client_create(client) Then
        plum_init = PLUM_ERR_FAILED
        Exit Function
    End If
    client_set = True

    plum_init = PLUM_ERR_SUCCESS
End Function

' PLUM_EXPORT int plum_cleanup()
Public Function plum_cleanup() As Long
    If Not client_set Then
        plum_cleanup = PLUM_ERR_FAILED
        Exit Function
    End If

    client_destroy client
    client_set = False

    dummytls_cleanup

    plum_log_cleanup
    plum_random_cleanup

    plum_cleanup = PLUM_ERR_SUCCESS
End Function

' PLUM_EXPORT int plum_create_mapping(const plum_mapping_t *mapping, plum_mapping_callback_t callback)
Public Function plum_create_mapping(mapping As plum_mapping_t, ByVal callback As Long) As Long
    Dim id As Long
    If Not client_set Then
        plum_create_mapping = PLUM_ERR_FAILED
        Exit Function
    End If

    id = client_add_mapping(client, mapping, callback)
    If id < 0 Then
        plum_create_mapping = PLUM_ERR_FAILED
        Exit Function
    End If

    ' Ensure client is started
    If client_start(client) < 0 Then
        plum_create_mapping = PLUM_ERR_FAILED
        Exit Function
    End If

    plum_create_mapping = id
End Function

' PLUM_EXPORT int plum_query_mapping(int id, plum_state_t *state, plum_mapping_t *mapping)
Public Function plum_query_mapping(ByVal id As Long, state As Long, mapping As plum_mapping_t) As Long
    If Not client_set Then
        plum_query_mapping = PLUM_ERR_FAILED
        Exit Function
    End If

    If id < 0 Then
        plum_query_mapping = PLUM_ERR_INVALID
        Exit Function
    End If

    If client_get_mapping(client, id, state, mapping) < 0 Then
        plum_query_mapping = PLUM_ERR_NOT_AVAIL
        Exit Function
    End If

    plum_query_mapping = PLUM_ERR_SUCCESS
End Function

' PLUM_EXPORT int plum_destroy_mapping(int id)
Public Function plum_destroy_mapping(ByVal id As Long) As Long
    If Not client_set Then
        plum_destroy_mapping = PLUM_ERR_FAILED
        Exit Function
    End If

    If id < 0 Then
        plum_destroy_mapping = PLUM_ERR_INVALID
        Exit Function
    End If

    If client_remove_mapping(client, id) < 0 Then
        plum_destroy_mapping = PLUM_ERR_NOT_AVAIL
        Exit Function
    End If

    plum_destroy_mapping = PLUM_ERR_SUCCESS
End Function

' PLUM_EXPORT int plum_get_local_address(char *buffer, size_t size)
Public Function plum_get_local_address(buffer As String, ByVal size As Long) As Long
    Dim record As addr_record_t, len_ As Long
    If net_get_default_interface(AF_INET, record) < 0 Then
        plum_get_local_address = PLUM_ERR_NOT_AVAIL
        Exit Function
    End If

    len_ = addr_get_host(record.addr, buffer, size)
    If len_ < 0 Then
        plum_get_local_address = PLUM_ERR_FAILED
        Exit Function
    End If

    plum_get_local_address = len_
End Function

' (VB) the client thread's pass (the thread runs while a client exists and its client_run has not returned)
Public Sub plum_poll()
    If client_set Then client_thread_pass client
End Sub
