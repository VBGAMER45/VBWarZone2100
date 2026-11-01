Attribute VB_Name = "plum_client_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/client.c, client.h: the client (its mappings and the protocol it runs) and its thread's loop.
' (VB) No threads: client_start does not start a thread; the client thread's work is client_thread_pass, run by
' plum_c's plum_poll (each frame: PortMappingManager_poll, from NETflush). client_run and client_run_protocol are
' resumable as the protocols' waits are (plum_socket_h's PLUM_PENDING): a pass runs until a call would wait, and the
' next pass resumes it; client_run returning ends the thread. thread_join (client_destroy) runs passes until then.
' The mutexes are left out (one thread). The table of protocols (const protocol_t protocols[]) is a protocol
' number and the protocol_* procedures below; client->protocol is that number + 1 (0: NULL). A mapping's
' plum_mapping_callback_t is a function number (0: NULL): FN_PlumMappingCallback.
Option Explicit

Private Declare Function ws_WSAStartup Lib "ws2_32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSAData As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const AF_INET As Long = 2
Private Const AF_UNSPEC As Long = 0

Public Const CLIENT_MAX_DISCOVER_TIMEOUT As Long = 10000   ' msecs
Public Const CLIENT_MAX_MAPPING_TIMEOUT As Long = 10000    ' msecs
Public Const CLIENT_RECHECK_PERIOD As Long = 300000        ' msecs

Private Const DEFAULT_MAPPINGS_SIZE As Long = 16

Private Const PROTOCOL_PCP As Long = 0
Private Const PROTOCOL_UPNP As Long = 1
Private Const PROTOCOL_NOPROTOCOL As Long = 2
Private Const PROTOCOLS_COUNT As Long = 3

' typedef struct client_mapping
Public Type client_mapping_t
    protocol As Long                            ' plum_ip_protocol_t
    internal_port As Long                       ' uint16_t
    suggested_addr As addr_record_t
    external_addr As addr_record_t
    callback As Long                            ' plum_mapping_callback_t: a function number (0: NULL)
    user_ptr As Long
    state As Long                               ' plum_state_t
    refresh_timestamp As Double
    impl_record As Variant                      ' void *impl_record: Empty (NULL) or a Byte array
End Type

' typedef struct client_t
Public Type client_t
    mappings() As client_mapping_t
    mappings_size As Long
    protocol As Long                            ' const protocol_t *protocol: the protocol number + 1 (0: NULL)
    protocol_state As protocol_state_t
    is_started As Boolean
    is_stopping As Boolean                      ' atomic(bool)
    thread As Boolean                           ' thread_t thread: (VB) its passes run (client_run has not returned)
End Type

' (VB) client_run's and client_run_protocol's state while they wait
Private run_at As Long                          ' client_run: 1 = in client_run_protocol
Private run_old_local As addr_record_t
Private run_protocol_num As Long
Private rp_at As Long                           ' client_run_protocol: 1 discover, 2 unmap, 3 map, 4 idle
Private rp_end_timestamp As Double
Private rp_next_timestamp As Double
Private rp_i As Long
Private rp_mapping As client_mapping_t
Private rp_output As protocol_map_output_t

Private wsaData(0 To 511) As Byte

' (VB) static const protocol_t protocols[PROTOCOLS_COUNT]: protocols[p].init(state) ...
Private Function protocol_init(ByVal p As Long, state As protocol_state_t) As Long
    Select Case p
    Case PROTOCOL_PCP: protocol_init = pcp_init(state)
    Case PROTOCOL_UPNP: protocol_init = upnp_init(state)
    Case PROTOCOL_NOPROTOCOL: protocol_init = noprotocol_init(state)
    End Select
End Function

Private Function protocol_cleanup(ByVal p As Long, state As protocol_state_t) As Long
    Select Case p
    Case PROTOCOL_PCP: protocol_cleanup = pcp_cleanup(state)
    Case PROTOCOL_UPNP: protocol_cleanup = upnp_cleanup(state)
    Case PROTOCOL_NOPROTOCOL: protocol_cleanup = noprotocol_cleanup(state)
    End Select
End Function

Private Function protocol_discover(ByVal p As Long, state As protocol_state_t, ByVal duration As Double) As Long
    Select Case p
    Case PROTOCOL_PCP: protocol_discover = pcp_discover(state, duration)
    Case PROTOCOL_UPNP: protocol_discover = upnp_discover(state, duration)
    Case PROTOCOL_NOPROTOCOL: protocol_discover = noprotocol_discover(state, duration)
    End Select
End Function

Private Function protocol_map(ByVal p As Long, state As protocol_state_t, mapping As client_mapping_t, output As protocol_map_output_t, ByVal duration As Double) As Long
    Select Case p
    Case PROTOCOL_PCP: protocol_map = pcp_map(state, mapping, output, duration)
    Case PROTOCOL_UPNP: protocol_map = upnp_map(state, mapping, output, duration)
    Case PROTOCOL_NOPROTOCOL: protocol_map = noprotocol_map(state, mapping, output, duration)
    End Select
End Function

Private Function protocol_unmap(ByVal p As Long, state As protocol_state_t, mapping As client_mapping_t, ByVal duration As Double) As Long
    Select Case p
    Case PROTOCOL_PCP: protocol_unmap = pcp_unmap(state, mapping, duration)
    Case PROTOCOL_UPNP: protocol_unmap = upnp_unmap(state, mapping, duration)
    Case PROTOCOL_NOPROTOCOL: protocol_unmap = noprotocol_unmap(state, mapping, duration)
    End Select
End Function

Private Function protocol_idle(ByVal p As Long, state As protocol_state_t, ByVal duration As Double) As Long
    Select Case p
    Case PROTOCOL_PCP: protocol_idle = pcp_idle(state, duration)
    Case PROTOCOL_UPNP: protocol_idle = upnp_idle(state, duration)
    Case PROTOCOL_NOPROTOCOL: protocol_idle = noprotocol_idle(state, duration)
    End Select
End Function

Private Function protocol_interrupt(ByVal p As Long, state As protocol_state_t, ByVal hard As Boolean) As Long
    Select Case p
    Case PROTOCOL_PCP: protocol_interrupt = pcp_interrupt(state, hard)
    Case PROTOCOL_UPNP: protocol_interrupt = upnp_interrupt(state, hard)
    Case PROTOCOL_NOPROTOCOL: protocol_interrupt = noprotocol_interrupt(state, hard)
    End Select
End Function

' static int find_empty_mapping_index(client_t *client)
Private Function find_empty_mapping_index(client As client_t) As Long
    Dim i As Long
    For i = 0 To client.mappings_size - 1
        If client.mappings(i).internal_port = 0 Then
            find_empty_mapping_index = i
            Exit Function
        End If
    Next

    find_empty_mapping_index = -1
End Function

' static void import_mapping(const plum_mapping_t *mapping, plum_mapping_callback_t callback, client_mapping_t *cm)
Private Sub import_mapping(mapping As plum_mapping_t, ByVal callback As Long, cm As client_mapping_t)
    Dim zeroMapping As client_mapping_t, zeroAddr As addr_record_t, set_ As Boolean
    cm = zeroMapping
    cm.callback = callback
    cm.protocol = mapping.protocol
    cm.internal_port = mapping.internal_port
    cm.user_ptr = mapping.user_ptr

    If mapping.external_host <> "" Then set_ = addr_set(AF_UNSPEC, mapping.external_host, mapping.external_port, cm.suggested_addr) >= 0
    If Not set_ Then
        If mapping.external_port <> 0 Then set_ = addr_set(AF_INET, "0.0.0.0", mapping.external_port, cm.suggested_addr) >= 0
        If Not set_ Then cm.suggested_addr = zeroAddr
    End If
End Sub

' static void export_mapping(const client_mapping_t *cm, plum_mapping_t *mapping)
Private Sub export_mapping(cm As client_mapping_t, mapping As plum_mapping_t)
    Dim zeroMapping As plum_mapping_t
    mapping = zeroMapping
    mapping.protocol = cm.protocol
    mapping.internal_port = cm.internal_port
    mapping.user_ptr = cm.user_ptr
    If cm.external_addr.len_ > 0 Then
        ' const struct sockaddr *sa = (const struct sockaddr *)&cm->external_addr.addr;
        mapping.external_port = addr_get_port(cm.external_addr.addr)
        If dummytls_get_host(cm.external_addr.addr, mapping.external_host, PLUM_MAX_HOST_LEN) < 0 Then _
            addr_get_host cm.external_addr.addr, mapping.external_host, PLUM_MAX_HOST_LEN
    End If
End Sub

' thread_return_t THREAD_CALL client_thread_entry(void *arg)
' (VB) the thread's pass (see above): client_run from where the last pass stopped; the thread ends when it returns
Public Sub client_thread_pass(client As client_t)
    If Not client.thread Then Exit Sub
    If client_run(client) <> PLUM_PENDING Then client.thread = False
End Sub

' (VB) thread_join(client->thread, NULL): the thread's passes until client_run returns
Private Sub thread_join(client As client_t)
    Do While client.thread
        client_thread_pass client
        If client.thread Then Sleep 1
    Loop
End Sub

' client_t *client_create(void): (VB) into client; False for NULL
Public Function client_create(client As client_t) As Boolean
    Dim zeroClient As client_t
    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Creating client"

    If ws_WSAStartup(&H202, wsaData(0)) <> 0 Then   ' MAKEWORD(2, 2)
        plum_log_write PLUM_LOG_LEVEL_FATAL, "WSAStartup failed"
        Exit Function
    End If

    client = zeroClient                         ' malloc(sizeof(client_t)); memset(client, 0, sizeof(*client));

    ReDim client.mappings(0 To DEFAULT_MAPPINGS_SIZE - 1)
    client.mappings_size = DEFAULT_MAPPINGS_SIZE

    ' mutex_init(&client->mappings_mutex, MUTEX_RECURSIVE); mutex_init(&client->protocol_mutex, 0);

    client_create = True
End Function

' void client_destroy(client_t *client)
Public Sub client_destroy(client As client_t)
    Dim zeroClient As client_t
    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Destroying client"

    If client.is_started Then
        client_interrupt client, True           ' stop
        thread_join client
    End If

    ' mutex_destroy(&client->mappings_mutex); mutex_destroy(&client->protocol_mutex);

    client = zeroClient                         ' free(client->mappings); free(client);

    WSACleanup

    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Destroyed client"
End Sub

' int client_start(client_t *client)
Public Function client_start(client As client_t) As Long
    If client.is_started Then
        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Client is already started"
        client_start = 0
        Exit Function
    End If

    ' thread_init(&client->thread, client_thread_entry, client): (VB) the passes run client_run from its start
    run_at = 0
    rp_at = 0
    client.thread = True

    client.is_started = True
    client_start = 0
End Function

' int client_add_mapping(client_t *client, const plum_mapping_t *mapping, plum_mapping_callback_t callback)
Public Function client_add_mapping(client As client_t, mapping As plum_mapping_t, ByVal callback As Long) As Long
    Dim i As Long
    i = find_empty_mapping_index(client)
    If i < 0 Then
        i = client.mappings_size

        ' realloc(client->mappings, client->mappings_size * 2 * sizeof(client_mapping_t)) and memset of the new half
        ReDim Preserve client.mappings(0 To client.mappings_size * 2 - 1)
        client.mappings_size = client.mappings_size * 2
    End If

    import_mapping mapping, callback, client.mappings(i)
    client.mappings(i).state = PLUM_STATE_PENDING
    client.mappings(i).refresh_timestamp = 0

    plum_log_write PLUM_LOG_LEVEL_INFO, "Added mapping %d for internal port %hu (callback=%d)", i, client.mappings(i).internal_port, _
                   IIf(client.mappings(i).callback <> 0, 1, 0)

    client_interrupt client, False
    client_add_mapping = i
End Function

' int client_get_mapping(client_t *client, int i, plum_state_t *state, plum_mapping_t *mapping)
Public Function client_get_mapping(client As client_t, ByVal i As Long, state As Long, mapping As plum_mapping_t) As Long
    If i >= client.mappings_size Then
        client_get_mapping = -1
        Exit Function
    End If
    If client.mappings(i).state = PLUM_STATE_DESTROYED Then
        client_get_mapping = -1
        Exit Function
    End If

    export_mapping client.mappings(i), mapping

    state = client.mappings(i).state

    client_get_mapping = 0
End Function

' int client_remove_mapping(client_t *client, int i)
Public Function client_remove_mapping(client As client_t, ByVal i As Long) As Long
    If i >= client.mappings_size Then
        client_remove_mapping = -1
        Exit Function
    End If
    If client.mappings(i).state = PLUM_STATE_DESTROYED Then
        client_remove_mapping = -1
        Exit Function
    End If

    client.mappings(i).state = PLUM_STATE_DESTROYING

    plum_log_write PLUM_LOG_LEVEL_INFO, "Removed mapping %d for internal port %hu", i, client.mappings(i).internal_port

    client_interrupt client, False
    client_remove_mapping = 0
End Function

' static void trigger_mapping_callback(const client_mapping_t *cm, int i)
Private Sub trigger_mapping_callback(client As client_t, ByVal i As Long)
    Dim mapping As plum_mapping_t, callback As Long, state As Long
    export_mapping client.mappings(i), mapping
    callback = client.mappings(i).callback
    state = client.mappings(i).state
    If callback <> 0 Then plum_mapping_callback_call callback, i, state, mapping
End Sub

' (VB) cm->callback(i, cm->state, &mapping): the function number's function
Private Sub plum_mapping_callback_call(ByVal callback As Long, ByVal id As Long, ByVal state As Long, mapping As plum_mapping_t)
    Select Case callback
    Case FN_PlumMappingCallback: PlumMappingCallback id, state, mapping
    End Select
End Sub

' static void update_mapping(client_mapping_t *cm, int i, plum_state_t state, const addr_record_t *external)
' (VB) hasExternal: external is not NULL
Private Sub update_mapping(client As client_t, ByVal i As Long, ByVal state As Long, external As addr_record_t, ByVal hasExternal As Boolean)
    Dim changed As Boolean, zeroAddr As addr_record_t
    If client.mappings(i).state = PLUM_STATE_DESTROYED Or client.mappings(i).state = PLUM_STATE_DESTROYING Then Exit Sub

    changed = False
    If state = PLUM_STATE_SUCCESS And hasExternal Then
        changed = Not addr_record_is_equal(client.mappings(i).external_addr, external, True)
        client.mappings(i).external_addr = external
    Else
        client.mappings(i).external_addr = zeroAddr
    End If

    If state <> client.mappings(i).state Or changed Then
        client.mappings(i).state = state

        If state <> PLUM_STATE_PENDING And state <> PLUM_STATE_DESTROYING Then trigger_mapping_callback client, i
    End If
End Sub

' static void destroy_mapping(client_mapping_t *cm, int i)
Private Sub destroy_mapping(client As client_t, ByVal i As Long)
    Dim zeroAddr As addr_record_t, zeroMapping As client_mapping_t
    client.mappings(i).external_addr = zeroAddr
    update_mapping client, i, PLUM_STATE_DESTROYED, zeroAddr, False
    client.mappings(i).impl_record = Empty      ' free(cm->impl_record);
    client.mappings(i) = zeroMapping
End Sub

' static bool has_destroying_mappings(client_t *client)
Private Function has_destroying_mappings(client As client_t) As Boolean
    Dim i As Long
    For i = 0 To client.mappings_size - 1
        If client.mappings(i).state = PLUM_STATE_DESTROYING Then
            has_destroying_mappings = True
            Exit Function
        End If
    Next
    has_destroying_mappings = False
End Function

' static void reset_protocol(client_t *client)
Private Sub reset_protocol(client As client_t)
    Dim i As Long, zeroAddr As addr_record_t
    ' protocol_mutex must be locked

    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Resetting protocol state"

    If client.protocol <> 0 Then
        protocol_cleanup client.protocol - 1, client.protocol_state
        client.protocol = 0
    End If

    ' Also reset timestamps and records
    For i = 0 To client.mappings_size - 1
        If client.mappings(i).state = PLUM_STATE_DESTROYED Then GoTo continue_

        client.mappings(i).refresh_timestamp = 0
        client.mappings(i).impl_record = Empty
        client.mappings(i).external_addr = zeroAddr
        If client.mappings(i).state = PLUM_STATE_DESTROYING Then
            destroy_mapping client, i           ' as good as destroyed now
        Else
            update_mapping client, i, PLUM_STATE_PENDING, zeroAddr, False
        End If
continue_:
    Next
End Sub

' void client_run(client_t *client)
' (VB) PLUM_PENDING when the pass stops at a wait (the next pass resumes it), 0 when the thread's function returns
Public Function client_run(client As client_t) As Long
    Dim local_ As addr_record_t, changed As Boolean, err_ As Long, zeroAddr As addr_record_t
    If run_at = 1 Then run_at = 0: GoTo resume_protocol      ' (VB) resumed

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Starting client thread"

    run_old_local = zeroAddr
    run_protocol_num = 0
    Do While True
        If net_get_default_interface(AF_INET, local_) < 0 Then
            plum_log_write PLUM_LOG_LEVEL_ERROR, "Unable to get default interface address"
            Exit Do
        End If

        changed = False
        If run_old_local.len_ > 0 Then changed = Not addr_record_is_equal(run_old_local, local_, False)
        run_old_local = local_
        If changed Then
            plum_log_write PLUM_LOG_LEVEL_INFO, "Local address changed, restarting"
            reset_protocol client
            run_protocol_num = 0
        End If

        If run_protocol_num <> PROTOCOL_NOPROTOCOL And addr_is_public(local_.addr) Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Local address is public, skipping discovery"
            reset_protocol client
            run_protocol_num = PROTOCOL_NOPROTOCOL
        End If

        err_ = PROTOCOL_ERR_SUCCESS
        If client.protocol = 0 Then
            client.protocol = run_protocol_num + 1  ' protocols + protocol_num
            err_ = protocol_init(client.protocol - 1, client.protocol_state)
            If err_ <> PROTOCOL_ERR_SUCCESS Then client.protocol = 0
        End If

        If err_ = PROTOCOL_ERR_SUCCESS Then
resume_protocol:
            err_ = client_run_protocol(client, client.protocol - 1, client.protocol_state, CLIENT_RECHECK_PERIOD)
            If err_ = PLUM_PENDING Then run_at = 1: client_run = PLUM_PENDING: Exit Function
        End If

        If err_ = PROTOCOL_ERR_SUCCESS Or err_ = PROTOCOL_ERR_INTERRUPTED Then
            If err_ = PROTOCOL_ERR_INTERRUPTED Then plum_log_write PLUM_LOG_LEVEL_DEBUG, "Protocol was interrupted"

            If client.is_stopping Then
                If Not has_destroying_mappings(client) Then
                    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Client is stopping, exiting"
                    Exit Do
                End If
                plum_log_write PLUM_LOG_LEVEL_DEBUG, "Mappings are marked for destruction, continuing"
            End If
            GoTo continue_
        End If

        If client.is_stopping Then Exit Do

        reset_protocol client

        If err_ = PROTOCOL_ERR_RESET Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Protocol was reset"
            GoTo continue_
        End If

        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Protocol failed"

        run_protocol_num = run_protocol_num + 1
        If run_protocol_num >= PROTOCOLS_COUNT Then
            plum_log_write PLUM_LOG_LEVEL_FATAL, "Client failed, exiting"
            Exit Do
        End If
continue_:
    Loop

    reset_protocol client

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Exiting client thread"
    client_run = 0
End Function

' int client_run_protocol(client_t *client, const protocol_t *protocol, protocol_state_t *protocol_state,
'                         timediff_t duration)
Public Function client_run_protocol(client As client_t, ByVal protocol As Long, protocol_state As protocol_state_t, ByVal duration As Double) As Long
    Dim err_ As Long, external_str As String, refresh_delay As Long, now As Double, diff As Double, zeroAddr As addr_record_t
    Select Case rp_at                           ' (VB) resumed
    Case 1: rp_at = 0: GoTo resume_discover
    Case 2: rp_at = 0: GoTo resume_unmap
    Case 3: rp_at = 0: GoTo resume_map
    Case 4: rp_at = 0: GoTo resume_idle
    End Select
    rp_end_timestamp = current_timestamp() + duration

resume_discover:
    err_ = protocol_discover(protocol, protocol_state, CLIENT_MAX_DISCOVER_TIMEOUT)
    If err_ = PLUM_PENDING Then rp_at = 1: client_run_protocol = PLUM_PENDING: Exit Function
    If err_ <> PROTOCOL_ERR_SUCCESS Then
        client_run_protocol = err_
        Exit Function
    End If

    Do While current_timestamp() < rp_end_timestamp
        rp_next_timestamp = rp_end_timestamp
        rp_i = 0
        Do While True
            If rp_i >= client.mappings_size Then Exit Do
            If client.mappings(rp_i).state = PLUM_STATE_DESTROYED Then
                rp_i = rp_i + 1
                GoTo continue_
            End If
            rp_mapping = client.mappings(rp_i)  ' client_mapping_t mapping = *cm;

            plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Mapping %d for internal port %hu is alive", rp_i, rp_mapping.internal_port

            If rp_mapping.state = PLUM_STATE_DESTROYING Then
                plum_log_write PLUM_LOG_LEVEL_INFO, "Performing unmapping for internal port %hu", rp_mapping.internal_port

resume_unmap:
                err_ = protocol_unmap(protocol, protocol_state, rp_mapping, CLIENT_MAX_MAPPING_TIMEOUT)
                If err_ = PLUM_PENDING Then rp_at = 2: client_run_protocol = PLUM_PENDING: Exit Function
                If err_ <> PROTOCOL_ERR_SUCCESS Then
                    client_run_protocol = err_
                    Exit Function
                End If

                plum_log_write PLUM_LOG_LEVEL_INFO, "Unmapped internal port %hu", rp_mapping.internal_port

                destroy_mapping client, rp_i

            ElseIf current_timestamp() >= rp_mapping.refresh_timestamp Then
                plum_log_write PLUM_LOG_LEVEL_INFO, "Performing mapping for internal port %hu", rp_mapping.internal_port

resume_map:
                err_ = protocol_map(protocol, protocol_state, rp_mapping, rp_output, CLIENT_MAX_MAPPING_TIMEOUT)
                If err_ = PLUM_PENDING Then rp_at = 3: client_run_protocol = PLUM_PENDING: Exit Function
                If err_ <> PROTOCOL_ERR_SUCCESS Then
                    client_run_protocol = err_
                    Exit Function
                End If

                If plum_log_is_enabled(PLUM_LOG_LEVEL_INFO) Then
                    If rp_output.state = PROTOCOL_MAP_STATE_SUCCESS Then
                        addr_record_to_string rp_output.external_addr, external_str, ADDR_MAX_STRING_LEN
                        plum_log_write PLUM_LOG_LEVEL_INFO, "Mapping succeeded for internal port %hu, external address is %s", _
                                       rp_mapping.internal_port, external_str
                    Else
                        plum_log_write PLUM_LOG_LEVEL_INFO, "Mapping failed for internal port %hu", rp_mapping.internal_port
                    End If
                End If

                refresh_delay = CLng(rp_output.refresh_timestamp - current_timestamp())
                plum_log_write PLUM_LOG_LEVEL_DEBUG, "Mapping will be refreshed in %dms", refresh_delay

                ' client_mapping_t *cm = client->mappings + i;
                client.mappings(rp_i).impl_record = rp_output.impl_record   ' free(cm->impl_record); cm->impl_record = ...
                client.mappings(rp_i).refresh_timestamp = rp_output.refresh_timestamp

                If client.mappings(rp_i).state <> PLUM_STATE_DESTROYING Then    ' mapping might have been destroyed
                    If rp_output.state = PROTOCOL_MAP_STATE_SUCCESS Then
                        update_mapping client, rp_i, PLUM_STATE_SUCCESS, rp_output.external_addr, True
                    Else
                        update_mapping client, rp_i, PLUM_STATE_FAILURE, zeroAddr, False
                    End If
                End If
            End If

            If rp_next_timestamp > rp_mapping.refresh_timestamp Then rp_next_timestamp = rp_mapping.refresh_timestamp

            rp_i = rp_i + 1
continue_:
        Loop

        If client.is_stopping Then Exit Do

        now = current_timestamp()
        If now < rp_next_timestamp Then
            diff = rp_next_timestamp - now
            If diff > CLIENT_RECHECK_PERIOD Then diff = CLIENT_RECHECK_PERIOD

resume_idle:
            err_ = protocol_idle(protocol, protocol_state, diff)
            If err_ = PLUM_PENDING Then rp_at = 4: client_run_protocol = PLUM_PENDING: Exit Function
            If err_ <> PROTOCOL_ERR_SUCCESS And err_ <> PROTOCOL_ERR_TIMEOUT Then
                client_run_protocol = err_
                Exit Function
            End If
        End If
    Loop

    client_run_protocol = PROTOCOL_ERR_SUCCESS
End Function

' int client_interrupt(client_t *client, bool stop)
Public Function client_interrupt(client As client_t, ByVal stop_ As Boolean) As Long
    Dim err_ As Long
    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Interrupting protocol"

    If Not client.is_started Then
        client_interrupt = PROTOCOL_ERR_SUCCESS
        Exit Function
    End If

    If stop_ Then client.is_stopping = True

    If client.protocol <> 0 Then
        err_ = protocol_interrupt(client.protocol - 1, client.protocol_state, False)
        If err_ < 0 Then
            client_interrupt = err_
            Exit Function
        End If
    End If

    client_interrupt = PROTOCOL_ERR_SUCCESS
End Function
