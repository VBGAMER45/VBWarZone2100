Attribute VB_Name = "plum_pcp_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/pcp.c, pcp.h: the PCP client, and the NAT-PMP one's driver (plum_natpmp_c). The protocol's state
' (state->impl) is this module's impl. The packed structs are byte offsets into the datagram buffers (below); the
' uint32_t fields are Longs holding their bits (ULT, SHRU for the unsigned operations; Long arithmetic wraps).
' (VB) The waits are resumable (plum_socket_h's PLUM_PENDING): each function that waits keeps its state here (one
' call at a time: the client thread's) and resumes at its call that returned PLUM_PENDING.
Option Explicit

Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long

Private Const INVALID_SOCKET As Long = -1
Private Const AF_INET As Long = 2

' RFC 6687 specifies version 2
Public Const PCP_VERSION As Long = 2
Public Const PCP_CLIENT_PORT As Long = 5350
Public Const PCP_SERVER_PORT As Long = 5351

' RFC 6887: All PCP messages are sent over UDP, with a maximum UDP payload length of 1100 octets.
Public Const PCP_MAX_PAYLOAD_LENGTH As Long = 1100

Public Const PCP_MAX_ATTEMPTS As Long = 4       ' RFC 6886 recommends 9

' typedef enum pcp_interrupt_t
Public Const PCP_INTERRUPT_NONE As Long = 0
Public Const PCP_INTERRUPT_SOFT As Long = 1
Public Const PCP_INTERRUPT_HARD As Long = 2

' typedef struct pcp_impl_t
Public Type pcp_impl_t
    sock As Long
    mcast_sock As Long
    external_addr As addr_record_t
    prev_server_time As Long                    ' uint32_t
    prev_client_time As Long                    ' uint32_t
    has_prev_server_time As Boolean
    use_natpmp As Boolean
    interrupt As Long                           ' atomic(pcp_interrupt_t)
End Type

' struct pcp_common_header (4 bytes): version 0, opcode 1, reserved 2, result 3
Public Const PCP_COMMON_HEADER_SIZE As Long = 4

' typedef enum pcp_opcode
Public Const PCP_OPCODE_ANNOUNCE As Long = 0
Public Const PCP_OPCODE_MAP As Long = 1
Public Const PCP_OPCODE_PEER As Long = 2

Public Const PCP_OPCODE_RESPONSE_BIT As Long = &H80

' RFC 6887 7.1. struct pcp_request_header (24 bytes): version 0, opcode 1, reserved 2 (uint16_t), lifetime 4
' (uint32_t), client_address 8 (16 bytes)
Private Const PCP_REQUEST_HEADER_SIZE As Long = 24
' RFC 6887 7.2. struct pcp_response_header (24 bytes): version 0, opcode 1, reserved 2, result 3, lifetime 4,
' epoch_time 8, reserved96 12
Private Const PCP_RESPONSE_HEADER_SIZE As Long = 24

' RFC 6887 7.4. Result Codes (typedef enum pcp_result_code)
Public Const PCP_RESULT_SUCCESS As Long = 0
Public Const PCP_RESULT_UNSUPP_VERSION As Long = 1
Public Const PCP_RESULT_NOT_AUTHORIZED As Long = 2
Public Const PCP_RESULT_MALFORMED_REQUEST As Long = 3
Public Const PCP_RESULT_UNSUPP_OPCODE As Long = 4
Public Const PCP_RESULT_UNSUPP_OPTION As Long = 5
Public Const PCP_RESULT_MALFORMED_OPTION As Long = 6
Public Const PCP_RESULT_NETWORK_FAILURE As Long = 7
Public Const PCP_RESULT_NO_RESOURCES As Long = 8
Public Const PCP_RESULT_UNSUPP_PROTOCOL As Long = 9
Public Const PCP_RESULT_USER_EX_QUOTA As Long = 10
Public Const PCP_RESULT_CANNOT_PROVIDE_EXTERNAL As Long = 11
Public Const PCP_RESULT_ADDRESS_MISMATCH As Long = 12
Public Const PCP_RESULT_EXCESSIVE_REMOTE_PEERS As Long = 13

' RFC 6887 11.1. struct pcp_map_request / struct pcp_map_response (36 bytes, after the header): nonce 0 (12 bytes),
' protocol 12, reserved 13 (3 bytes), internal_port 16, suggested_external_port / external_port 18,
' suggested_external_addr / external_addr 20 (16 bytes)
Public Const PCP_MAP_NONCE_SIZE As Long = 12
Private Const PCP_MAP_SIZE As Long = 36

' typedef enum pcp_protocol (the IANA protocol registry)
Private Const PCP_PROTOCOL_ALL As Long = 0
Private Const PCP_PROTOCOL_TCP As Long = 6
Private Const PCP_PROTOCOL_UDP As Long = 17

Private impl As pcp_impl_t                      ' state->impl

' (VB) the resumable functions' state
Private disc_at As Long                         ' pcp_discover: 1 = in natpmp_impl_probe, 2 = in pcp_impl_probe
Private disc_end_timestamp As Double
Private disc_probe_count As Long
Private disc_probe_duration As Double
Private disc_probe_end_timestamp As Double
Private map_at As Long                          ' pcp_map: 1 = in natpmp_impl_map, 2 = in pcp_impl_map
Private map_end_timestamp As Double
Private map_count As Long
Private map_duration As Double
Private map_map_end_timestamp As Double
Private unmap_at As Long                        ' pcp_unmap: as pcp_map
Private unmap_end_timestamp As Double
Private unmap_count As Long
Private unmap_duration As Double
Private unmap_map_end_timestamp As Double
Private unmap_output As protocol_map_output_t
Private idle_at As Long                         ' pcp_idle: 1 = in pcp_natpmp_impl_wait_response
Private idle_end_timestamp As Double
Private idle_buffer(0 To PCP_MAX_PAYLOAD_LENGTH - 1) As Byte
Private probe_at As Long                        ' pcp_impl_probe: 1 = waiting for the response
Private probe_gateway As addr_record_t
Private probe_buffer(0 To PCP_MAX_PAYLOAD_LENGTH - 1) As Byte
Private implmap_at As Long                      ' pcp_impl_map: 1 = waiting for the response
Private implmap_buffer(0 To PCP_MAX_PAYLOAD_LENGTH - 1) As Byte
Private implmap_nonce(0 To PCP_MAP_NONCE_SIZE - 1) As Byte
Private wait_at As Long                         ' pcp_natpmp_impl_wait_response: 1 = in its loop
Private wait_end_timestamp As Double
Private wait_is_reset As Boolean

Private Function closesocket(ByVal s As Long) As Long
    closesocket = ws_closesocket(s)
    plum_sock_called
End Function

' (VB) a uint32_t's value modulo m (m > 0), as a Double
Private Function umod(ByVal x As Long, ByVal m As Double) As Double
    Dim v As Double
    v = U32(x)
    umod = v - Int(v / m) * m
End Function

' int pcp_init(protocol_state_t *state)
Public Function pcp_init(state As protocol_state_t) As Long
    Dim zeroState As protocol_state_t, zeroImpl As pcp_impl_t, udp_config As udp_socket_config_t, mcast_udp_config As udp_socket_config_t
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Initializing PCP/NAT-PMP state"
    state = zeroState

    state.impl = True                           ' malloc(sizeof(pcp_impl_t))

    impl = zeroImpl
    impl.sock = INVALID_SOCKET
    impl.mcast_sock = INVALID_SOCKET
    impl.has_prev_server_time = False
    impl.use_natpmp = False
    impl.interrupt = PCP_INTERRUPT_NONE

    udp_config.family = AF_INET
    udp_config.port = 0
    impl.sock = udp_create_socket(udp_config)
    If impl.sock = INVALID_SOCKET Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP socket creation failed"
        GoTo error_
    End If

    mcast_udp_config.family = AF_INET
    mcast_udp_config.port = PCP_CLIENT_PORT
    mcast_udp_config.multicast_group = "224.0.0.1"
    mcast_udp_config.enable_reuseaddr = True
    impl.mcast_sock = udp_create_socket(mcast_udp_config)
    If impl.mcast_sock = INVALID_SOCKET Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Multicast UDP socket creation on port %d failed", PCP_CLIENT_PORT
        GoTo error_
    End If

    pcp_init = PROTOCOL_ERR_SUCCESS
    Exit Function

error_:
    If impl.sock <> INVALID_SOCKET Then closesocket impl.sock

    If impl.mcast_sock <> INVALID_SOCKET Then closesocket impl.mcast_sock

    state.impl = False                          ' free(state->impl); state->impl = NULL;
    pcp_init = PROTOCOL_ERR_INSUFF_RESOURCES
End Function

' int pcp_cleanup(protocol_state_t *state)
Public Function pcp_cleanup(state As protocol_state_t) As Long
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Cleaning up PCP/NAT-PMP state"

    closesocket impl.sock
    closesocket impl.mcast_sock

    state.impl = False
    pcp_cleanup = PROTOCOL_ERR_SUCCESS
End Function

' int pcp_discover(protocol_state_t *state, timediff_t duration)
Public Function pcp_discover(state As protocol_state_t, ByVal duration As Double) As Long
    Dim err_ As Long, gateway_str As String
    Select Case disc_at                         ' (VB) resumed
    Case 1: disc_at = 0: GoTo resume_natpmp
    Case 2: disc_at = 0: GoTo resume_pcp
    End Select
    disc_end_timestamp = current_timestamp() + duration

    ' RFC 6886: To determine the external IPv4 address, or to request a port mapping, a NAT-PMP
    ' client sends its request packet to port 5351 of its configured gateway address, and waits 250
    ' ms for a response.  If no NAT-PMP response is received from the gateway after 250 ms, the
    ' client retransmits its request and waits 500 ms.  The client SHOULD repeat this process with
    ' the interval between attempts doubling each time.
    disc_probe_count = 0
    disc_probe_duration = 250
    Do
        disc_probe_end_timestamp = current_timestamp() + disc_probe_duration
        If disc_probe_end_timestamp > disc_end_timestamp Then disc_probe_end_timestamp = disc_end_timestamp

        If impl.use_natpmp Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Probing NAT-PMP..."
resume_natpmp:
            err_ = natpmp_impl_probe(impl, state.gateway, disc_probe_end_timestamp)
            If err_ = PLUM_PENDING Then disc_at = 1: pcp_discover = PLUM_PENDING: Exit Function
        Else
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Probing PCP..."
resume_pcp:
            err_ = pcp_impl_probe(impl, state.gateway, disc_probe_end_timestamp)
            If err_ = PLUM_PENDING Then disc_at = 2: pcp_discover = PLUM_PENDING: Exit Function
            If err_ = PROTOCOL_ERR_NEED_DOWNGRADE Then
                plum_log_write PLUM_LOG_LEVEL_DEBUG, "Falling back to NAT-PMP"
                impl.use_natpmp = True
                GoTo continue_
            End If
        End If
        If err_ = PROTOCOL_ERR_SUCCESS Then
            If plum_log_is_enabled(PLUM_LOG_LEVEL_INFO) Then
                addr_record_to_string state.gateway, gateway_str, ADDR_MAX_STRING_LEN
                If impl.use_natpmp Then
                    plum_log_write PLUM_LOG_LEVEL_INFO, "Success probing NAT-PMP, gateway address is %s", gateway_str
                Else
                    plum_log_write PLUM_LOG_LEVEL_INFO, "Success probing PCP, gateway address is %s", gateway_str
                End If
            End If
            pcp_discover = PROTOCOL_ERR_SUCCESS
            Exit Function
        End If

        If err_ <> PROTOCOL_ERR_TIMEOUT Then
            pcp_discover = err_
            Exit Function
        End If

        disc_probe_duration = disc_probe_duration * 2
continue_:
        disc_probe_count = disc_probe_count + 1
    Loop While disc_probe_count < PCP_MAX_ATTEMPTS And current_timestamp() < disc_end_timestamp

    pcp_discover = PROTOCOL_ERR_TIMEOUT
End Function

' int pcp_map(protocol_state_t *state, const client_mapping_t *mapping, protocol_map_output_t *output,
'             timediff_t duration)
Public Function pcp_map(state As protocol_state_t, mapping As client_mapping_t, output As protocol_map_output_t, ByVal duration As Double) As Long
    Dim err_ As Long, lifetime As Long
    Select Case map_at                          ' (VB) resumed
    Case 1: map_at = 0: lifetime = 7200: GoTo resume_natpmp
    Case 2: map_at = 0: lifetime = 7200: GoTo resume_pcp
    End Select
    map_end_timestamp = current_timestamp() + duration

    map_count = 0
    map_duration = 250
    Do
        map_map_end_timestamp = current_timestamp() + map_duration
        If map_map_end_timestamp > map_end_timestamp Then map_map_end_timestamp = map_end_timestamp

        If impl.use_natpmp Then
            ' RFC 6886: The RECOMMENDED Port Mapping Lifetime is 7200 seconds (two hours).
            lifetime = 7200 ' seconds

            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Mapping with NAT-PMP..."
resume_natpmp:
            err_ = natpmp_impl_map(impl, mapping, output, lifetime, state.gateway, map_map_end_timestamp)
            If err_ = PLUM_PENDING Then map_at = 1: pcp_map = PLUM_PENDING: Exit Function
        Else
            ' RFC 6887: The PCP client requests a certain lifetime, and the PCP server responds
            ' with the assigned lifetime. The PCP server MAY grant a lifetime smaller or larger
            ' than the requested lifetime. The PCP server SHOULD be configurable for permitted
            ' minimum and maximum lifetime, and the minimum value SHOULD be 120 seconds. The
            ' maximum value SHOULD be the remaining lifetime of the IP address assigned to the PCP
            ' client if that information is available (e.g., from the DHCP server), or half the
            ' lifetime of IP address assignments on that network if the remaining lifetime is not
            ' available, or 24 hours. RFC 6886: The RECOMMENDED Port Mapping Lifetime is 7200
            ' seconds (two hours).
            lifetime = 7200 ' seconds

            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Mapping with PCP..."
resume_pcp:
            err_ = pcp_impl_map(impl, mapping, output, lifetime, state.gateway, map_map_end_timestamp)
            If err_ = PLUM_PENDING Then map_at = 2: pcp_map = PLUM_PENDING: Exit Function
        End If

        If err_ = PROTOCOL_ERR_SUCCESS Then
            If impl.use_natpmp Then
                plum_log_write PLUM_LOG_LEVEL_DEBUG, "Success mapping with NAT-PMP"
            Else
                plum_log_write PLUM_LOG_LEVEL_DEBUG, "Success mapping with PCP"
            End If

            pcp_map = PROTOCOL_ERR_SUCCESS
            Exit Function
        End If

        If err_ <> PROTOCOL_ERR_TIMEOUT Then
            pcp_map = err_
            Exit Function
        End If

        map_duration = map_duration * 2
        map_count = map_count + 1
    Loop While map_count < PCP_MAX_ATTEMPTS And current_timestamp() < map_end_timestamp

    pcp_map = PROTOCOL_ERR_TIMEOUT
End Function

' int pcp_unmap(protocol_state_t *state, const client_mapping_t *mapping, timediff_t duration)
Public Function pcp_unmap(state As protocol_state_t, mapping As client_mapping_t, ByVal duration As Double) As Long
    Dim err_ As Long, lifetime As Long
    lifetime = 0                                ' We want to unmap
    ' protocol_map_output_t output; dummy: unmap_output
    Select Case unmap_at                        ' (VB) resumed
    Case 1: unmap_at = 0: GoTo resume_natpmp
    Case 2: unmap_at = 0: GoTo resume_pcp
    End Select
    unmap_end_timestamp = current_timestamp() + duration

    unmap_count = 0
    unmap_duration = 250
    Do
        unmap_map_end_timestamp = current_timestamp() + unmap_duration
        If unmap_map_end_timestamp > unmap_end_timestamp Then unmap_map_end_timestamp = unmap_end_timestamp

        If impl.use_natpmp Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unmapping with NAT-PMP..."
resume_natpmp:
            err_ = natpmp_impl_map(impl, mapping, unmap_output, lifetime, state.gateway, unmap_map_end_timestamp)
            If err_ = PLUM_PENDING Then unmap_at = 1: pcp_unmap = PLUM_PENDING: Exit Function
        Else
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unmapping with PCP..."
resume_pcp:
            err_ = pcp_impl_map(impl, mapping, unmap_output, lifetime, state.gateway, unmap_map_end_timestamp)
            If err_ = PLUM_PENDING Then unmap_at = 2: pcp_unmap = PLUM_PENDING: Exit Function
        End If

        If err_ = PROTOCOL_ERR_SUCCESS Then
            If impl.use_natpmp Then
                plum_log_write PLUM_LOG_LEVEL_DEBUG, "Success unmapping with NAT-PMP"
            Else
                plum_log_write PLUM_LOG_LEVEL_DEBUG, "Success unmapping with PCP"
            End If

            unmap_output.impl_record = Empty    ' free(output.impl_record);
            pcp_unmap = PROTOCOL_ERR_SUCCESS
            Exit Function
        End If

        If err_ <> PROTOCOL_ERR_TIMEOUT Then
            pcp_unmap = err_
            Exit Function
        End If

        unmap_duration = unmap_duration * 2
        unmap_count = unmap_count + 1
    Loop While unmap_count < PCP_MAX_ATTEMPTS And current_timestamp() < unmap_end_timestamp

    pcp_unmap = PROTOCOL_ERR_TIMEOUT
End Function

' int pcp_idle(protocol_state_t *state, timediff_t duration)
Public Function pcp_idle(state As protocol_state_t, ByVal duration As Double) As Long
    Dim src As addr_record_t, len_ As Long
    If idle_at = 1 Then
        idle_at = 0                             ' (VB) resumed
    Else
        idle_end_timestamp = current_timestamp() + duration
    End If

    Do
        len_ = pcp_natpmp_impl_wait_response(impl, idle_buffer, src, idle_end_timestamp, True)
        If len_ = PLUM_PENDING Then idle_at = 1: pcp_idle = PLUM_PENDING: Exit Function
        If len_ < 0 Then Exit Do
        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected datagram, ignoring"
    Loop

    pcp_idle = len_                             ' len < 0
End Function

' int pcp_interrupt(protocol_state_t *state, bool hard)
Public Function pcp_interrupt(state As protocol_state_t, ByVal hard As Boolean) As Long
    Dim dummy(0 To 0) As Byte
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Interrupting PCP/NAT-PMP operation"
    If hard Then impl.interrupt = PCP_INTERRUPT_HARD Else impl.interrupt = PCP_INTERRUPT_SOFT

    dummy(0) = 0                                ' Some C libraries might error out on NULL pointers
    If udp_sendto_self(impl.sock, dummy, 0) < 0 Then
        If sockerrno() <> SEAGAIN And sockerrno() <> SEWOULDBLOCK Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to interrupt PCP/NAT-PMP operation by triggering socket, errno=%d", sockerrno()
            pcp_interrupt = PROTOCOL_ERR_UNKNOWN
            Exit Function
        End If
    End If

    pcp_interrupt = PROTOCOL_ERR_SUCCESS
End Function

' static int write_pcp_header(struct pcp_request_header *header, uint8_t opcode, uint32_t lifetime)
Private Function write_pcp_header(header() As Byte, ByVal opcode As Long, ByVal lifetime As Long) As Long
    Dim i As Long, local_ As addr_record_t
    For i = 0 To PCP_REQUEST_HEADER_SIZE - 1    ' memset(header, 0, sizeof(*header));
        header(i) = 0
    Next
    header(0) = PCP_VERSION
    header(1) = opcode
    ' header->lifetime = lifetime; (libplum writes it in host order: little-endian)
    header(4) = lifetime And &HFF&
    header(5) = (lifetime And &HFF00&) \ &H100&
    header(6) = (lifetime And &HFF0000) \ &H10000
    header(7) = ((lifetime And &HFF000000) \ &H1000000) And &HFF&

    If net_get_default_interface(AF_INET, local_) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Unable to get default interface address"
        write_pcp_header = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If

    ' RFC 6887: When the address field holds an IPv4 address, an IPv4-mapped IPv6 address
    ' [RFC4291] is used.
    addr_map_inet6_v4mapped local_.addr, local_.len_
    For i = 0 To 15                             ' memcpy(header->client_address, &local_sin6->sin6_addr, 16);
        header(8 + i) = local_.addr(8 + i)
    Next
    write_pcp_header = 0
End Function

' int pcp_impl_probe(pcp_impl_t *impl, addr_record_t *found_gateway, timestamp_t end_timestamp)
Public Function pcp_impl_probe(impl As pcp_impl_t, found_gateway As addr_record_t, ByVal end_timestamp As Double) As Long
    Dim gateway_str As String, src As addr_record_t, len_ As Long, result As Long, version As Long, epoch_time As Long, err_ As Long
    If probe_at = 1 Then probe_at = 0: GoTo resume_wait     ' (VB) resumed

    If net_get_default_gateway(AF_INET, probe_gateway) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Unable to get the default gateway address"
        pcp_impl_probe = PROTOCOL_ERR_NETWORK_FAILED
        Exit Function
    End If

    addr_set_port probe_gateway.addr, PCP_SERVER_PORT

    If found_gateway.len_ <> 0 Then
        If Not addr_record_is_equal(probe_gateway, found_gateway, False) Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Default gateway changed"
            pcp_impl_probe = PROTOCOL_ERR_RESET
            Exit Function
        End If
    End If

    If plum_log_is_enabled(PLUM_LOG_LEVEL_DEBUG) Then
        addr_record_to_string probe_gateway, gateway_str, ADDR_MAX_STRING_LEN
        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Probing gateway at %s", gateway_str
    End If

    If write_pcp_header(probe_buffer, PCP_OPCODE_ANNOUNCE, 0) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Unable to write PCP header"
        pcp_impl_probe = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Sending PCP announce request"
    If udp_sendto(impl.sock, probe_buffer, PCP_REQUEST_HEADER_SIZE, probe_gateway) < 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP send failed, errno=%d", sockerrno()
        pcp_impl_probe = PROTOCOL_ERR_NETWORK_FAILED
        Exit Function
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Waiting for PCP announce response..."
    Do
resume_wait:
        len_ = pcp_natpmp_impl_wait_response(impl, probe_buffer, src, end_timestamp, False)
        If len_ = PLUM_PENDING Then probe_at = 1: pcp_impl_probe = PLUM_PENDING: Exit Function
        If len_ < 0 Then Exit Do
        If len_ < PCP_COMMON_HEADER_SIZE Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Datagram of length %d is too short, ignoring", len_
            GoTo continue_
        End If
        If probe_buffer(1) <> (PCP_OPCODE_ANNOUNCE Or PCP_OPCODE_RESPONSE_BIT) Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected response opcode, ignoring"
            GoTo continue_
        End If

        result = probe_buffer(3)
        version = probe_buffer(0)
        If result <> PCP_RESULT_SUCCESS Then
            If result = PCP_RESULT_UNSUPP_VERSION Then
                If version = 0 Then
                    ' RFC 6887 9. Version Negotiation:
                    ' If the version number in the UNSUPP_VERSION response is zero then that means
                    ' this is a NAT-PMP server [RFC6886], and a client MAY choose to communicate
                    ' with it using the older NAT-PMP protocol
                    plum_log_write PLUM_LOG_LEVEL_DEBUG, "PCP is not supported, server is NAT-PMP only"
                    pcp_impl_probe = PROTOCOL_ERR_NEED_DOWNGRADE
                    Exit Function
                Else
                    plum_log_write PLUM_LOG_LEVEL_WARN, "PCP server version is unsupported, version=%u", version
                    pcp_impl_probe = PROTOCOL_ERR_UNSUPP_VERSION
                    Exit Function
                End If
            Else
                plum_log_write PLUM_LOG_LEVEL_WARN, "Got PCP error response, result=%u", result
                pcp_impl_probe = PROTOCOL_ERR_PROTOCOL_FAILED
                Exit Function
            End If
        End If

        If len_ < PCP_RESPONSE_HEADER_SIZE Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Announce response of length %d is too short", len_
            GoTo continue_
        End If

        epoch_time = get_be32(probe_buffer, 8)  ' ntohl(header->epoch_time)
        err_ = pcp_impl_check_epoch_time(impl, epoch_time)
        If err_ = PROTOCOL_ERR_RESET Then
            pcp_impl_probe = err_
            Exit Function
        End If

        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Success probing PCP compatible gateway"
        found_gateway = probe_gateway

        pcp_impl_probe = PROTOCOL_ERR_SUCCESS
        Exit Function
continue_:
    Loop

    pcp_impl_probe = len_                       ' len < 0
End Function

' int pcp_impl_map(pcp_impl_t *impl, const client_mapping_t *mapping, protocol_map_output_t *output,
'                  uint32_t lifetime, const addr_record_t *gateway, timestamp_t end_timestamp)
Public Function pcp_impl_map(impl As pcp_impl_t, mapping As client_mapping_t, output As protocol_map_output_t, _
        ByVal lifetime As Long, gateway As addr_record_t, ByVal end_timestamp As Double) As Long
    Dim zeroOutput As protocol_map_output_t, i As Long, nonce() As Byte, protocol As Long, external As addr_record_t
    Dim src As addr_record_t, len_ As Long, epoch_time As Long, err_ As Long, response_lifetime As Long
    Dim expiry_delay As Double, refresh_delay As Double, record() As Byte
    If implmap_at = 1 Then implmap_at = 0: GoTo resume_wait     ' (VB) resumed
    output = zeroOutput

    If write_pcp_header(implmap_buffer, PCP_OPCODE_MAP, lifetime) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Unable to write PCP header"
        pcp_impl_map = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If

    If Not IsEmpty(mapping.impl_record) Then
        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Using saved nonce for PCP map request"
        nonce = mapping.impl_record
        For i = 0 To PCP_MAP_NONCE_SIZE - 1
            implmap_nonce(i) = nonce(i)
        Next
    Else
        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Generating new nonce for PCP map request"
        plum_random implmap_nonce, 0, PCP_MAP_NONCE_SIZE
    End If

    Select Case mapping.protocol
    Case PLUM_IP_PROTOCOL_TCP
        protocol = PCP_PROTOCOL_TCP
    Case PLUM_IP_PROTOCOL_UDP
        protocol = PCP_PROTOCOL_UDP
    Case Else
        protocol = PCP_PROTOCOL_ALL
    End Select

    ' struct pcp_map_request *map = (struct pcp_map_request *)(buffer + sizeof(struct pcp_request_header));
    For i = 0 To PCP_MAP_SIZE - 1               ' memset(map, 0, sizeof(*map));
        implmap_buffer(PCP_REQUEST_HEADER_SIZE + i) = 0
    Next
    For i = 0 To PCP_MAP_NONCE_SIZE - 1
        implmap_buffer(PCP_REQUEST_HEADER_SIZE + i) = implmap_nonce(i)
    Next
    implmap_buffer(PCP_REQUEST_HEADER_SIZE + 12) = protocol
    put_be16 implmap_buffer, PCP_REQUEST_HEADER_SIZE + 16, mapping.internal_port

    If mapping.external_addr.len_ > 0 Then external = mapping.external_addr Else external = mapping.suggested_addr

    ' RFC 6887: If the PCP client does not know the external address, or does not have a
    ' preference, it MUST use the address-family-specific all-zeros address
    If external.len_ = 0 Or lifetime = 0 Then addr_set AF_INET, "0.0.0.0", 0, external

    ' RFC 6887: When the address field holds an IPv4 address, an IPv4-mapped IPv6 address [RFC4291]
    ' is used.
    addr_map_inet6_v4mapped external.addr, external.len_
    implmap_buffer(PCP_REQUEST_HEADER_SIZE + 18) = external.addr(2)   ' network byte-order
    implmap_buffer(PCP_REQUEST_HEADER_SIZE + 19) = external.addr(3)
    For i = 0 To 15
        implmap_buffer(PCP_REQUEST_HEADER_SIZE + 20 + i) = external.addr(8 + i)
    Next

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Sending PCP map request"
    If udp_sendto(impl.sock, implmap_buffer, PCP_REQUEST_HEADER_SIZE + PCP_MAP_SIZE, gateway) < 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP send failed, errno=%d", sockerrno()
        pcp_impl_map = PROTOCOL_ERR_NETWORK_FAILED
        Exit Function
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Waiting for PCP map response..."
    Do
resume_wait:
        len_ = pcp_natpmp_impl_wait_response(impl, implmap_buffer, src, end_timestamp, False)
        If len_ = PLUM_PENDING Then implmap_at = 1: pcp_impl_map = PLUM_PENDING: Exit Function
        If len_ < 0 Then Exit Do
        If len_ < PCP_RESPONSE_HEADER_SIZE Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Mapping response of length %d is too short", len_
            GoTo continue_
        End If
        If implmap_buffer(1) <> (PCP_OPCODE_MAP Or PCP_OPCODE_RESPONSE_BIT) Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected response opcode, ignoring"
            GoTo continue_
        End If

        If len_ < PCP_RESPONSE_HEADER_SIZE + PCP_MAP_SIZE Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Mapping success response of length=%d is too short", len_
            GoTo continue_
        End If

        ' const struct pcp_map_response *map = (buffer + sizeof(struct pcp_response_header));
        For i = 0 To PCP_MAP_NONCE_SIZE - 1     ' memcmp(map->nonce, nonce, PCP_MAP_NONCE_SIZE) != 0
            If implmap_buffer(PCP_RESPONSE_HEADER_SIZE + i) <> implmap_nonce(i) Then Exit For
        Next
        If i < PCP_MAP_NONCE_SIZE Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected nonce in map response, ignoring"
            GoTo continue_
        End If

        epoch_time = get_be32(implmap_buffer, 8)
        err_ = pcp_impl_check_epoch_time(impl, epoch_time)
        If err_ = PROTOCOL_ERR_RESET Then
            pcp_impl_map = err_
            Exit Function
        End If

        If implmap_buffer(3) <> PCP_RESULT_SUCCESS Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Got PCP error response, result=%d", implmap_buffer(3)
            pcp_impl_map = PROTOCOL_ERR_PROTOCOL_FAILED
            Exit Function
        End If

        response_lifetime = get_be32(implmap_buffer, 4)
        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Server mapping lifetime is %us", U32(response_lifetime)

        If ULT(response_lifetime, lifetime) Then lifetime = response_lifetime

        output.state = PROTOCOL_MAP_STATE_SUCCESS

        ' RFC 6887: The PCP client SHOULD renew the mapping before its expiry time; otherwise, it
        ' will be removed by the PCP server. To reduce the risk of inadvertent synchronization of
        ' renewal requests, a random jitter component should be included.
        If response_lifetime <> 0 Then
            expiry_delay = U32(lifetime) * 1000
            ' (VB) expiry_delay / 4 is 0 when the requested lifetime is (an unmapping answered with a lifetime):
            ' C divides by zero there (undefined); the port takes no jitter
            If Int(expiry_delay / 4) > 0 Then
                refresh_delay = Int(expiry_delay / 2) + umod(plum_rand32(), Int(expiry_delay / 4))
            Else
                refresh_delay = Int(expiry_delay / 2)
            End If
            plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Renewing mapping in %us", Int(refresh_delay / 1000)
            output.refresh_timestamp = current_timestamp() + refresh_delay
        Else
            output.refresh_timestamp = 0
        End If

        ' struct sockaddr_in6 *external_sin6 = (struct sockaddr_in6 *)&impl->external_addr.addr;
        For i = 0 To 27                         ' memset(external_sin6, 0, sizeof(*external_sin6));
            impl.external_addr.addr(i) = 0
        Next
        impl.external_addr.len_ = 28
        impl.external_addr.addr(0) = 23         ' external_sin6->sin6_family = AF_INET6;
        impl.external_addr.addr(2) = implmap_buffer(PCP_RESPONSE_HEADER_SIZE + 18)    ' sin6_port = map->external_port
        impl.external_addr.addr(3) = implmap_buffer(PCP_RESPONSE_HEADER_SIZE + 19)
        For i = 0 To 15
            impl.external_addr.addr(8 + i) = implmap_buffer(PCP_RESPONSE_HEADER_SIZE + 20 + i)
        Next
        addr_unmap_inet6_v4mapped impl.external_addr.addr, impl.external_addr.len_
        output.external_addr = impl.external_addr

        If lifetime <> 0 Then
            ReDim record(0 To PCP_MAP_NONCE_SIZE - 1)   ' output->impl_record = malloc(PCP_MAP_NONCE_SIZE);
            For i = 0 To PCP_MAP_NONCE_SIZE - 1
                record(i) = implmap_nonce(i)
            Next
            output.impl_record = record
        End If

        pcp_impl_map = PROTOCOL_ERR_SUCCESS
        Exit Function
continue_:
    Loop

    pcp_impl_map = len_                         ' len < 0
End Function

' int pcp_impl_process_mcast_response(pcp_impl_t *impl, const char *buffer, int len)
Public Function pcp_impl_process_mcast_response(impl As pcp_impl_t, buffer() As Byte, ByVal len_ As Long) As Long
    Dim epoch_time As Long, err_ As Long
    If len_ < PCP_RESPONSE_HEADER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Datagram of length %d is too short, ignoring", len_
        pcp_impl_process_mcast_response = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If

    If buffer(1) <> (PCP_OPCODE_ANNOUNCE Or PCP_OPCODE_RESPONSE_BIT) Then
        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected opcode received multicast socket, ignoring"
        pcp_impl_process_mcast_response = PROTOCOL_ERR_SUCCESS
        Exit Function
    End If

    epoch_time = get_be32(buffer, 8)
    err_ = pcp_impl_check_epoch_time(impl, epoch_time)
    If err_ = PROTOCOL_ERR_RESET Then
        pcp_impl_process_mcast_response = err_
        Exit Function
    End If

    pcp_impl_process_mcast_response = PROTOCOL_ERR_SUCCESS
End Function

' int pcp_impl_check_epoch_time(pcp_impl_t *impl, uint32_t curr_server_time)
Public Function pcp_impl_check_epoch_time(impl As pcp_impl_t, ByVal curr_server_time As Long) As Long
    Dim curr_client_time As Long, is_valid As Boolean, client_delta As Long, server_delta As Long
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "PCP server epoch time is %u", U32(curr_server_time)
    curr_client_time = FromU32(Int(current_timestamp() / 1000))     ' seconds

    ' RFC 6887: Whenever a client receives a PCP response, the client validates the received Epoch
    ' Time value according to the procedure below, using integer arithmetic:

    ' ... If this is the first PCP response the client has received from this PCP server, the Epoch
    ' Time value is treated as necessarily valid
    If Not impl.has_prev_server_time Then
        is_valid = True
    ' ... If the current PCP server Epoch time (curr_server_time) is less than the previously
    ' received PCP server Epoch time (prev_server_time) by more than one second, then the
    ' client treats the Epoch time as obviously invalid (time should not go backwards).
    ElseIf ULT(curr_server_time, impl.prev_server_time - 1) Then
        is_valid = False
    ' If the server Epoch time passes this check, then further validation checks are performed:
    Else
        ' ... The client computes the difference between its current local time (curr_client_time)
        ' and the time the previous PCP response was received from this PCP server
        ' (prev_client_time)
        client_delta = curr_client_time - impl.prev_client_time

        ' ... The client computes the difference between the current PCP server Epoch time
        ' (curr_server_time) and the previously received Epoch time (prev_server_time)
        server_delta = curr_server_time - impl.prev_server_time

        ' ... If client_delta+2 < server_delta - server_delta/16 or server_delta+2 < client_delta -
        ' client_delta/16, then the client treats the Epoch Time value as invalid, else the client
        ' treats the Epoch Time value as valid.
        is_valid = Not (ULT(client_delta + 2, server_delta - SHRU(server_delta, 4)) Or _
                        ULT(server_delta + 2, client_delta - SHRU(client_delta, 4)))
    End If

    ' ... The client records the current time values for use in its next comparison
    impl.prev_client_time = curr_client_time
    impl.prev_server_time = curr_server_time

    If Not is_valid Then
        plum_log_write PLUM_LOG_LEVEL_INFO, "PCP reset detected"
        pcp_impl_check_epoch_time = PROTOCOL_ERR_RESET
        Exit Function
    End If

    pcp_impl_check_epoch_time = PROTOCOL_ERR_SUCCESS
End Function

' static int process_mcast_response(pcp_impl_t *impl, const char *buffer, int len)
Private Function process_mcast_response(impl As pcp_impl_t, buffer() As Byte, ByVal len_ As Long) As Long
    Dim opcode As Long, version As Long
    If len_ < PCP_COMMON_HEADER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Datagram of length %d is too short, ignoring", len_
        process_mcast_response = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If
    opcode = buffer(1)
    If opcode <> (PCP_OPCODE_ANNOUNCE Or PCP_OPCODE_RESPONSE_BIT) Then
        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected response opcode, ignoring"
        process_mcast_response = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If

    version = buffer(0)
    If Not impl.use_natpmp And version = PCP_VERSION Then
        process_mcast_response = pcp_impl_process_mcast_response(impl, buffer, len_)
        Exit Function
    End If
    If impl.use_natpmp And version = NATPMP_VERSION Then
        process_mcast_response = natpmp_impl_process_mcast_response(impl, buffer, len_)
    Else
        process_mcast_response = PROTOCOL_ERR_UNSUPP_VERSION
    End If
End Function

' int pcp_natpmp_impl_wait_response(pcp_impl_t *impl, char *buffer, addr_record_t *src, timestamp_t end_timestamp,
'                                   bool interruptible)
' (VB) resumable: a resumed call goes on with its own end_timestamp (a reset moves it) and is_reset
Public Function pcp_natpmp_impl_wait_response(impl As pcp_impl_t, buffer() As Byte, src As addr_record_t, _
        ByVal end_timestamp As Double, ByVal interruptible As Boolean) As Long
    Dim timediff As Double, interrupt As Long, pfd(0 To 1) As pollfd, ret As Long, i As Long, len_ As Long
    Dim err_ As Long, wait_duration As Double
    If wait_at = 1 Then
        wait_at = 0                             ' (VB) resumed
    Else
        wait_end_timestamp = end_timestamp
        wait_is_reset = False
    End If
    Do
        timediff = wait_end_timestamp - current_timestamp()
        If timediff <= 0 Then Exit Do

        interrupt = impl.interrupt
        If interrupt = PCP_INTERRUPT_HARD Or (interrupt = PCP_INTERRUPT_SOFT And interruptible) Then
            plum_log_write PLUM_LOG_LEVEL_VERBOSE, "PCP/NAT-PMP interrupted"
            impl.interrupt = PCP_INTERRUPT_NONE
            pcp_natpmp_impl_wait_response = PROTOCOL_ERR_INTERRUPTED
            Exit Function
        End If

        pfd(0).fd = impl.mcast_sock
        pfd(0).events = POLLIN
        pfd(1).fd = impl.sock
        pfd(1).events = POLLIN

        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Entering poll for %d ms", CLng(timediff)
        ret = poll(pfd, 2, 0)                   ' poll(pfd, 2, (int)timediff)
        If ret = 0 Then wait_at = 1: pcp_natpmp_impl_wait_response = PLUM_PENDING: Exit Function  ' (VB)
        If ret < 0 Then
            If sockerrno() = SEINTR Or sockerrno() = SEAGAIN Then
                plum_log_write PLUM_LOG_LEVEL_VERBOSE, "poll interrupted"
                GoTo continue_
            Else
                plum_log_write PLUM_LOG_LEVEL_ERROR, "poll failed, errno=%d", sockerrno()
                pcp_natpmp_impl_wait_response = PROTOCOL_ERR_UNKNOWN
                Exit Function
            End If
        End If

        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Exiting poll"

        ' if (ret == 0) break; timeout: (VB) a pass that finds nothing ends above

        For i = 0 To 1
            If (pfd(i).revents And POLLNVAL) <> 0 Or (pfd(i).revents And POLLERR) <> 0 Then
                plum_log_write PLUM_LOG_LEVEL_ERROR, "Error when polling socket"
                pcp_natpmp_impl_wait_response = PROTOCOL_ERR_UNKNOWN
                Exit Function
            End If

            If (pfd(i).revents And POLLIN) <> 0 Then
                len_ = udp_recvfrom(pfd(i).fd, buffer, PCP_MAX_PAYLOAD_LENGTH, src)
                If len_ < 0 Then
                    If sockerrno() = SEAGAIN Or sockerrno() = SEWOULDBLOCK Then GoTo next_i

                    plum_log_write PLUM_LOG_LEVEL_WARN, "UDP recvfrom failed, errno=%d", sockerrno()
                    If sockerrno() = SECONNRESET Then  ' ICMP Port Unreachable
                        pcp_natpmp_impl_wait_response = PROTOCOL_ERR_UNSUPP_PROTOCOL
                    Else
                        pcp_natpmp_impl_wait_response = PROTOCOL_ERR_NETWORK_FAILED
                    End If
                    Exit Function
                End If

                If i = 0 Then                   ' mcast_sock
                    err_ = process_mcast_response(impl, buffer, len_)  ' handles PCP and NAT-PMP
                    If err_ = PROTOCOL_ERR_RESET Then
                        ' RFC 6887: [The client] MUST wait a random amount of time between 0 and 5
                        ' seconds
                        wait_duration = umod(plum_rand32(), 5000)
                        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "PCP/NAT-PMP reset, waiting %ums...", wait_duration
                        wait_end_timestamp = current_timestamp() + wait_duration
                        wait_is_reset = True
                    End If
                Else                            ' sock
                    If len_ > 0 And Not wait_is_reset Then  ' 0-length datagrams are used to interrupt, ignore them
                        pcp_natpmp_impl_wait_response = len_
                        Exit Function
                    End If
                End If
            End If
next_i:
        Next
continue_:
    Loop

    If wait_is_reset Then
        pcp_natpmp_impl_wait_response = PROTOCOL_ERR_RESET
        Exit Function
    End If

    pcp_natpmp_impl_wait_response = PROTOCOL_ERR_TIMEOUT
End Function
