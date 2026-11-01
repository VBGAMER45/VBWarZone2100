Attribute VB_Name = "plum_natpmp_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/natpmp.c, natpmp.h: the NAT-PMP client, run by plum_pcp_c with the PCP client's state (pcp_impl_t).
' The packed structs are byte offsets (below). natpmp_impl_t is declared in natpmp.h but used by no code: left out.
' (VB) The waits are resumable as plum_pcp_c's (their state is kept here).
Option Explicit

Private Const AF_INET As Long = 2

' RFC 6686 specifies version 0
Public Const NATPMP_VERSION As Long = 0

' typedef enum natpmp_opcode
Private Const NATPMP_OPCODE_ANNOUNCE As Long = 0
Private Const NATPMP_OPCODE_MAP_UDP As Long = 1
Private Const NATPMP_OPCODE_MAP_TCP As Long = 2

Private Const NATPMP_OPCODE_RESPONSE_BIT As Long = &H80

' RFC 6886 3.5. Result Codes (typedef enum natpmp_result_code)
Private Const NATPMP_RESULT_SUCCESS As Long = 0
Private Const NATPMP_RESULT_UNSUPP_VERSION As Long = 1
Private Const NATPMP_RESULT_NOT_AUTHORIZED As Long = 2
Private Const NATPMP_RESULT_NETWORK_FAILURE As Long = 3
Private Const NATPMP_RESULT_NO_RESOURCES As Long = 4
Private Const NATPMP_RESULT_UNSUPP_OPCODE As Long = 5

' RFC 6886 3.2. struct natpmp_announce_request (2 bytes): version 0, opcode 1
Private Const NATPMP_ANNOUNCE_REQUEST_SIZE As Long = 2
' struct natpmp_announce_response (12 bytes): version 0, opcode 1, result 2 (uint16_t), epoch_time 4, external_addr 8
Private Const NATPMP_ANNOUNCE_RESPONSE_SIZE As Long = 12
' RFC 6886 3.3. struct natpmp_map_request (12 bytes): version 0, opcode 1, reserved 2, internal_port 4,
' suggested_external_port 6, lifetime 8
Private Const NATPMP_MAP_REQUEST_SIZE As Long = 12
' struct natpmp_map_response (16 bytes): version 0, opcode 1, result 2, epoch_time 4, internal_port 8,
' external_port 10, lifetime 12
Private Const NATPMP_MAP_RESPONSE_SIZE As Long = 16

' (VB) the resumable functions' state
Private probe_at As Long                        ' natpmp_impl_probe: 1 = waiting for the response
Private probe_gateway As addr_record_t
Private probe_buffer(0 To PCP_MAX_PAYLOAD_LENGTH - 1) As Byte
Private map_at As Long                          ' natpmp_impl_map: 1 = waiting for the response
Private map_request_opcode As Long
Private map_buffer(0 To PCP_MAX_PAYLOAD_LENGTH - 1) As Byte

' int natpmp_impl_probe(pcp_impl_t *impl, addr_record_t *found_gateway, timestamp_t end_timestamp)
Public Function natpmp_impl_probe(impl As pcp_impl_t, found_gateway As addr_record_t, ByVal end_timestamp As Double) As Long
    Dim gateway_str As String, request(0 To NATPMP_ANNOUNCE_REQUEST_SIZE - 1) As Byte, src As addr_record_t
    Dim len_ As Long, epoch_time As Long, err_ As Long, result As Long, external_str As String
    If probe_at = 1 Then probe_at = 0: GoTo resume_wait     ' (VB) resumed

    If net_get_default_gateway(AF_INET, probe_gateway) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Unable to get the default gateway address"
        natpmp_impl_probe = PROTOCOL_ERR_NETWORK_FAILED
        Exit Function
    End If

    addr_set_port probe_gateway.addr, PCP_SERVER_PORT

    If found_gateway.len_ <> 0 Then
        If Not addr_record_is_equal(probe_gateway, found_gateway, False) Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Default gateway changed"
            natpmp_impl_probe = PROTOCOL_ERR_RESET
            Exit Function
        End If
    End If

    If plum_log_is_enabled(PLUM_LOG_LEVEL_DEBUG) Then
        addr_record_to_string probe_gateway, gateway_str, ADDR_MAX_STRING_LEN
        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Probing gateway at %s", gateway_str
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Sending NAT-PMP announce request"
    request(0) = NATPMP_VERSION
    request(1) = NATPMP_OPCODE_ANNOUNCE
    If udp_sendto(impl.sock, request, NATPMP_ANNOUNCE_REQUEST_SIZE, probe_gateway) < 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP send failed, errno=%d", sockerrno()
        natpmp_impl_probe = PROTOCOL_ERR_NETWORK_FAILED
        Exit Function
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Waiting for NAT-PMP announce response..."
    Do
resume_wait:
        len_ = pcp_natpmp_impl_wait_response(impl, probe_buffer, src, end_timestamp, False)
        If len_ = PLUM_PENDING Then probe_at = 1: natpmp_impl_probe = PLUM_PENDING: Exit Function
        If len_ < 0 Then Exit Do
        If len_ < NATPMP_ANNOUNCE_RESPONSE_SIZE Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Announce response of length %d is too short", len_
            GoTo continue_
        End If
        If probe_buffer(1) <> (NATPMP_OPCODE_ANNOUNCE Or NATPMP_OPCODE_RESPONSE_BIT) Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected response opcode, ignoring"
            GoTo continue_
        End If

        epoch_time = get_be32(probe_buffer, 4)
        err_ = natpmp_impl_check_epoch_time(impl, epoch_time)
        If err_ = PROTOCOL_ERR_RESET Then
            natpmp_impl_probe = err_
            Exit Function
        End If

        result = get_be16(probe_buffer, 2)
        If result <> NATPMP_RESULT_SUCCESS Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Got NAT-PMP error response, result=%d", result
            natpmp_impl_probe = PROTOCOL_ERR_PROTOCOL_FAILED
            Exit Function
        End If

        addr_set_binary AF_INET, probe_buffer, 8, 0, impl.external_addr

        If plum_log_is_enabled(PLUM_LOG_LEVEL_DEBUG) Then
            addr_record_to_string impl.external_addr, external_str, ADDR_MAX_STRING_LEN
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Success probing NAT-PMP compatible gateway, external address is %s", external_str
        End If

        found_gateway = probe_gateway

        natpmp_impl_probe = PROTOCOL_ERR_SUCCESS
        Exit Function
continue_:
    Loop

    natpmp_impl_probe = len_                    ' len < 0
End Function

' int natpmp_impl_map(pcp_impl_t *impl, const client_mapping_t *mapping, protocol_map_output_t *output,
'                     uint32_t lifetime, const addr_record_t *gateway, timestamp_t end_timestamp)
Public Function natpmp_impl_map(impl As pcp_impl_t, mapping As client_mapping_t, output As protocol_map_output_t, _
        ByVal lifetime As Long, gateway As addr_record_t, ByVal end_timestamp As Double) As Long
    Dim zeroOutput As protocol_map_output_t, request(0 To NATPMP_MAP_REQUEST_SIZE - 1) As Byte, external_port As Long
    Dim src As addr_record_t, len_ As Long, epoch_time As Long, err_ As Long, result As Long, response_lifetime As Long
    Dim expiry_delay As Double, refresh_delay As Double
    If map_at = 1 Then map_at = 0: GoTo resume_wait         ' (VB) resumed
    output = zeroOutput

    If impl.external_addr.len_ = 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Attempted to map with NAT-PMP while external address is unknown"
        natpmp_impl_map = PROTOCOL_ERR_RESET
        Exit Function
    End If

    If mapping.protocol <> PLUM_IP_PROTOCOL_TCP And mapping.protocol <> PLUM_IP_PROTOCOL_UDP Then
        natpmp_impl_map = PROTOCOL_ERR_UNSUPP_PROTOCOL
        Exit Function
    End If

    request(0) = NATPMP_VERSION
    If mapping.protocol = PLUM_IP_PROTOCOL_UDP Then request(1) = NATPMP_OPCODE_MAP_UDP Else request(1) = NATPMP_OPCODE_MAP_TCP
    map_request_opcode = request(1)
    put_be16 request, 4, mapping.internal_port
    put_be32 request, 8, lifetime

    If mapping.external_addr.len_ > 0 Then external_port = addr_get_port(mapping.external_addr.addr) Else external_port = 0
    If external_port = 0 Then
        If mapping.suggested_addr.len_ > 0 Then external_port = addr_get_port(mapping.suggested_addr.addr) Else external_port = 0
    End If

    put_be16 request, 6, external_port

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Sending map request"
    If udp_sendto(impl.sock, request, NATPMP_MAP_REQUEST_SIZE, gateway) < 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP send failed, errno=%d", sockerrno()
        natpmp_impl_map = -1
        Exit Function
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Waiting for map response..."
    Do
resume_wait:
        len_ = pcp_natpmp_impl_wait_response(impl, map_buffer, src, end_timestamp, False)
        If len_ = PLUM_PENDING Then map_at = 1: natpmp_impl_map = PLUM_PENDING: Exit Function
        If len_ < 0 Then Exit Do
        If len_ < NATPMP_MAP_RESPONSE_SIZE Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Mapping response of length %d is too short", len_
            GoTo continue_
        End If
        If map_buffer(1) <> (map_request_opcode Or PCP_OPCODE_RESPONSE_BIT) Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected response opcode, ignoring"
            GoTo continue_
        End If

        epoch_time = get_be32(map_buffer, 4)
        err_ = natpmp_impl_check_epoch_time(impl, epoch_time)
        If err_ = PROTOCOL_ERR_RESET Then
            natpmp_impl_map = err_
            Exit Function
        End If

        result = get_be16(map_buffer, 2)
        If result <> NATPMP_RESULT_SUCCESS Then
            ' (int)response->result: the field as it is in memory (network order read little-endian)
            plum_log_write PLUM_LOG_LEVEL_WARN, "Got NAT-PMP error response, result=%d", CLng(map_buffer(2)) + CLng(map_buffer(3)) * 256
            natpmp_impl_map = PROTOCOL_ERR_PROTOCOL_FAILED
            Exit Function
        End If

        response_lifetime = get_be32(map_buffer, 12)
        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Server mapping lifetime is %us", U32(response_lifetime)

        If ULT(response_lifetime, lifetime) Then lifetime = response_lifetime

        output.state = PROTOCOL_MAP_STATE_SUCCESS

        ' RFC 6886: The client SHOULD begin trying to renew the mapping halfway to expiry time,
        ' like DHCP.
        If response_lifetime <> 0 Then
            expiry_delay = U32(lifetime) * 1000
            refresh_delay = Int(expiry_delay / 2)
            plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Renewing mapping in %us", Int(refresh_delay / 1000)
            output.refresh_timestamp = current_timestamp() + refresh_delay
        Else
            output.refresh_timestamp = 0
        End If

        external_port = get_be16(map_buffer, 10)
        output.external_addr = impl.external_addr
        addr_set_port output.external_addr.addr, external_port
        natpmp_impl_map = PROTOCOL_ERR_SUCCESS
        Exit Function
continue_:
    Loop

    natpmp_impl_map = len_                      ' len < 0
End Function

' int natpmp_impl_process_mcast_response(pcp_impl_t *impl, const char *buffer, int len)
Public Function natpmp_impl_process_mcast_response(impl As pcp_impl_t, buffer() As Byte, ByVal len_ As Long) As Long
    Dim epoch_time As Long, err_ As Long, result As Long, external_str As String
    If len_ < NATPMP_ANNOUNCE_RESPONSE_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Datagram of length %d is too short, ignoring", len_
        natpmp_impl_process_mcast_response = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If

    If buffer(1) <> (PCP_OPCODE_ANNOUNCE Or PCP_OPCODE_RESPONSE_BIT) Then
        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected opcode received multicast socket, ignoring"
        natpmp_impl_process_mcast_response = PROTOCOL_ERR_SUCCESS
        Exit Function
    End If

    epoch_time = get_be32(buffer, 4)
    err_ = natpmp_impl_check_epoch_time(impl, epoch_time)
    If err_ = PROTOCOL_ERR_RESET Then
        natpmp_impl_process_mcast_response = err_
        Exit Function
    End If

    result = get_be16(buffer, 2)
    If result = NATPMP_RESULT_SUCCESS Then
        addr_set_binary AF_INET, buffer, 8, 0, impl.external_addr

        If plum_log_is_enabled(PLUM_LOG_LEVEL_DEBUG) Then
            addr_record_to_string impl.external_addr, external_str, ADDR_MAX_STRING_LEN
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Got announce from NAT-PMP compatible gateway, external address is %s", external_str
        End If
    End If

    natpmp_impl_process_mcast_response = PROTOCOL_ERR_SUCCESS
End Function

' int natpmp_impl_check_epoch_time(pcp_impl_t *impl, uint32_t curr_server_time)
Public Function natpmp_impl_check_epoch_time(impl As pcp_impl_t, ByVal curr_server_time As Long) As Long
    Dim curr_client_time As Long, elapsed As Long, estimated_server_time As Long
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "NAT-PMP server epoch time is %u", U32(curr_server_time)
    curr_client_time = FromU32(Int(current_timestamp() / 1000))     ' seconds

    If Not impl.has_prev_server_time Then
        impl.has_prev_server_time = True
        impl.prev_client_time = curr_client_time
        impl.prev_server_time = curr_server_time
        natpmp_impl_check_epoch_time = PROTOCOL_ERR_SUCCESS
        Exit Function
    End If

    ' RFC 6886: Every packet sent by the NAT gateway includes a Seconds Since Start of Epoch
    ' (SSSoE) field. [...] Whenever a client receives any packet from the NAT gateway, either
    ' unsolicited or in response to a client request, the client computes its own conservative
    ' estimate of the expected SSSoE value by taking the SSSoE value in the last packet it received
    ' from the gateway and adding 7/8 (87.5%) of the time elapsed according to the client's local
    ' clock since that packet was received.
    elapsed = curr_client_time - impl.prev_client_time
    ' (libplum adds to curr_server_time where RFC 6886 says the last packet's)
    estimated_server_time = curr_server_time + elapsed - SHRU(elapsed, 3)

    impl.prev_client_time = curr_client_time
    impl.prev_server_time = curr_server_time

    ' If the SSSoE in the newly received packet is less than the client's conservative estimate by
    ' more than 2 seconds, then the client concludes that the NAT gateway has undergone a reboot or
    ' other loss of port mapping state, and the client MUST immediately renew all its active port
    ' mapping leases
    If ULT(curr_server_time + 2, estimated_server_time) Then
        plum_log_write PLUM_LOG_LEVEL_INFO, "NAT-PMP reset detected"
        natpmp_impl_check_epoch_time = PROTOCOL_ERR_RESET
        Exit Function
    End If

    natpmp_impl_check_epoch_time = PROTOCOL_ERR_SUCCESS
End Function
