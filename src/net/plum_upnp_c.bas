Attribute VB_Name = "plum_upnp_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/upnp.c, upnp.h: the UPnP-IGD client (SSDP search, the device's description, its SOAP calls). The
' protocol's state (state->impl) is this module's impl. SSDP's datagrams and HTTP's texts are ANSI text in Strings.
' (VB) TEST ONLY: upnp_impl_probe sends its SSDP search where plum_net_c's net_test_override_ssdp says when the
' test override is set (unicast to the simulated router on 127.0.0.1; never the multicast group then).
' (VB) The waits are resumable (plum_socket_h's PLUM_PENDING): each function that waits keeps its state here (one
' call at a time: the client thread's) and resumes at its call that returned PLUM_PENDING.
Option Explicit

Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long

Private Const INVALID_SOCKET As Long = -1
Private Const AF_INET As Long = 2

Private Const UPNP_SSDP_ADDRESS As String = "239.255.255.250"
Private Const UPNP_SSDP_PORT As Long = 1900

Private Const UPNP_SSDP_MAX_ATTEMPTS As Long = 4
Private Const UPNP_MAP_MAX_ATTEMPTS As Long = 10
Private Const UPNP_QUERY_TIMEOUT As Long = 5000 ' msecs

Private Const UPNP_BUFFER_SIZE As Long = 2048

' typedef enum upnp_interrupt_t
Private Const UPNP_INTERRUPT_NONE As Long = 0
Private Const UPNP_INTERRUPT_SOFT As Long = 1
Private Const UPNP_INTERRUPT_HARD As Long = 2

' typedef struct upnp_impl_t
Private Type upnp_impl_t
    sock As Long
    external_addr_str As String                 ' char external_addr_str[ADDR_MAX_STRING_LEN]
    location_url As String                      ' char *location_url ("" for NULL)
    service As String                           ' const char *service
    version As Long
    control_url As String                       ' char *control_url ("" for NULL)
    interrupt As Long                           ' atomic(upnp_interrupt_t)
End Type

Private impl As upnp_impl_t                     ' state->impl

' (VB) the resumable functions' state
Private disc_at As Long                         ' upnp_discover: 1 = in upnp_impl_probe
Private disc_end_timestamp As Double
Private disc_probe_count As Long
Private disc_probe_duration As Double
Private disc_probe_end_timestamp As Double
Private disc_query_end_timestamp As Double
Private map_at As Long                          ' upnp_map: 1 = in upnp_impl_query_external_addr, 2 = in upnp_impl_map
Private map_end_timestamp As Double
Private map_duration As Double
Private map_external_port As Long
Private map_lifetime As Long
Private map_query_count As Long
Private map_query_end_timestamp As Double
Private unmap_at As Long                        ' upnp_unmap: 1 = in upnp_impl_unmap
Private unmap_end_timestamp As Double
Private unmap_external_port As Long
Private idle_at As Long                         ' upnp_idle: 1 = in upnp_impl_wait_response
Private idle_end_timestamp As Double
Private idle_buffer(0 To UPNP_BUFFER_SIZE - 1) As Byte
Private probe_at As Long                        ' upnp_impl_probe: 1 = waiting, 2 = in query_control_url, 3 = in query_external_addr
Private probe_buffer(0 To UPNP_BUFFER_SIZE - 1) As Byte
Private query_at As Long                        ' upnp_impl_query_control_url, _query_external_addr, _map, _unmap:
                                                ' 1 = in http_perform (one of them at a time)
Private http_request As http_request_t
Private http_response As http_response_t

Private Function closesocket(ByVal s As Long) As Long
    closesocket = ws_closesocket(s)
    plum_sock_called
End Function

' static uint16_t random_port()
Private Function random_port() As Long
    Dim v As Double
    v = U32(plum_rand32())
    random_port = 1024 + CLng(v - Int(v / (49151 - 1024)) * (49151 - 1024))   ' TODO
End Function

' int upnp_init(protocol_state_t *state)
Public Function upnp_init(state As protocol_state_t) As Long
    Dim zeroState As protocol_state_t, zeroImpl As upnp_impl_t, udp_config As udp_socket_config_t
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Initializing UPnP state"
    state = zeroState

    state.impl = True                           ' malloc(sizeof(upnp_impl_t))

    impl = zeroImpl
    impl.sock = INVALID_SOCKET
    impl.external_addr_str = ""
    impl.location_url = ""
    impl.control_url = ""
    impl.interrupt = UPNP_INTERRUPT_NONE

    udp_config.family = AF_INET
    udp_config.port = 0     ' Binding explicitly to the SSDP port prevents udp_sendto_self from working on Windows
    udp_config.enable_broadcast = True
    impl.sock = udp_create_socket(udp_config)
    If impl.sock = INVALID_SOCKET Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP socket creation failed"
        GoTo error_
    End If

    upnp_init = PROTOCOL_ERR_SUCCESS
    Exit Function

error_:
    If impl.sock <> INVALID_SOCKET Then closesocket impl.sock

    state.impl = False
    upnp_init = PROTOCOL_ERR_INSUFF_RESOURCES
End Function

' int upnp_cleanup(protocol_state_t *state)
Public Function upnp_cleanup(state As protocol_state_t) As Long
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Cleaning up UPnP state"

    closesocket impl.sock
    impl.location_url = ""                      ' free(impl->location_url);
    impl.control_url = ""                       ' free(impl->control_url);

    state.impl = False
    upnp_cleanup = PROTOCOL_ERR_SUCCESS
End Function

' int upnp_discover(protocol_state_t *state, timediff_t duration)
Public Function upnp_discover(state As protocol_state_t, ByVal duration As Double) As Long
    Dim err_ As Long, gateway_str As String
    If disc_at = 1 Then disc_at = 0: GoTo resume_probe     ' (VB) resumed
    disc_end_timestamp = current_timestamp() + duration

    disc_probe_count = 0
    disc_probe_duration = 250
    Do
        disc_probe_end_timestamp = current_timestamp() + disc_probe_duration
        If disc_probe_end_timestamp > disc_end_timestamp Then disc_probe_end_timestamp = disc_end_timestamp

        disc_query_end_timestamp = current_timestamp() + UPNP_QUERY_TIMEOUT
        If disc_query_end_timestamp > disc_end_timestamp Then disc_query_end_timestamp = disc_end_timestamp

        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Probing UPnP..."
resume_probe:
        err_ = upnp_impl_probe(state.gateway, disc_probe_end_timestamp, disc_query_end_timestamp)
        If err_ = PLUM_PENDING Then disc_at = 1: upnp_discover = PLUM_PENDING: Exit Function
        If err_ = PROTOCOL_ERR_SUCCESS Then
            If plum_log_is_enabled(PLUM_LOG_LEVEL_INFO) Then
                addr_record_to_string state.gateway, gateway_str, ADDR_MAX_STRING_LEN
                plum_log_write PLUM_LOG_LEVEL_INFO, "Success probing UPnP, external address is %s", impl.external_addr_str
            End If
            upnp_discover = PROTOCOL_ERR_SUCCESS
            Exit Function
        End If

        If err_ <> PROTOCOL_ERR_TIMEOUT Then
            upnp_discover = err_
            Exit Function
        End If

        disc_probe_duration = disc_probe_duration * 2
        disc_probe_count = disc_probe_count + 1
    Loop While disc_probe_count < UPNP_SSDP_MAX_ATTEMPTS And current_timestamp() < disc_end_timestamp

    upnp_discover = PROTOCOL_ERR_TIMEOUT
End Function

' int upnp_map(protocol_state_t *state, const client_mapping_t *mapping, protocol_map_output_t *output,
'              timediff_t duration)
Public Function upnp_map(state As protocol_state_t, mapping As client_mapping_t, output As protocol_map_output_t, ByVal duration As Double) As Long
    Dim zeroOutput As protocol_map_output_t, err_ As Long, query_end_timestamp As Double
    Select Case map_at                          ' (VB) resumed
    Case 1: map_at = 0: GoTo resume_query
    Case 2: map_at = 0: GoTo resume_map
    End Select
    map_duration = duration
    map_end_timestamp = current_timestamp() + map_duration
    output = zeroOutput

    If impl.external_addr_str = "" Then
        map_query_end_timestamp = current_timestamp() + UPNP_QUERY_TIMEOUT
        If map_query_end_timestamp > map_end_timestamp Then map_query_end_timestamp = map_end_timestamp

resume_query:
        err_ = upnp_impl_query_external_addr(map_query_end_timestamp)
        If err_ = PLUM_PENDING Then map_at = 1: upnp_map = PLUM_PENDING: Exit Function
        If err_ <> PROTOCOL_ERR_SUCCESS Then
            upnp_map = err_
            Exit Function
        End If
    End If

    If mapping.external_addr.len_ > 0 Then map_external_port = addr_get_port(mapping.external_addr.addr) Else map_external_port = 0
    If map_external_port = 0 Then
        If mapping.suggested_addr.len_ > 0 Then map_external_port = addr_get_port(mapping.suggested_addr.addr) Else map_external_port = 0
        If map_external_port = 0 Then map_external_port = random_port()
    End If

    map_lifetime = 7200 ' seconds

    map_query_count = 0
    Do
        map_query_end_timestamp = current_timestamp() + UPNP_QUERY_TIMEOUT
        If map_query_end_timestamp > map_end_timestamp Then map_query_end_timestamp = map_end_timestamp

        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Mapping with UPnP..."
resume_map:
        err_ = upnp_impl_map(mapping.protocol, map_external_port, mapping.internal_port, map_lifetime, map_query_end_timestamp)
        If err_ = PLUM_PENDING Then map_at = 2: upnp_map = PLUM_PENDING: Exit Function
        If err_ = PROTOCOL_ERR_SUCCESS Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Success mapping with UPnP"
            output.state = PROTOCOL_MAP_STATE_SUCCESS
            output.refresh_timestamp = current_timestamp() + (map_lifetime \ 2) * 1000   ' halfway expiry time
            addr_set AF_INET, impl.external_addr_str, map_external_port, output.external_addr
            upnp_map = PROTOCOL_ERR_SUCCESS
            Exit Function

        ElseIf err_ > 0 Then                    ' if it's an UPnP error code
            Select Case err_
            Case 718        ' The port mapping entry specified conflicts with a object assigned
                            ' previously to another client
                map_external_port = random_port()
            Case 725        ' The NAT implementation only supports permanent lease times on port mapping
                map_duration = 0
            Case 729        ' Attempted port mapping is not allowed due to conflict with other mechanisms
                ' NAT port mapping rules can be created by other mechanisms besides UPnP IGD.
                ' Therefore, it is possible that port mappings done by independent mechanisms MAY overlap or conflict.
                map_external_port = random_port()
            Case Else
                upnp_map = PROTOCOL_ERR_PROTOCOL_FAILED
                Exit Function
            End Select

        ElseIf err_ <> PROTOCOL_ERR_TIMEOUT Then
            upnp_map = err_
            Exit Function
        End If

        map_query_count = map_query_count + 1
    Loop While map_query_count < UPNP_MAP_MAX_ATTEMPTS And current_timestamp() < map_end_timestamp

    upnp_map = PROTOCOL_ERR_TIMEOUT
End Function

' int upnp_unmap(protocol_state_t *state, const client_mapping_t *mapping, timediff_t duration)
Public Function upnp_unmap(state As protocol_state_t, mapping As client_mapping_t, ByVal duration As Double) As Long
    Dim err_ As Long
    If unmap_at = 1 Then unmap_at = 0: GoTo resume_unmap   ' (VB) resumed
    unmap_end_timestamp = current_timestamp() + duration

    If mapping.external_addr.len_ > 0 Then unmap_external_port = addr_get_port(mapping.external_addr.addr) Else unmap_external_port = 0

    If unmap_external_port = 0 Then
        upnp_unmap = PROTOCOL_ERR_SUCCESS       ' Nothing to do
        Exit Function
    End If

resume_unmap:
    err_ = upnp_impl_unmap(mapping.protocol, unmap_external_port, unmap_end_timestamp)
    If err_ = PLUM_PENDING Then unmap_at = 1
    upnp_unmap = err_
End Function

' int upnp_idle(protocol_state_t *state, timediff_t duration)
Public Function upnp_idle(state As protocol_state_t, ByVal duration As Double) As Long
    Dim src As addr_record_t, len_ As Long
    If idle_at = 1 Then
        idle_at = 0                             ' (VB) resumed
    Else
        idle_end_timestamp = current_timestamp() + duration
    End If

    Do
        len_ = upnp_impl_wait_response(idle_buffer, UPNP_BUFFER_SIZE - 1, src, idle_end_timestamp, True)
        If len_ = PLUM_PENDING Then idle_at = 1: upnp_idle = PLUM_PENDING: Exit Function
        If len_ < 0 Then Exit Do
        plum_log_write PLUM_LOG_LEVEL_DEBUG, "Unexpected datagram, ignoring"
    Loop

    upnp_idle = len_                            ' len < 0
End Function

' int upnp_interrupt(protocol_state_t *state, bool hard)
Public Function upnp_interrupt(state As protocol_state_t, ByVal hard As Boolean) As Long
    Dim dummy(0 To 0) As Byte
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Interrupting UPnP operation"
    If hard Then impl.interrupt = UPNP_INTERRUPT_HARD Else impl.interrupt = UPNP_INTERRUPT_SOFT

    dummy(0) = 0                                ' Some C libraries might error out on NULL pointers
    If udp_sendto_self(impl.sock, dummy, 0) < 0 Then
        If sockerrno() <> SEAGAIN And sockerrno() <> SEWOULDBLOCK Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to interrupt UPnP operation by triggering socket, errno=%d", sockerrno()
            upnp_interrupt = PROTOCOL_ERR_UNKNOWN
            Exit Function
        End If
    End If

    upnp_interrupt = PROTOCOL_ERR_SUCCESS
End Function

' (VB) the text of the first len bytes of a datagram buffer (buffer[len] = '\0')
Private Function buffer_text(buffer() As Byte, ByVal len_ As Long) As String
    Dim s As String, p As Long, b() As Byte, i As Long
    If len_ <= 0 Then Exit Function
    ReDim b(0 To len_ - 1)
    For i = 0 To len_ - 1
        b(i) = buffer(i)
    Next
    s = StrConv(b, vbUnicode)
    p = InStr(s, vbNullChar)
    If p > 0 Then s = Left$(s, p - 1)
    buffer_text = s
End Function

' int upnp_impl_probe(upnp_impl_t *impl, addr_record_t *found_gateway, timestamp_t end_timestamp,
'                     timestamp_t query_end_timestamp)
Private Function upnp_impl_probe(found_gateway As addr_record_t, ByVal end_timestamp As Double, ByVal query_end_timestamp As Double) As Long
    Dim broadcast As addr_record_t, broadcast_str As String, text As String, len_ As Long, src As addr_record_t
    Dim tmp As String, location_url As String, err_ As Long, request() As Byte
    Select Case probe_at                        ' (VB) resumed
    Case 1: probe_at = 0: GoTo resume_wait
    Case 2: probe_at = 0: GoTo resume_control_url
    Case 3: probe_at = 0: GoTo resume_external_addr
    End Select

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Probing gateway with SSDP"
    addr_set AF_INET, UPNP_SSDP_ADDRESS, UPNP_SSDP_PORT, broadcast
    If net_test_override_ssdp(broadcast) < 0 Then   ' (VB) TEST ONLY: the simulated router's address, or nothing
        upnp_impl_probe = PROTOCOL_ERR_NETWORK_FAILED
        Exit Function
    End If

    addr_record_to_string broadcast, broadcast_str, ADDR_MAX_STRING_LEN

    text = va("M-SEARCH * HTTP/1.1\r\nHOST: %s\r\nMAN: \""ssdp:discover\""\r\nMX: 10\r\nST: urn:schemas-upnp-org:device:InternetGatewayDevice:1\r\n\r\n", _
              broadcast_str)
    len_ = Len(text)

    If len_ <= 0 Or len_ >= UPNP_BUFFER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to write SSDP message to buffer"
        upnp_impl_probe = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If
    request = StrConv(text, vbFromUnicode)

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Broadcasting SSDP request"
    If udp_sendto(impl.sock, request, len_, broadcast) < 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "UDP broadcast failed, errno=%d", sockerrno()
        upnp_impl_probe = PROTOCOL_ERR_NETWORK_FAILED
        Exit Function
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Waiting for SSDP response..."
    Do
resume_wait:
        len_ = upnp_impl_wait_response(probe_buffer, UPNP_BUFFER_SIZE - 1, src, end_timestamp, False)
        If len_ = PLUM_PENDING Then probe_at = 1: upnp_impl_probe = PLUM_PENDING: Exit Function
        If len_ < 0 Then Exit Do
        text = buffer_text(probe_buffer, len_) ' buffer[len] = '\0'; null-terminate the string

        If header_extract(text, "NT", tmp, UPNP_BUFFER_SIZE) <= 0 Then
            If header_extract(text, "ST", tmp, UPNP_BUFFER_SIZE) <= 0 Then GoTo continue_
        End If

        If InStr(tmp, "device:InternetGatewayDevice") = 0 Then GoTo continue_

        If header_extract(text, "SERVER", tmp, UPNP_BUFFER_SIZE) > 0 Then
            plum_log_write PLUM_LOG_LEVEL_INFO, "Found UPnP-IGD device: %s", tmp
        End If

        If header_extract(text, "LOCATION", location_url, HTTP_MAX_URL_LEN) <= 0 Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Missing location for UPnP-IGD device"
            GoTo continue_
        End If

        found_gateway = src

        plum_log_write PLUM_LOG_LEVEL_DEBUG, "UPnP-IGP location URL: %s", location_url
        If impl.location_url <> "" Then
            If impl.location_url <> location_url Then
                plum_log_write PLUM_LOG_LEVEL_WARN, "UPnP-IGP location URL has changed, resetting"
                upnp_impl_probe = PROTOCOL_ERR_RESET
                Exit Function
            End If
            upnp_impl_probe = PROTOCOL_ERR_SUCCESS
            Exit Function
        End If

        impl.location_url = location_url

resume_control_url:
        err_ = upnp_impl_query_control_url(query_end_timestamp)
        If err_ = PLUM_PENDING Then probe_at = 2: upnp_impl_probe = PLUM_PENDING: Exit Function
        If err_ <> PROTOCOL_ERR_SUCCESS Then
            upnp_impl_probe = err_
            Exit Function
        End If

resume_external_addr:
        err_ = upnp_impl_query_external_addr(query_end_timestamp)
        If err_ = PLUM_PENDING Then probe_at = 3
        upnp_impl_probe = err_
        Exit Function
continue_:
    Loop

    upnp_impl_probe = len_                      ' len < 0
End Function

' (VB) the HTTP request's error as upnp.c's functions return it (ret < 0)
Private Function http_error(ByVal ret As Long) As Long
    If ret = HTTP_ERR_TIMEOUT Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Timed-out sending HTTP request to UPnP-IGD device"
        http_error = PROTOCOL_ERR_TIMEOUT
    Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to send HTTP request to UPnP-IGD device"
        http_error = PROTOCOL_ERR_NETWORK_FAILED
    End If
End Function

' int upnp_impl_query_control_url(upnp_impl_t *impl, timestamp_t end_timestamp)
Private Function upnp_impl_query_control_url(ByVal end_timestamp As Double) As Long
    Dim zeroRequest As http_request_t, zeroResponse As http_response_t, ret As Long, serviceType As String
    Dim service As String, found As Boolean, control_url As String, base_url As String, p As Long, tmp As String
    If query_at = 1 Then query_at = 0: GoTo resume_http     ' (VB) resumed
    If impl.location_url = "" Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Attempted to query UPnP control URL with unknown location URL"
        upnp_impl_query_control_url = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If

    http_request = zeroRequest
    http_request.method = HTTP_METHOD_GET
    http_request.url = impl.location_url
    http_request.headers = ""

    http_response = zeroResponse
resume_http:
    ret = http_perform(http_request, http_response, end_timestamp)
    If ret = PLUM_PENDING Then query_at = 1: upnp_impl_query_control_url = PLUM_PENDING: Exit Function
    If ret < 0 Then
        upnp_impl_query_control_url = http_error(ret)
        Exit Function
    End If

    If ret <> 200 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "HTTP request to UPnP-IGD device failed, code=%d", ret
        http_free http_response
        upnp_impl_query_control_url = PROTOCOL_ERR_NETWORK_FAILED
        Exit Function
    End If

    ' Try to find WANIPConnection:2
    serviceType = "urn:schemas-upnp-org:service:WANIPConnection:2"
    service = xml_find_matching_child(http_response.body, "service", "serviceType", serviceType, found)
    If found Then
        impl.service = "WANIPConnection"
        impl.version = 2
    End If
    If Not found Then
        ' Try to find WANPPPConnection:2
        serviceType = "urn:schemas-upnp-org:service:WANPPPConnection:2"
        service = xml_find_matching_child(http_response.body, "service", "serviceType", serviceType, found)
        If found Then
            impl.service = "WANPPPConnection"
            impl.version = 2
        End If
    End If
    If Not found Then
        ' Try to find WANIPConnection:1
        serviceType = "urn:schemas-upnp-org:service:WANIPConnection:1"
        service = xml_find_matching_child(http_response.body, "service", "serviceType", serviceType, found)
        If found Then
            impl.service = "WANIPConnection"
            impl.version = 1
        End If
    End If
    If Not found Then
        ' Try to find WANPPPConnection:1
        serviceType = "urn:schemas-upnp-org:service:WANPPPConnection:1"
        service = xml_find_matching_child(http_response.body, "service", "serviceType", serviceType, found)
        If found Then
            impl.service = "WANPPPConnection"
            impl.version = 1
        End If
    End If
    If Not found Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "WANIPConnection not found in UPnP-IGD services"
        http_free http_response
        upnp_impl_query_control_url = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If

    If xml_extract(service, "controlURL", control_url, HTTP_MAX_URL_LEN) <= 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Missing control URL for UPnP-IGN service"
        http_free http_response
        upnp_impl_query_control_url = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If

    If Left$(control_url, 1) = "/" Then
        base_url = impl.location_url            ' strcpy(base_url, impl->location_url);
        p = InStr(base_url, "://")              ' char *p = strstr(base_url, "://");
        If p > 0 Then p = InStr(p + 3, base_url, "/") Else p = InStr(base_url, "/")
        If p > 0 Then base_url = Left$(base_url, p - 1)

        tmp = base_url & control_url
        If Len(tmp) >= HTTP_MAX_URL_LEN Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "WANIPConnection service control URL is too long"
            http_free http_response
            upnp_impl_query_control_url = PROTOCOL_ERR_PROTOCOL_FAILED
            Exit Function
        End If

        control_url = tmp
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "UPnP-IGP %s:%d control URL: %s", impl.service, impl.version, control_url
    impl.control_url = control_url

    http_free http_response
    upnp_impl_query_control_url = PROTOCOL_ERR_SUCCESS
End Function

' int upnp_impl_query_external_addr(upnp_impl_t *impl, timestamp_t end_timestamp)
Private Function upnp_impl_query_external_addr(ByVal end_timestamp As Double) As Long
    Dim zeroRequest As http_request_t, zeroResponse As http_response_t, ret As Long, header_buffer As String, body_buffer As String
    If query_at = 1 Then query_at = 0: GoTo resume_http     ' (VB) resumed
    If impl.control_url = "" Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Attempted to query external address with unknown control URL"
        upnp_impl_query_external_addr = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If

    http_request = zeroRequest
    http_request.method = HTTP_METHOD_POST
    http_request.url = impl.control_url

    header_buffer = va("SOAPAction: urn:schemas-upnp-org:service:%s:%d#GetExternalIPAddress\r\n", impl.service, impl.version)
    If Len(header_buffer) <= 0 Or Len(header_buffer) >= UPNP_BUFFER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to format SOAP request headers"
        upnp_impl_query_external_addr = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If
    http_request.headers = header_buffer

    body_buffer = va("<?xml version=\""1.0\""?>\r\n" & _
                     "<s:Envelope xmlns:s=\""http://schemas.xmlsoap.org/soap/envelope/\"" " & _
                     "s:encodingStyle=\""http://schemas.xmlsoap.org/soap/encoding/\"">" & _
                     "<s:Body>" & _
                     "<m:GetExternalIPAddress " & _
                     "xmlns:m=\""urn:schemas-upnp-org:service:%s:%d\"">" & _
                     "</m:GetExternalIPAddress>" & _
                     "</s:Body>" & _
                     "</s:Envelope>\r\n", _
                     impl.service, impl.version)
    If Len(body_buffer) <= 0 Or Len(body_buffer) >= UPNP_BUFFER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to format SOAP request body"
        upnp_impl_query_external_addr = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If
    http_request.body = body_buffer
    http_request.body_size = Len(body_buffer)
    http_request.body_type = "text/xml; charset=""utf-8"""

    http_response = zeroResponse
resume_http:
    ret = http_perform(http_request, http_response, end_timestamp)
    If ret = PLUM_PENDING Then query_at = 1: upnp_impl_query_external_addr = PLUM_PENDING: Exit Function
    If ret < 0 Then
        upnp_impl_query_external_addr = http_error(ret)
        Exit Function
    End If

    If ret <> 200 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "HTTP request to UPnP-IGD device failed, code=%d", ret
        http_free http_response
        upnp_impl_query_external_addr = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If

    If xml_extract(http_response.body, "NewExternalIPAddress", impl.external_addr_str, ADDR_MAX_STRING_LEN) <= 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Missing external address in response"
        http_free http_response
        upnp_impl_query_external_addr = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "UPnP-IGP WANIPConnection external address: %s", impl.external_addr_str

    http_free http_response
    upnp_impl_query_external_addr = PROTOCOL_ERR_SUCCESS
End Function

' (VB) a non-200 answer's UPnP error code, as upnp_impl_map and upnp_impl_unmap read it
Private Function upnp_error_code(ByVal ret As Long) As Long
    Dim error_code_buffer As String, error_code As Long
    error_code = 0
    If xml_extract(http_response.body, "errorCode", error_code_buffer, 10) > 0 Then error_code = atoi(error_code_buffer)
    If error_code <= 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Got unknown UPnP error"
        upnp_error_code = PROTOCOL_ERR_PROTOCOL_FAILED
        Exit Function
    End If

    plum_log_write PLUM_LOG_LEVEL_WARN, "Got UPnP error, code=%d", error_code
    http_free http_response
    upnp_error_code = error_code
End Function

' int upnp_impl_map(upnp_impl_t *impl, plum_ip_protocol_t protocol, uint16_t external_port, uint16_t internal_port,
'                   unsigned int lifetime, timestamp_t end_timestamp)
Private Function upnp_impl_map(ByVal protocol As Long, ByVal external_port As Long, ByVal internal_port As Long, _
        ByVal lifetime As Long, ByVal end_timestamp As Double) As Long
    Dim local_ As addr_record_t, local_str As String, description As String, zeroRequest As http_request_t
    Dim zeroResponse As http_response_t, header_buffer As String, body_buffer As String, ret As Long
    If query_at = 1 Then query_at = 0: GoTo resume_http     ' (VB) resumed
    If impl.control_url = "" Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Attempted to map with UPnP with unknown control URL"
        upnp_impl_map = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If

    If net_get_default_interface(AF_INET, local_) <> 0 Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Unable to get default interface address"
        upnp_impl_map = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If

    addr_record_to_string local_, local_str, ADDR_MAX_STRING_LEN

    description = "libplum"                     ' TODO

    http_request = zeroRequest
    http_request.method = HTTP_METHOD_POST
    http_request.url = impl.control_url

    header_buffer = va("SOAPAction: urn:schemas-upnp-org:service:%s:%d#AddPortMapping\r\n", impl.service, impl.version)
    If Len(header_buffer) <= 0 Or Len(header_buffer) >= UPNP_BUFFER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to format SOAP request headers"
        upnp_impl_map = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If
    http_request.headers = header_buffer

    body_buffer = va("<?xml version=\""1.0\""?>\r\n" & _
                     "<s:Envelope xmlns:s=\""http://schemas.xmlsoap.org/soap/envelope/\"" " & _
                     "s:encodingStyle=\""http://schemas.xmlsoap.org/soap/encoding/\"">" & _
                     "<s:Body>" & _
                     "<m:AddPortMapping " & _
                     "xmlns:m=\""urn:schemas-upnp-org:service:%s:%d\"">" & _
                     "<NewRemoteHost></NewRemoteHost>" & _
                     "<NewExternalPort>%hu</NewExternalPort>" & _
                     "<NewProtocol>%s</NewProtocol>" & _
                     "<NewInternalPort>%hu</NewInternalPort>" & _
                     "<NewInternalClient>%s</NewInternalClient>" & _
                     "<NewEnabled>1</NewEnabled>" & _
                     "<NewPortMappingDescription>%s</NewPortMappingDescription>" & _
                     "<NewLeaseDuration>%u</NewLeaseDuration>" & _
                     "</m:AddPortMapping>" & _
                     "</s:Body>" & _
                     "</s:Envelope>\r\n", _
                     impl.service, impl.version, _
                     external_port, IIf(protocol = PLUM_IP_PROTOCOL_UDP, "UDP", "TCP"), _
                     internal_port, local_str, description, U32(lifetime))
    If Len(body_buffer) <= 0 Or Len(body_buffer) >= UPNP_BUFFER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to format SOAP request body"
        upnp_impl_map = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If
    http_request.body = body_buffer
    http_request.body_size = Len(body_buffer)
    http_request.body_type = "text/xml; charset=""utf-8"""

    http_response = zeroResponse
resume_http:
    ret = http_perform(http_request, http_response, end_timestamp)
    If ret = PLUM_PENDING Then query_at = 1: upnp_impl_map = PLUM_PENDING: Exit Function
    If ret < 0 Then
        upnp_impl_map = http_error(ret)
        Exit Function
    End If

    If ret <> 200 Then
        upnp_impl_map = upnp_error_code(ret)
        Exit Function
    End If

    http_free http_response
    upnp_impl_map = PROTOCOL_ERR_SUCCESS
End Function

' int upnp_impl_unmap(upnp_impl_t *impl, plum_ip_protocol_t protocol, uint16_t external_port, timestamp_t end_timestamp)
Private Function upnp_impl_unmap(ByVal protocol As Long, ByVal external_port As Long, ByVal end_timestamp As Double) As Long
    Dim zeroRequest As http_request_t, zeroResponse As http_response_t, header_buffer As String, body_buffer As String
    Dim ret As Long
    If query_at = 1 Then query_at = 0: GoTo resume_http     ' (VB) resumed
    If impl.control_url = "" Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Attempted to unmap with UPnP with unknown control URL"
        upnp_impl_unmap = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If

    http_request = zeroRequest
    http_request.method = HTTP_METHOD_POST
    http_request.url = impl.control_url

    header_buffer = va("SOAPAction: urn:schemas-upnp-org:service:%s:%d#DeletePortMapping\r\n", impl.service, impl.version)
    If Len(header_buffer) <= 0 Or Len(header_buffer) >= UPNP_BUFFER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to format SOAP request headers"
        upnp_impl_unmap = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If
    http_request.headers = header_buffer

    body_buffer = va("<?xml version=\""1.0\""?>\r\n" & _
                     "<s:Envelope xmlns:s=\""http://schemas.xmlsoap.org/soap/envelope/\"" " & _
                     "s:encodingStyle=\""http://schemas.xmlsoap.org/soap/encoding/\"">" & _
                     "<s:Body>" & _
                     "<m:DeletePortMapping xmlns:m=\""urn:schemas-upnp-org:service:%s:%d\"">" & _
                     "<NewRemoteHost></NewRemoteHost>" & _
                     "<NewExternalPort>%hu</NewExternalPort>" & _
                     "<NewProtocol>%s</NewProtocol>" & _
                     "</m:DeletePortMapping>" & _
                     "</s:Body>" & _
                     "</s:Envelope>\r\n", _
                     impl.service, impl.version, external_port, _
                     IIf(protocol = PLUM_IP_PROTOCOL_UDP, "UDP", "TCP"))
    If Len(body_buffer) <= 0 Or Len(body_buffer) >= UPNP_BUFFER_SIZE Then
        plum_log_write PLUM_LOG_LEVEL_ERROR, "Failed to format SOAP request body"
        upnp_impl_unmap = PROTOCOL_ERR_UNKNOWN
        Exit Function
    End If
    http_request.body = body_buffer
    http_request.body_size = Len(body_buffer)
    http_request.body_type = "text/xml; charset=""utf-8"""

    http_response = zeroResponse
resume_http:
    ret = http_perform(http_request, http_response, end_timestamp)
    If ret = PLUM_PENDING Then query_at = 1: upnp_impl_unmap = PLUM_PENDING: Exit Function
    If ret < 0 Then
        upnp_impl_unmap = http_error(ret)
        Exit Function
    End If

    If ret <> 200 Then
        upnp_impl_unmap = upnp_error_code(ret)
        Exit Function
    End If

    http_free http_response
    upnp_impl_unmap = PROTOCOL_ERR_SUCCESS
End Function

' int upnp_impl_wait_response(upnp_impl_t *impl, char *buffer, size_t size, addr_record_t *src,
'                             timestamp_t end_timestamp, bool interruptible)
' (VB) a call that returns PLUM_PENDING keeps no state: made again, it goes on as its loop would
Private Function upnp_impl_wait_response(buffer() As Byte, ByVal size As Long, src As addr_record_t, _
        ByVal end_timestamp As Double, ByVal interruptible As Boolean) As Long
    Dim timediff As Double, interrupt As Long, pfd(0 To 0) As pollfd, ret As Long, len_ As Long
    Do
        timediff = end_timestamp - current_timestamp()
        If timediff <= 0 Then Exit Do

        interrupt = impl.interrupt
        If interrupt = UPNP_INTERRUPT_HARD Or (interrupt = UPNP_INTERRUPT_SOFT And interruptible) Then
            plum_log_write PLUM_LOG_LEVEL_VERBOSE, "UPnP interrupted"
            impl.interrupt = UPNP_INTERRUPT_NONE
            upnp_impl_wait_response = PROTOCOL_ERR_INTERRUPTED
            Exit Function
        End If

        pfd(0).fd = impl.sock
        pfd(0).events = POLLIN

        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Entering poll for %d ms", CLng(timediff)
        ret = poll(pfd, 1, 0)                   ' poll(&pfd, 1, (int)timediff)
        If ret = 0 Then upnp_impl_wait_response = PLUM_PENDING: Exit Function  ' (VB)
        If ret < 0 Then
            If sockerrno() = SEINTR Or sockerrno() = SEAGAIN Then
                plum_log_write PLUM_LOG_LEVEL_VERBOSE, "poll interrupted"
                GoTo continue_
            Else
                plum_log_write PLUM_LOG_LEVEL_ERROR, "poll failed, errno=%d", sockerrno()
                upnp_impl_wait_response = PROTOCOL_ERR_UNKNOWN
                Exit Function
            End If
        End If

        plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Exiting poll"

        ' if (ret == 0) break; timeout: (VB) a pass that finds nothing ends above

        If (pfd(0).revents And POLLNVAL) <> 0 Or (pfd(0).revents And POLLERR) <> 0 Then
            plum_log_write PLUM_LOG_LEVEL_ERROR, "Error when polling socket"
            upnp_impl_wait_response = PROTOCOL_ERR_UNKNOWN
            Exit Function
        End If

        If (pfd(0).revents And POLLIN) <> 0 Then
            len_ = udp_recvfrom(pfd(0).fd, buffer, size, src)
            If len_ < 0 Then
                If sockerrno() = SEAGAIN Or sockerrno() = SEWOULDBLOCK Then GoTo continue_

                plum_log_write PLUM_LOG_LEVEL_WARN, "UDP recvfrom failed, errno=%d", sockerrno()
                upnp_impl_wait_response = PROTOCOL_ERR_NETWORK_FAILED
                Exit Function
            End If

            If len_ > 0 Then                    ' 0-length datagrams are used to interrupt, ignore them
                upnp_impl_wait_response = len_
                Exit Function
            End If
        End If
continue_:
    Loop

    upnp_impl_wait_response = PROTOCOL_ERR_TIMEOUT
End Function
