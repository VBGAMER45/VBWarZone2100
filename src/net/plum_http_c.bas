Attribute VB_Name = "plum_http_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/http.c, http.h. The texts (URL, headers, body) are Strings; the socket's bytes are ANSI text.
' (VB) http_perform is resumable (plum_socket_h's PLUM_PENDING): its state is kept here (one request at a time, the
' client thread's), the call that got PLUM_PENDING is made again with the same arguments and goes on where it
' stopped. http_perform_rec's tail call on a redirection is a jump back to its start with the new request (C leaves
' the old connection's socket open: so does this).
Option Explicit

Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long

Private Const INVALID_SOCKET As Long = -1
Private Const HTTP_MAX_RECORDS As Long = 4
Private Const HTTP_MAX_REDIRECTIONS As Long = 5
Private Const DEFAULT_BUFFER_SIZE As Long = 10 * 1024&
Private Const MAX_BUFFER_SIZE As Long = 10 * 1024& * 1024&

Public Const HTTP_MAX_HOST_LEN As Long = 256
Public Const HTTP_MAX_URL_LEN As Long = 1024

Public Const HTTP_ERR_UNKNOWN As Long = -1
Public Const HTTP_ERR_TIMEOUT As Long = -2

' typedef enum http_method
Public Const HTTP_METHOD_GET As Long = 0
Public Const HTTP_METHOD_POST As Long = 1

' typedef struct http_request
Public Type http_request_t
    method As Long
    url As String
    headers As String                           ' vbNullString: NULL
    body As String
    body_type As String
    body_size As Long
End Type

' typedef struct http_response
Public Type http_response_t
    headers As String
    body As String
    body_size As Long
End Type

' (VB) http_perform_rec's state while it waits (see above)
Private rec_at As Long                          ' 0: a new call; 1: connecting; 2: sending the head; 3: the body; 4: receiving
Private rec_request As http_request_t           ' the request (the redirection's new_request after a redirection)
Private rec_redirections As Long
Private rec_records(0 To HTTP_MAX_RECORDS - 1) As addr_record_t
Private rec_n As Long
Private rec_i As Long
Private rec_sock As Long
Private rec_buffer() As Byte
Private rec_size As Long
Private rec_total_len As Long
Private rec_head() As Byte                      ' the request's head (C's buffer before it receives)
Private rec_len As Long
Private rec_body() As Byte
Private rec_path As String                      ' the URL from its path on ("" when it has none)
Private rec_host As String
Private rec_service As String
Private rec_ret As Long

Private Function closesocket(ByVal s As Long) As Long
    closesocket = ws_closesocket(s)
    plum_sock_called
End Function

' static int http_perform_rec(const http_request_t *request, http_response_t *response,
'                             timestamp_t end_timestamp, int redirections)
Private Function http_perform_rec(response As http_response_t, ByVal end_timestamp As Double) As Long
    Dim scheme As String, scheme_len As Long, url As String, url_len As Long, p As Long
    Dim host_len As Long, separator As Long, method_str As String
    Dim text As String, sent As Long, new_size As Long, code As Long, headers_begin As Long, headers_end As Long
    Dim new_url As String, location As String, body_begin As Long, s As String, len_ As Long
    Dim head() As Byte
    Select Case rec_at                          ' (VB) resumed
    Case 1: rec_at = 0: GoTo resume_connect
    Case 2: rec_at = 0: GoTo resume_send
    Case 3: rec_at = 0: GoTo resume_send_body
    Case 4: rec_at = 0: GoTo resume_recv
    End Select

start_:
    scheme = "http://"
    scheme_len = Len(scheme)
    rec_ret = HTTP_ERR_UNKNOWN

    url = rec_request.url
    url_len = Len(url)
    If Left$(url, scheme_len) <> scheme Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "HTTP request URL is invalid: %s", rec_request.url
        http_perform_rec = rec_ret
        Exit Function
    End If

    p = InStr(scheme_len + 1, url, "/")         ' const char *path = strchr(url + scheme_len, '/');
    If p = 0 Then p = url_len + 1               ' path = url + url_len;
    rec_path = Mid$(url, p)

    host_len = p - (scheme_len + 1)
    ' snprintf(host, HTTP_MAX_HOST_LEN, "%.*s", host_len, url + scheme_len) != host_len
    If host_len > HTTP_MAX_HOST_LEN - 1 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to retrieve HTTP host from URL"
        http_perform_rec = rec_ret
        Exit Function
    End If
    rec_host = Mid$(url, scheme_len + 1, host_len)

    separator = InStr(rec_host, ":")
    If separator > 0 Then
        rec_service = Mid$(rec_host, separator + 1)
        rec_host = Left$(rec_host, separator - 1)
    Else
        rec_service = "80"
    End If

    rec_n = addr_resolve(rec_host, rec_service, rec_records, HTTP_MAX_RECORDS)
    If rec_n <= 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to resolve HTTP host %s:%s", rec_host, rec_service
        http_perform_rec = rec_ret
        Exit Function
    End If
    If rec_n > HTTP_MAX_RECORDS Then rec_n = HTTP_MAX_RECORDS

    rec_sock = INVALID_SOCKET
    rec_i = 0
    Do While rec_i < rec_n
resume_connect:
        rec_sock = tcp_connect_socket(rec_records(rec_i), end_timestamp)
        If rec_sock = PLUM_PENDING Then rec_at = 1: http_perform_rec = PLUM_PENDING: Exit Function
        If rec_sock <> INVALID_SOCKET Then Exit Do
        rec_i = rec_i + 1
    Loop

    If rec_sock = INVALID_SOCKET Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to connect to HTTP host %s:%s", rec_host, rec_service
        http_perform_rec = rec_ret
        Exit Function
    End If

    rec_size = DEFAULT_BUFFER_SIZE
    ReDim rec_buffer(0 To rec_size - 1)

    If rec_request.method = HTTP_METHOD_POST Then method_str = "POST" Else method_str = "GET"

    If rec_request.body_size > 0 Then
        text = va("%s %s HTTP/1.0\r\nHost: %s\r\nConnection: close\r\nContent-Length: %zu\r\nContent-Type: %s\r\n%s\r\n", _
                  method_str, IIf(rec_path <> "", rec_path, "/"), rec_host, rec_request.body_size, _
                  rec_request.body_type, IIf(rec_request.headers <> vbNullString, rec_request.headers, ""))
    Else
        text = va("%s %s HTTP/1.0\r\nHost: %s\r\nConnection: close\r\n%s\r\n", _
                  method_str, IIf(rec_path <> "", rec_path, "/"), rec_host, _
                  IIf(rec_request.headers <> vbNullString, rec_request.headers, ""))
    End If
    rec_len = Len(text)

    If rec_len < 0 Or rec_len >= rec_size Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to format HTTP request"
        GoTo error_
    End If
    rec_head = StrConv(text, vbFromUnicode)

    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Sending HTTP request: %s%s", text, IIf(rec_request.body_size > 0, rec_request.body, "")

resume_send:
    sent = tcp_send(rec_sock, rec_head, 0, rec_len, end_timestamp)
    If sent = PLUM_PENDING Then rec_at = 2: http_perform_rec = PLUM_PENDING: Exit Function
    If sent <> rec_len Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to send HTTP request"
        If sent = TCP_ERR_TIMEOUT Then rec_ret = HTTP_ERR_TIMEOUT
        GoTo error_
    End If

    If rec_request.body_size > 0 Then
        rec_body = StrConv(rec_request.body, vbFromUnicode)
resume_send_body:
        sent = tcp_send(rec_sock, rec_body, 0, rec_request.body_size, end_timestamp)
        If sent = PLUM_PENDING Then rec_at = 3: http_perform_rec = PLUM_PENDING: Exit Function
        If sent <> rec_request.body_size Then
            plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to send HTTP request body"
            If sent = TCP_ERR_TIMEOUT Then rec_ret = HTTP_ERR_TIMEOUT
            GoTo error_
        End If
    End If

    rec_total_len = 0
    Do
resume_recv:
        rec_len = tcp_recv(rec_sock, rec_buffer, rec_total_len, rec_size - rec_total_len, end_timestamp)
        If rec_len = PLUM_PENDING Then rec_at = 4: http_perform_rec = PLUM_PENDING: Exit Function
        If rec_len <= 0 Then Exit Do
        rec_total_len = rec_total_len + rec_len

        If rec_total_len = rec_size Then
            If rec_size = MAX_BUFFER_SIZE Then
                plum_log_write PLUM_LOG_LEVEL_WARN, "HTTP response is too big"
                GoTo error_
            End If

            new_size = rec_size * 2
            If new_size > MAX_BUFFER_SIZE Then new_size = MAX_BUFFER_SIZE

            ReDim Preserve rec_buffer(0 To new_size - 1)
            rec_size = new_size
        End If
    Loop

    If rec_len < 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to receive HTTP response"
        If rec_len = TCP_ERR_TIMEOUT Then rec_ret = HTTP_ERR_TIMEOUT
        GoTo error_
    End If

    ' buffer[total_len] = '\0'; null-terminate response
    If rec_total_len > 0 Then
        head = rec_buffer
        ReDim Preserve head(0 To rec_total_len - 1)
        s = StrConv(head, vbUnicode)
    Else
        s = ""
    End If
    p = InStr(s, vbNullChar)
    If p > 0 Then s = Left$(s, p - 1)
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Received HTTP response: %s", s

    code = 0
    If sscanf(s, "HTTP/%*s %d %*s" & vbLf & "%n", code, len_) <> 1 Or code <= 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to parse HTTP response status"
        GoTo error_
    End If

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Got HTTP response code %d", code

    headers_begin = len_ + 1                    ' char *headers_begin = buffer + len;
    headers_end = InStr(headers_begin, s, vbCrLf & vbCrLf)
    If headers_end = 0 Then
        plum_log_write PLUM_LOG_LEVEL_WARN, "Failed to parse HTTP response headers"
        GoTo error_
    End If
    headers_end = headers_end + 2               ' headers_end += 2; *headers_end = '\0';

    ' Handle redirections
    If code >= 300 And code <= 399 And rec_redirections < HTTP_MAX_REDIRECTIONS Then
        len_ = header_extract(Mid$(s, headers_begin, headers_end - headers_begin), "Location", location, HTTP_MAX_URL_LEN)
        If len_ > 0 Then
            If Left$(location, 1) = "/" Then
                new_url = Left$(rec_request.url, p_path_offset(rec_request.url)) & location
                len_ = Len(new_url)
                If len_ > HTTP_MAX_URL_LEN - 1 Then new_url = Left$(new_url, HTTP_MAX_URL_LEN - 1)
            Else
                new_url = location
            End If
        End If
        If len_ > 0 And len_ <= HTTP_MAX_URL_LEN Then
            plum_log_write PLUM_LOG_LEVEL_DEBUG, "Got HTTP redirection to %s", new_url
            rec_request.url = new_url           ' http_request_t new_request = *request; new_request.url = new_url;
            Erase rec_buffer                    ' free(buffer);
            rec_redirections = rec_redirections + 1
            GoTo start_                         ' return http_perform_rec(&new_request, response, end_timestamp, redirections + 1);
        End If
    End If

    response.headers = Mid$(s, headers_begin, headers_end - headers_begin)

    body_begin = headers_end + 2
    response.body = Mid$(s, body_begin)
    response.body_size = rec_total_len - (body_begin - 1)

    plum_log_write PLUM_LOG_LEVEL_DEBUG, "Got HTTP body of size %zu", response.body_size

    closesocket rec_sock
    Erase rec_buffer
    http_perform_rec = code
    Exit Function

error_:
    Erase rec_buffer
    closesocket rec_sock
    http_perform_rec = rec_ret
End Function

' (VB) (int)(path - url): the length of url before its path
Private Function p_path_offset(ByVal url As String) As Long
    Dim p As Long
    p = InStr(Len("http://") + 1, url, "/")
    If p = 0 Then p = Len(url) + 1
    p_path_offset = p - 1
End Function

' int http_perform(const http_request_t *request, http_response_t *response, timestamp_t end_timestamp)
Public Function http_perform(request As http_request_t, response As http_response_t, ByVal end_timestamp As Double) As Long
    If rec_at = 0 Then                          ' (VB) a new call (not resumed)
        rec_request = request
        rec_redirections = 0
    End If
    http_perform = http_perform_rec(response, end_timestamp)
End Function

' void http_free(http_response_t *response)
Public Sub http_free(response As http_response_t)
    response.headers = vbNullString
    response.body = vbNullString
    response.body_size = 0
End Sub
