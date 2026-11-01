Attribute VB_Name = "error_categories_c"
' lib/netplay/error_categories.cpp, error_categories.h, net_result.h, and the parts of MSVC's <system_error> they
' rely on (cluster J's Contract C).
' A std::error_code is a Long, "ec": 0 is no error; otherwise the category in the high byte and the value in the low
' 24 bits (sign-extended back), so an error whose value is 0 still reads as an error:
'   1 EC_GENERIC          std::generic_category(): std::make_error_code(std::errc::x), MSVC's errno numbering;
'   2 EC_GENERIC_SYSTEM   GenericSystemErrorCategory ("generic system (locale-safe)"): make_network_error_code, i.e.
'                         a Winsock code (WSAGetLastError) or, from the files that do not include tcp/sock_error.h
'                         (client_connection.cpp, pending_writes_manager.cpp, ...), an MSVC errno value (EBADF 9,
'                         EINVAL 22, ETIMEDOUT 138, ECONNRESET 108);
'   3 EC_GETADDRINFO      GetaddrinfoErrorCategory: make_getaddrinfo_error_code (getaddrinfo's return value);
'   4 EC_ZLIB             ZlibErrorCategory: make_zlib_error_code.
' net::result<T> functions return T and set ec (0 on success). error_code::message() is net_errorMessage(ec);
' `ec == std::errc::x` is net_errorEqualsErrc(ec, errc_x): MSVC compares the code's default_error_condition, which
' for EC_GENERIC_SYSTEM is system_category()'s (_Winerror_map: the Winsock codes and some Win32 codes map to errc;
' an errno value such as 138 does not, so client_connection's ETIMEDOUT is not errc::timed_out on Windows, as in
' 4.7.0's MSVC build).
Option Explicit

Private Declare Function FormatMessageA Lib "kernel32" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByVal Arguments As Long) As Long

Public Const EC_GENERIC As Long = 1
Public Const EC_GENERIC_SYSTEM As Long = 2
Public Const EC_GETADDRINFO As Long = 3
Public Const EC_ZLIB As Long = 4

' std::errc with MSVC's values (errno.h; its POSIX supplement for the networking ones)
Public Const errc_interrupted As Long = 4
Public Const errc_bad_file_descriptor As Long = 9
Public Const errc_resource_unavailable_try_again As Long = 11
Public Const errc_bad_address As Long = 14
Public Const errc_invalid_argument As Long = 22
Public Const errc_address_in_use As Long = 100
Public Const errc_address_not_available As Long = 101
Public Const errc_address_family_not_supported As Long = 102
Public Const errc_connection_already_in_progress As Long = 103
Public Const errc_connection_aborted As Long = 106
Public Const errc_connection_refused As Long = 107
Public Const errc_connection_reset As Long = 108
Public Const errc_host_unreachable As Long = 110
Public Const errc_operation_in_progress As Long = 112
Public Const errc_already_connected As Long = 113
Public Const errc_network_down As Long = 116
Public Const errc_network_reset As Long = 117
Public Const errc_network_unreachable As Long = 118
Public Const errc_not_connected As Long = 126
Public Const errc_not_a_socket As Long = 128
Public Const errc_timed_out As Long = 138
Public Const errc_operation_would_block As Long = 140

Private Function make_ec(ByVal category As Long, ByVal ev As Long) As Long
    make_ec = (category * &H1000000) Or (ev And &HFFFFFF)
End Function

' std::error_code make_network_error_code(int ev)
Public Function make_network_error_code(ByVal ev As Long) As Long
    make_network_error_code = make_ec(EC_GENERIC_SYSTEM, ev)
End Function

' std::error_code make_getaddrinfo_error_code(int ev)
Public Function make_getaddrinfo_error_code(ByVal ev As Long) As Long
    make_getaddrinfo_error_code = make_ec(EC_GETADDRINFO, ev)
End Function

' std::error_code make_zlib_error_code(int ev)
Public Function make_zlib_error_code(ByVal ev As Long) As Long
    make_zlib_error_code = make_ec(EC_ZLIB, ev)
End Function

' (VB) std::make_error_code(std::errc e): { e, std::generic_category() }
Public Function make_generic_error_code(ByVal errcValue As Long) As Long
    make_generic_error_code = make_ec(EC_GENERIC, errcValue)
End Function

' (VB) ec.value()
Public Function net_errorValue(ByVal ec As Long) As Long
    Dim v As Long
    v = ec And &HFFFFFF
    If (v And &H800000) <> 0 Then v = v Or &HFF000000   ' sign-extend the 24 bits
    net_errorValue = v
End Function

' (VB) which of the categories above (0: no error)
Public Function net_errorCategory(ByVal ec As Long) As Long
    net_errorCategory = (ec And &H7F000000) \ &H1000000
End Function

' (VB) ec.category().name()
Public Function net_errorCategoryName(ByVal ec As Long) As String
    Select Case net_errorCategory(ec)
    Case EC_GENERIC: net_errorCategoryName = "generic"
    Case EC_GENERIC_SYSTEM: net_errorCategoryName = "generic system (locale-safe)"
    Case EC_GETADDRINFO: net_errorCategoryName = "getaddrinfo"
    Case EC_ZLIB: net_errorCategoryName = "zlib"
    Case Else: net_errorCategoryName = "system"
    End Select
End Function

' (VB) ec.message()
Public Function net_errorMessage(ByVal ec As Long) As String
    Dim ev As Long
    ev = net_errorValue(ec)
    Select Case net_errorCategory(ec)
    Case EC_GENERIC_SYSTEM
        net_errorMessage = GenericSystemErrorCategory_message(ev)
    Case EC_GETADDRINFO
        net_errorMessage = GetaddrinfoErrorCategory_message(ev)
    Case EC_ZLIB
        net_errorMessage = ZlibErrorCategory_message(ev)
    Case Else
        net_errorMessage = generic_category_message(ev)
    End Select
End Function

' std::string GenericSystemErrorCategory::message(int ev) const (the WZ_OS_WIN branch)
Private Function GenericSystemErrorCategory_message(ByVal ev As Long) As String
    Dim s As String
    Select Case ev
    Case 0:     s = "No error"
    Case 10004: s = "Interrupted system call"               ' WSAEINTR
    Case 10009: s = "Bad file number"                       ' WSAEBADF
    Case 10013: s = "Permission denied"                     ' WSAEACCES
    Case 10014: s = "Bad address"                           ' WSAEFAULT
    Case 10022: s = "Invalid argument"                      ' WSAEINVAL
    Case 10024: s = "Too many open sockets"                 ' WSAEMFILE
    Case 10035: s = "Operation would block"                 ' WSAEWOULDBLOCK
    Case 10036: s = "Operation now in progress"             ' WSAEINPROGRESS
    Case 10037: s = "Operation already in progress"         ' WSAEALREADY
    Case 10038: s = "Socket operation on non-socket"        ' WSAENOTSOCK
    Case 10039: s = "Destination address required"          ' WSAEDESTADDRREQ
    Case 10040: s = "Message too long"                      ' WSAEMSGSIZE
    Case 10041: s = "Protocol wrong type for socket"        ' WSAEPROTOTYPE
    Case 10042: s = "Bad protocol option"                   ' WSAENOPROTOOPT
    Case 10043: s = "Protocol not supported"                ' WSAEPROTONOSUPPORT
    Case 10044: s = "Socket type not supported"             ' WSAESOCKTNOSUPPORT
    Case 10045: s = "Operation not supported on socket"     ' WSAEOPNOTSUPP
    Case 10046: s = "Protocol family not supported"         ' WSAEPFNOSUPPORT
    Case 10047: s = "Address family not supported"          ' WSAEAFNOSUPPORT
    Case 10048: s = "Address already in use"                ' WSAEADDRINUSE
    Case 10049: s = "Can't assign requested address"        ' WSAEADDRNOTAVAIL
    Case 10050: s = "Network is down"                       ' WSAENETDOWN
    Case 10051: s = "Network is unreachable"                ' WSAENETUNREACH
    Case 10052: s = "Net connection reset"                  ' WSAENETRESET
    Case 10053: s = "Software caused connection abort"      ' WSAECONNABORTED
    Case 10054: s = "Connection reset by peer"              ' WSAECONNRESET
    Case 10055: s = "No buffer space available"             ' WSAENOBUFS
    Case 10056: s = "Socket is already connected"           ' WSAEISCONN
    Case 10057: s = "Socket is not connected"               ' WSAENOTCONN
    Case 10058: s = "Can't send after socket shutdown"      ' WSAESHUTDOWN
    Case 10059: s = "Too many references, can't splice"     ' WSAETOOMANYREFS
    Case 10060: s = "Connection timed out"                  ' WSAETIMEDOUT
    Case 10061: s = "Connection refused"                    ' WSAECONNREFUSED
    Case 10062: s = "Too many levels of symbolic links"     ' WSAELOOP
    Case 10063: s = "File name too long"                    ' WSAENAMETOOLONG
    Case 10064: s = "Host is down"                          ' WSAEHOSTDOWN
    Case 10065: s = "No route to host"                      ' WSAEHOSTUNREACH
    Case 10066: s = "Directory not empty"                   ' WSAENOTEMPTY
    Case 10067: s = "Too many processes"                    ' WSAEPROCLIM
    Case 10068: s = "Too many users"                        ' WSAEUSERS
    Case 10069: s = "Disc quota exceeded"                   ' WSAEDQUOT
    Case 10070: s = "Stale NFS file handle"                 ' WSAESTALE
    Case 10071: s = "Too many levels of remote in path"     ' WSAEREMOTE
    Case 10091: s = "Network system is unavailable"         ' WSASYSNOTREADY
    Case 10092: s = "Winsock version out of range"          ' WSAVERNOTSUPPORTED
    Case 10093: s = "WSAStartup not yet called"             ' WSANOTINITIALISED
    Case 10101: s = "Graceful shutdown in progress"         ' WSAEDISCON
    Case 11001: s = "Host not found"                        ' WSAHOST_NOT_FOUND
    Case 11004: s = "No host data of that type was found"   ' WSANO_DATA
    Case Else:  s = "Unknown error"
    End Select
    GenericSystemErrorCategory_message = s
End Function

' std::string GetaddrinfoErrorCategory::message(int ev) const: gai_strerror(ev), on Windows ws2tcpip.h's inline
' gai_strerrorA: FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS |
' FORMAT_MESSAGE_MAX_WIDTH_MASK, NULL, ecode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), buff,
' GAI_STRERROR_BUFFER_SIZE (1024), NULL) into a static buffer (the system's text, in its language)
Private Function GetaddrinfoErrorCategory_message(ByVal ev As Long) As String
    Dim buff As String, n As Long
    buff = String$(1025, vbNullChar)
    n = FormatMessageA(&H1000 Or &H200 Or &HFF, 0, ev, &H400, buff, 1024, 0)
    GetaddrinfoErrorCategory_message = Left$(buff, n)
End Function

' std::string ZlibErrorCategory::message(int ev) const
Private Function ZlibErrorCategory_message(ByVal ev As Long) As String
    Select Case ev
    Case Z_STREAM_ERROR
        ZlibErrorCategory_message = "Z_STREAM_ERROR"
    Case Z_NEED_DICT
        ZlibErrorCategory_message = "Z_NEED_DICT"
    Case Z_DATA_ERROR
        ZlibErrorCategory_message = "Z_DATA_ERROR"
    Case Z_MEM_ERROR
        ZlibErrorCategory_message = "Z_MEM_ERROR"
    Case Else
        ZlibErrorCategory_message = "Unknown zlib error"
    End Select
End Function

' (VB) MSVC's std::generic_category().message(ev): its _Syserror_map table ("unknown error" otherwise)
Private Function generic_category_message(ByVal ev As Long) As String
    Dim s As String
    Select Case ev
    Case 102: s = "address family not supported"
    Case 100: s = "address in use"
    Case 101: s = "address not available"
    Case 113: s = "already connected"
    Case 7:   s = "argument list too long"
    Case 33:  s = "argument out of domain"
    Case 14:  s = "bad address"
    Case 9:   s = "bad file descriptor"
    Case 104: s = "bad message"
    Case 32:  s = "broken pipe"
    Case 106: s = "connection aborted"
    Case 103: s = "connection already in progress"
    Case 107: s = "connection refused"
    Case 108: s = "connection reset"
    Case 18:  s = "cross device link"
    Case 109: s = "destination address required"
    Case 16:  s = "device or resource busy"
    Case 41:  s = "directory not empty"
    Case 8:   s = "executable format error"
    Case 17:  s = "file exists"
    Case 27:  s = "file too large"
    Case 38:  s = "filename too long"
    Case 40:  s = "function not supported"
    Case 110: s = "host unreachable"
    Case 111: s = "identifier removed"
    Case 42:  s = "illegal byte sequence"
    Case 25:  s = "inappropriate io control operation"
    Case 4:   s = "interrupted"
    Case 22:  s = "invalid argument"
    Case 29:  s = "invalid seek"
    Case 5:   s = "io error"
    Case 21:  s = "is a directory"
    Case 115: s = "message size"
    Case 116: s = "network down"
    Case 117: s = "network reset"
    Case 118: s = "network unreachable"
    Case 119: s = "no buffer space"
    Case 10:  s = "no child process"
    Case 121: s = "no link"
    Case 39:  s = "no lock available"
    Case 120: s = "no message available"
    Case 122: s = "no message"
    Case 123: s = "no protocol option"
    Case 28:  s = "no space on device"
    Case 124: s = "no stream resources"
    Case 6:   s = "no such device or address"
    Case 19:  s = "no such device"
    Case 2:   s = "no such file or directory"
    Case 3:   s = "no such process"
    Case 20:  s = "not a directory"
    Case 128: s = "not a socket"
    Case 125: s = "not a stream"
    Case 126: s = "not connected"
    Case 12:  s = "not enough memory"
    Case 129: s = "not supported"
    Case 105: s = "operation canceled"
    Case 112: s = "operation in progress"
    Case 1:   s = "operation not permitted"
    Case 130: s = "operation not supported"
    Case 140: s = "operation would block"
    Case 133: s = "owner dead"
    Case 13:  s = "permission denied"
    Case 134: s = "protocol error"
    Case 135: s = "protocol not supported"
    Case 30:  s = "read only file system"
    Case 36:  s = "resource deadlock would occur"
    Case 11:  s = "resource unavailable try again"
    Case 34:  s = "result out of range"
    Case 127: s = "state not recoverable"
    Case 137: s = "stream timeout"
    Case 139: s = "text file busy"
    Case 138: s = "timed out"
    Case 23:  s = "too many files open in system"
    Case 24:  s = "too many files open"
    Case 31:  s = "too many links"
    Case 114: s = "too many symbolic link levels"
    Case 132: s = "value too large"
    Case 136: s = "wrong protocol type"
    Case Else: s = "unknown error"
    End Select
    generic_category_message = s
End Function

' (VB) MSVC's _Winerror_map: system_category().default_error_condition's mapping of a Win32 / Winsock code to an
' errc (0: none, the condition stays in system_category)
Private Function Winerror_map(ByVal ev As Long) As Long
    Dim p As Long
    Select Case ev
    Case 5, 82, 12, 16, 998, 32, 19, 10013: p = 13         ' permission_denied (ERROR_ACCESS_DENIED, ...)
    Case 183, 80: p = 17                                   ' file_exists
    Case 53, 67, 123, 2, 3: p = 2                          ' no_such_file_or_directory
    Case 20, 55, 15: p = 19                                ' no_such_device
    Case 109: p = 32                                       ' broken_pipe
    Case 111, 10063: p = 38                                ' filename_too_long
    Case 170, 142, 2404, 2401: p = 16                      ' device_or_resource_busy
    Case 1011, 1012, 1013, 110, 30, 25, 29: p = 5          ' io_error
    Case 145: p = 41                                       ' directory_not_empty
    Case 267, 6, 87, 131, 10022: p = 22                    ' invalid_argument
    Case 112, 39: p = 28                                   ' no_space_on_device
    Case 1: p = 40                                         ' function_not_supported
    Case 33, 212: p = 39                                   ' no_lock_available
    Case 8, 14: p = 12                                     ' not_enough_memory
    Case 21, 1237: p = 11                                  ' resource_unavailable_try_again
    Case 17: p = 18                                        ' cross_device_link
    Case 50: p = 129                                       ' not_supported
    Case 995: p = 105                                      ' operation_canceled
    Case 4, 10024: p = 24                                  ' too_many_files_open
    Case 10048: p = 100                                    ' address_in_use
    Case 10049: p = 101                                    ' address_not_available
    Case 10047: p = 102                                    ' address_family_not_supported
    Case 10037: p = 103                                    ' connection_already_in_progress
    Case 10009: p = 9                                      ' bad_file_descriptor
    Case 10053: p = 106                                    ' connection_aborted
    Case 10061: p = 107                                    ' connection_refused
    Case 10054: p = 108                                    ' connection_reset
    Case 10039: p = 109                                    ' destination_address_required
    Case 10014: p = 14                                     ' bad_address
    Case 10065: p = 110                                    ' host_unreachable
    Case 10036: p = 112                                    ' operation_in_progress
    Case 10004: p = 4                                      ' interrupted
    Case 10056: p = 113                                    ' already_connected
    Case 10040: p = 115                                    ' message_size
    Case 10050: p = 116                                    ' network_down
    Case 10052: p = 117                                    ' network_reset
    Case 10051: p = 118                                    ' network_unreachable
    Case 10055: p = 119                                    ' no_buffer_space
    Case 10042: p = 123                                    ' no_protocol_option
    Case 10057: p = 126                                    ' not_connected
    Case 10038: p = 128                                    ' not_a_socket
    Case 10045: p = 130                                    ' operation_not_supported
    Case 10043: p = 135                                    ' protocol_not_supported
    Case 10041: p = 136                                    ' wrong_protocol_type
    Case 10060: p = 138                                    ' timed_out
    Case 10035: p = 140                                    ' operation_would_block
    Case Else: p = 0
    End Select
    Winerror_map = p
End Function

' (VB) `ec == std::errc::x`: ec.category().equivalent(ec.value(), errc) || errc's category's equivalent(ec, ...), i.e.
' ec's default_error_condition == error_condition(errc, generic_category())
Public Function net_errorEqualsErrc(ByVal ec As Long, ByVal errcValue As Long) As Boolean
    Dim ev As Long, p As Long
    If ec = 0 Then Exit Function
    ev = net_errorValue(ec)
    Select Case net_errorCategory(ec)
    Case EC_GENERIC
        net_errorEqualsErrc = (ev = errcValue)
    Case EC_GENERIC_SYSTEM
        ' GenericSystemErrorCategory::default_error_condition: std::system_category().default_error_condition(ev)
        p = Winerror_map(ev)
        If p <> 0 Then net_errorEqualsErrc = (p = errcValue)
    Case Else
        ' getaddrinfo and zlib keep the default: error_condition(ev, *this), never a generic errc
        net_errorEqualsErrc = False
    End Select
End Function
