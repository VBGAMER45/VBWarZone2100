Attribute VB_Name = "multijoin_helpers_c"
' src/multijoin_helpers.cpp, multijoin_helpers.h, whole: joining a game from an IP / hostname string, an IP and port,
' a lobby game id or a connection list. inet_pton is Winsock's InetPtonW (its WSAStartup done once here); std::stoi's
' parse (leading whitespace, a sign, digits up to the first other character, an exception when there are none or
' the value leaves int) is stoiLike. GNS_DIRECT connections exist only with WZ_GNS_NETWORK_BACKEND_ENABLED (not
' defined: GameNetworkingSockets is not ported).
Option Explicit

Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequested As Integer, lpWSAData As Any) As Long
Private Declare Function InetPtonW Lib "ws2_32" (ByVal Family As Long, ByVal pszAddrString As Long, pAddrBuf As Any) As Long

Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23

' enum class JoinGameFromConnectionStrOutcome
Public Enum JoinGameFromConnectionStrOutcome
    JoinGameFromConnectionStrOutcome_Attempting_IP_Join
    JoinGameFromConnectionStrOutcome_Attempting_Lobby_Join
End Enum

' MARK: String parsing helper functions

' struct HostAndPort
Private Type HostAndPort
    host As String
    port As Long                        ' uint16_t = 0
End Type

' struct IPAddressInfo
Private Type IPAddressInfo
    ipVersion As netlobby_IPVersion
    ip As String
    has_port As Boolean                 ' std::optional<int> port
    port As Long
End Type

Private wsaStarted As Boolean

' (VB) std::stoi(s): False where it throws
Private Function stoiLike(ByVal s As String, value As Long) As Boolean
    Dim i As Long, n As Long, c As String, neg As Boolean, digits As Long, acc As Double
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid$(s, i, 1)
        If c = " " Or c = vbTab Or c = vbLf Or c = vbCr Or c = vbVerticalTab Or c = vbFormFeed Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    If i <= n Then
        c = Mid$(s, i, 1)
        If c = "+" Or c = "-" Then
            neg = (c = "-")
            i = i + 1
        End If
    End If
    Do While i <= n
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then Exit Do
        acc = acc * 10 + (AscW(c) - 48)
        digits = digits + 1
        If acc > 2147483648# Then Exit Function    ' std::out_of_range
        i = i + 1
    Loop
    If digits = 0 Then Exit Function               ' std::invalid_argument
    If neg Then acc = -acc
    If acc > 2147483647# Or acc < -2147483648# Then Exit Function
    value = CLng(acc)
    stoiLike = True
End Function

' optional<HostAndPort> splitStrIntoHostAndPort(const std::string& input) (has_value: the result)
Private Function splitStrIntoHostAndPort(ByVal input_ As String, out As HostAndPort) As Boolean
    Dim lastColon As Long, hostPart As String, portPart As String, p As Long
    lastColon = InStrRev(input_, ":") - 1          ' find_last_of(':') (-1: npos)
    If lastColon <> -1 And lastColon < Len(input_) - 1 Then
        hostPart = Left$(input_, lastColon)
        portPart = Mid$(input_, lastColon + 2)

        ' Validate port
        If Not stoiLike(portPart, p) Then
            Exit Function
        End If
        If p < 0 Or p > 65535 Then
            Exit Function
        End If

        out.host = hostPart
        out.port = p
        splitStrIntoHostAndPort = True
        Exit Function
    End If

    splitStrIntoHostAndPort = False
End Function

' MARK: IP address parsing helper functions

' (VB) inet_pton(af, src, dst) == 1
Private Function inet_pton_ok(ByVal af As Long, ByVal src As String) As Boolean
    Dim wsaData(0 To 407) As Byte, buf(0 To 15) As Byte
    If Not wsaStarted Then
        wsaStarted = (WSAStartup(&H202, wsaData(0)) = 0)
    End If
    inet_pton_ok = (InetPtonW(af, StrPtr(src), buf(0)) = 1)
End Function

' optional<IPAddressInfo> parseAsIPAddress(const std::string& input) (has_value: the result)
Private Function parseAsIPAddress(ByVal input_ As String, out As IPAddressInfo) As Boolean
    Dim optHostAndPort As HostAndPort, ipPart As String, p As Long, rawIp6 As String
    If Len(input_) = 0 Then Exit Function

    ' 1. Try raw IPv4
    If inet_pton_ok(AF_INET, input_) Then
        out.ipVersion = netlobby_IPVersion_IPv4
        out.ip = input_
        out.has_port = False
        parseAsIPAddress = True
        Exit Function
    End If

    ' 2. Try raw IPv6
    If inet_pton_ok(AF_INET6, input_) Then
        out.ipVersion = netlobby_IPVersion_IPv6
        out.ip = input_
        out.has_port = False
        parseAsIPAddress = True
        Exit Function
    End If

    ' 3. Check for port suffix
    If splitStrIntoHostAndPort(input_, optHostAndPort) Then
        ipPart = optHostAndPort.host
        p = optHostAndPort.port

        ' Check: IPv4 with port (e.g. 1.2.3.4:80)
        If inet_pton_ok(AF_INET, ipPart) Then
            out.ipVersion = netlobby_IPVersion_IPv4
            out.ip = ipPart
            out.has_port = True
            out.port = p
            parseAsIPAddress = True
            Exit Function
        End If

        ' Check: Bracketed IPv6 with port (e.g. [::1]:80)
        If Len(ipPart) > 2 And Left$(ipPart, 1) = "[" And Right$(ipPart, 1) = "]" Then
            rawIp6 = Mid$(ipPart, 2, Len(ipPart) - 2)
            If inet_pton_ok(AF_INET6, rawIp6) Then
                out.ipVersion = netlobby_IPVersion_IPv6
                out.ip = rawIp6
                out.has_port = True
                out.port = p
                parseAsIPAddress = True
                Exit Function
            End If
        End If
    End If

    parseAsIPAddress = False
End Function

' MARK: - Public API

' Start a join with an expected ip, ip:port, hostname, or hostname:port string
' Note: Does not attempt a lobby join
Public Sub joinGameFromIPOrHostnameConnectionStr(ByVal ipOrHostnameConnectionStr As String, Optional ByVal asSpectator As Boolean = False)
    Dim parseIPAddressDetails As IPAddressInfo, host As String, port As Long, optHostAndPort As HostAndPort
    ' 1. Try to parse as an IP address or IP:port
    If parseAsIPAddress(ipOrHostnameConnectionStr, parseIPAddressDetails) Then
        ' Parsed as an IP address
        ' Start joining attempt
        host = parseIPAddressDetails.ip
        If parseIPAddressDetails.has_port Then port = parseIPAddressDetails.port Else port = 0 ' 0 = try default port
        debug_ 149, LOG_INFO, "joinGameFromIPOrHostnameConnectionStr", "Connecting to ip [%s] port %d", host, port
        joinGameFromIPPort host, port, asSpectator
        Exit Sub
    End If

    ' 2. Try to do a very basic split at the last ":", if one is present
    If splitStrIntoHostAndPort(ipOrHostnameConnectionStr, optHostAndPort) Then
        ' hopefully hostPart is either a valid ip or hostname - give it a try
        debug_ 159, LOG_INFO, "joinGameFromIPOrHostnameConnectionStr", "Connecting to host [%s] port %d", optHostAndPort.host, optHostAndPort.port
        joinGameFromIPPort optHostAndPort.host, optHostAndPort.port, asSpectator
        Exit Sub
    End If

    ' 3. Otherwise, perhaps this is just a hostname (without a port specified) - give that a try with the default port
    joinGameFromIPPort ipOrHostnameConnectionStr, 0, asSpectator
End Sub

' Start a join with an explicit ip & port
Public Sub joinGameFromIPPort(ByVal host As String, ByVal port As Long, Optional ByVal asSpectator As Boolean = False)
    Dim connList(0 To 0) As JoinConnectionDescription
    ' Construct connection list
    connList(0) = JoinConnectionDescription_make(JoinConnectionType_TCP_DIRECT, host, port)
    ' (WZ_GNS_NETWORK_BACKEND_ENABLED: not defined)
    ' Start joining attempt
    startJoiningAttempt sPlayer, connList, 1, asSpectator
End Sub

' Parses a user-provided connection string, which may be either:
' - An IP address or IP:port combination
' - A lobby game ID
' And attempts to start a join
' (optional<JoinGameFromConnectionStrOutcome>: has_value is the result, the outcome in outcome)
Public Function joinGameFromConnectionStr(ByVal userProvidedConnectionStr As String, outcome As JoinGameFromConnectionStrOutcome, _
        Optional ByVal asSpectator As Boolean = False) As Boolean
    Dim parseIPAddressDetails As IPAddressInfo, port As Long
    If Len(userProvidedConnectionStr) = 0 Then
        joinGameFromConnectionStr = False
        Exit Function
    End If

    If parseAsIPAddress(userProvidedConnectionStr, parseIPAddressDetails) Then
        ' Parsed as an IP address
        ' Start joining attempt
        If parseIPAddressDetails.has_port Then port = parseIPAddressDetails.port Else port = 0 ' 0 = try default port
        joinGameFromIPPort parseIPAddressDetails.ip, port, asSpectator
        outcome = JoinGameFromConnectionStrOutcome_Attempting_IP_Join
        joinGameFromConnectionStr = True
        Exit Function
    End If

    ' Otherwise, try treating as a lobby game id
    startLobbyJoiningAttempt sPlayer, NETgetLobbyserverAddress(), userProvidedConnectionStr, asSpectator
    outcome = JoinGameFromConnectionStrOutcome_Attempting_Lobby_Join
    joinGameFromConnectionStr = True
End Function

' Join a game given a list of JoinConnectionDescription
Public Sub joinGame(connection_list() As JoinConnectionDescription, ByVal connection_listCount As Long, Optional ByVal asSpectator As Boolean = False)
    startJoiningAttempt sPlayer, connection_list, connection_listCount, asSpectator
End Sub

' Explicitly joins a lobby game, with lobbyAddress and lobbyGameId
Public Sub joinLobbyGame(ByVal lobbyAddress As String, ByVal lobbyGameId As String, Optional ByVal asSpectator As Boolean = False)
    startLobbyJoiningAttempt sPlayer, lobbyAddress, lobbyGameId, asSpectator
End Sub
