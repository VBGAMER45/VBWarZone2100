Attribute VB_Name = "netpermissions_c"
' lib/netplay/netpermissions.cpp, netpermissions.h: the IP ban list (IPlist, a std::deque<PLAYER_IP>; the
' write dir's banlist.txt is read when hosting starts), isLoopbackIP and the identities' connect permissions
' (identityPermissions: std::unordered_map<std::string, IdentityPermissions>, arrays here).
' onBanList matches each entry with RE2::FullMatch (the entry is a regular expression): regexFullMatch covers
' literals, '.', classes ([...], [^...], ranges), \d \w \s and escaped characters, and the quantifiers * + ?
' (RE2's alternation, groups and counted repetitions are not supported: no ban entry the game writes uses them).
Option Explicit

Public Const MAX_BANS As Long = 1024

' (netpermissions.h's enum class ConnectPermissions: NetApi's)

' struct PLAYER_IP
Private IPlist_pname() As String                ' char pname[40]
Private IPlist_IPAddress() As String            ' char IPAddress[40]
Private IPlistN As Long
Public loadedUserConfigBanList As Boolean

' identityPermissions: identityStr -> IdentityPermissions { optional<ConnectPermissions> connectPermissions; }
Private permKey() As String
Private permHas() As Boolean                    ' connectPermissions.has_value()
Private permValue() As Long
Private permN As Long

' ////////////////////////////////////////////////////////////////////////
Public Sub NETpermissionsInit()
    NETloadUserConfigBanList
End Sub

Public Sub NETpermissionsShutdown()
    IPlistN = 0                                 ' IPlist.clear()
End Sub

' ////////////////////////////////////////////////////////////////////////
Private Sub NETloadUserConfigBanList()
    Dim BanListPath As String, f As Integer, BanStringBuf As String, ToBanIP As String, ToBanName As String, p As Long
    loadedUserConfigBanList = True
    BanListPath = Left$(PHYSFS_getWriteDir(), 4095)
    BanListPath = Left$(BanListPath & "/banlist.txt", 4095)
    If Dir$(BanListPath) = "" Then
        Exit Sub
    End If
    f = FreeFile
    On Error GoTo cannotOpen
    Open BanListPath For Input As #f
    On Error GoTo 0
    debug_ 77, LOG_INFO, "NETloadUserConfigBanList", "Reading banlist file: [%s]" & vbLf, BanListPath
    Do While Not EOF(f)
        Line Input #f, BanStringBuf
        ' sscanf(BanStringBuf, "%255s %255[^\n]", ToBanIP, ToBanName) != 2
        BanStringBuf = Replace(BanStringBuf, vbCr, "")
        p = 1
        Do While p <= Len(BanStringBuf) And InStr(" " & vbTab, Mid$(BanStringBuf, p, 1)) > 0
            p = p + 1
        Loop
        ToBanIP = ""
        Do While p <= Len(BanStringBuf) And InStr(" " & vbTab, Mid$(BanStringBuf, p, 1)) = 0
            ToBanIP = ToBanIP & Mid$(BanStringBuf, p, 1)
            p = p + 1
        Loop
        Do While p <= Len(BanStringBuf) And InStr(" " & vbTab, Mid$(BanStringBuf, p, 1)) > 0
            p = p + 1
        Loop
        ToBanName = Mid$(BanStringBuf, p)
        If ToBanIP = "" Or ToBanName = "" Then
            If Len(BanStringBuf) + 1 > 2 Then  ' strlen(BanStringBuf) counts the '\n' fgets keeps
                debug_ 84, LOG_ERROR, "NETloadUserConfigBanList", "Error reading banlist file!" & vbLf
            End If
        Else
            addIPToBanList Left$(ToBanIP, 255), Left$(ToBanName, 255)
        End If
    Loop
    Close #f
    Exit Sub
cannotOpen:
End Sub

Public Function isLoopbackIP(ByVal ip As String) As Boolean
    If Left$(ip, 4) = "127." Then
        isLoopbackIP = True
        Exit Function
    End If
    If ip = "::1" Then
        isLoopbackIP = True
        Exit Function
    End If
    If ip = "0000:0000:0000:0000:0000:0000:0000:0001" Then
        isLoopbackIP = True
        Exit Function
    End If
    isLoopbackIP = False
End Function

' Check if ip is on the banned list.
' \param ip IP address converted to text
Public Function onBanList(ByVal ip As String) As Boolean
    Dim k As Long
    If IPlistN = 0 Then
        onBanList = False    'if no bans are added, then don't check.
        Exit Function
    End If
    If isLoopbackIP(ip) Then
        onBanList = False   ' ignore loopback IPs
        Exit Function
    End If
    For k = 0 To IPlistN - 1
        If regexFullMatch(ip, IPlist_IPAddress(k)) Then
            onBanList = True
            Exit Function
        End If
    Next
    netGame.wzCommandInterfaceOutput va("WZEVENT: bancheck: %s" & vbLf, ip)
    onBanList = False
End Function

' Create the banned list.
' \param ip IP address in text format
' \param name Name of the player we are banning
Public Sub addIPToBanList(ByVal ip As String, ByVal name_ As String)
    Dim k As Long
    If isLoopbackIP(ip) Then
        Exit Sub
    End If
    If IPlistN = 0 Then
        ReDim IPlist_pname(0 To 15)
        ReDim IPlist_IPAddress(0 To 15)
    ElseIf IPlistN > UBound(IPlist_pname) Then
        ReDim Preserve IPlist_pname(0 To 2 * IPlistN - 1)
        ReDim Preserve IPlist_IPAddress(0 To 2 * IPlistN - 1)
    End If
    IPlist_IPAddress(IPlistN) = sstrcpyN(ip, 40)
    IPlist_pname(IPlistN) = sstrcpyN(name_, 40)
    IPlistN = IPlistN + 1
    sync_counter.banned = (sync_counter.banned + 1) And &HFFFF&
    If IPlistN > MAX_BANS Then
        debug_ 153, LOG_INFO, "addIPToBanList", "We have exceeded %d bans, clearing oldest", MAX_BANS
        For k = 0 To IPlistN - 2                ' IPlist.pop_front()
            IPlist_IPAddress(k) = IPlist_IPAddress(k + 1)
            IPlist_pname(k) = IPlist_pname(k + 1)
        Next
        IPlistN = IPlistN - 1
    End If
End Sub

Public Function removeIPFromBanList(ByVal ip As String) As Boolean
    Dim k As Long, j As Long
    For k = 0 To IPlistN - 1
        If IPlist_IPAddress(k) = ip Then
            For j = k To IPlistN - 2
                IPlist_IPAddress(j) = IPlist_IPAddress(j + 1)
                IPlist_pname(j) = IPlist_pname(j + 1)
            Next
            IPlistN = IPlistN - 1
            removeIPFromBanList = True
            Exit Function
        End If
    Next
    removeIPFromBanList = False
End Function

' std::vector<PLAYER_IP> NETgetIPBanList(): its size, and an entry's fields
Public Function NETgetIPBanList_size() As Long
    NETgetIPBanList_size = IPlistN
End Function

Public Sub NETgetIPBanList_item(ByVal k As Long, pname As String, IPAddress As String)
    pname = IPlist_pname(k)
    IPAddress = IPlist_IPAddress(k)
End Sub

' ////////////////////////////////////////////////////////////////////////

Private Function permFind(ByVal identityStr As String) As Long
    Dim k As Long
    For k = 0 To permN - 1
        If permKey(k) = identityStr Then
            permFind = k
            Exit Function
        End If
    Next
    permFind = -1
End Function

Private Sub permErase(ByVal n As Long)
    Dim k As Long
    For k = n To permN - 2
        permKey(k) = permKey(k + 1)
        permHas(k) = permHas(k + 1)
        permValue(k) = permValue(k + 1)
    Next
    permN = permN - 1
End Sub

' identityStr: Either the the (b64) public key (preferred) or the public hash
Public Sub netPermissionsSet_Connect(ByVal identityStr As String, ByVal perm As Long)
    Dim n As Long
    If identityStr = "" Then Exit Sub
    n = permFind(identityStr)
    If n < 0 Then
        If permN = 0 Then
            ReDim permKey(0 To 15)
            ReDim permHas(0 To 15)
            ReDim permValue(0 To 15)
        ElseIf permN > UBound(permKey) Then
            ReDim Preserve permKey(0 To 2 * permN - 1)
            ReDim Preserve permHas(0 To 2 * permN - 1)
            ReDim Preserve permValue(0 To 2 * permN - 1)
        End If
        n = permN
        permKey(n) = identityStr
        permN = permN + 1
    End If
    permHas(n) = True
    permValue(n) = perm
End Sub

Public Function netPermissionsUnset_Connect(ByVal identityStr As String) As Boolean
    Dim n As Long
    n = permFind(identityStr)
    If n < 0 Then
        netPermissionsUnset_Connect = False
        Exit Function
    End If
    If Not permHas(n) Then
        netPermissionsUnset_Connect = False
        Exit Function
    End If
    permHas(n) = False
    ' (!it->second.hasAnyPermissionsSet(): connectPermissions is the only permission)
    permErase n
    netPermissionsUnset_Connect = True
End Function

Public Function netPermissionsRemoveAll(ByVal identityStr As String) As Boolean
    Dim n As Long
    n = permFind(identityStr)
    If n < 0 Then
        netPermissionsRemoveAll = False
        Exit Function
    End If
    permErase n
    netPermissionsRemoveAll = True
End Function

' optional<ConnectPermissions> netPermissionsCheck_Connect(const std::string& identityStr): has_value() is the result
Public Function netPermissionsCheck_Connect(ByVal identityStr As String, perm As Long) As Boolean
    Dim n As Long
    n = permFind(identityStr)
    If n < 0 Then
        netPermissionsCheck_Connect = False
        Exit Function
    End If
    If permHas(n) Then perm = permValue(n)
    netPermissionsCheck_Connect = permHas(n)
End Function

' optional<ConnectPermissions> netPermissionsCheck_Connect(const EcKey& identity)
Public Function netPermissionsCheck_ConnectKey(identity As EcKey, perm As Long) As Boolean
    Dim key() As Byte, n As Long, checkPublicKeyB64 As String, checkIdentityHash As String
    ' Check both the (b64) public key and the public hash
    ' (b64 public key always takes precedence)
    n = EcKey_toBytes(identity, EcKey_Public, key)
    checkPublicKeyB64 = base64Encode(key, n)
    If netPermissionsCheck_Connect(checkPublicKeyB64, perm) Then
        netPermissionsCheck_ConnectKey = True
        Exit Function
    End If
    checkIdentityHash = EcKey_publicHashString(identity)
    netPermissionsCheck_ConnectKey = netPermissionsCheck_Connect(checkIdentityHash, perm)
End Function

'--------------------------------------------------------------------------
' (VB) RE2::FullMatch(text, pattern): a backtracking matcher for the subset described above

Public Function regexFullMatch(ByVal text As String, ByVal pattern As String) As Boolean
    regexFullMatch = matchHere(text, 1, pattern, 1)
End Function

' the atom at pattern(pi): its end (the index after it) and whether ch matches it
Private Function atomEnd(ByVal pattern As String, ByVal pi As Long) As Long
    Dim c As String
    c = Mid$(pattern, pi, 1)
    If c = "\" Then
        atomEnd = pi + 2
    ElseIf c = "[" Then
        pi = pi + 1
        If Mid$(pattern, pi, 1) = "^" Then pi = pi + 1
        If Mid$(pattern, pi, 1) = "]" Then pi = pi + 1  ' a ']' first is literal
        Do While pi <= Len(pattern) And Mid$(pattern, pi, 1) <> "]"
            If Mid$(pattern, pi, 1) = "\" Then pi = pi + 1
            pi = pi + 1
        Loop
        atomEnd = pi + 1
    Else
        atomEnd = pi + 1
    End If
End Function

Private Function classEscape(ByVal e As String, ByVal ch As String) As Boolean
    Select Case e
    Case "d": classEscape = ch Like "[0-9]"
    Case "w": classEscape = ch Like "[A-Za-z0-9_]"
    Case "s": classEscape = InStr(" " & vbTab & vbCr & vbLf & vbFormFeed & vbVerticalTab, ch) > 0
    Case Else: classEscape = (ch = e)
    End Select
End Function

Private Function atomMatches(ByVal pattern As String, ByVal pi As Long, ByVal ae As Long, ByVal ch As String) As Boolean
    Dim c As String, neg As Boolean, k As Long, lo As String, hi As String, hit As Boolean
    c = Mid$(pattern, pi, 1)
    If c = "." Then
        atomMatches = (ch <> vbLf)
    ElseIf c = "\" Then
        atomMatches = classEscape(Mid$(pattern, pi + 1, 1), ch)
    ElseIf c = "[" Then
        k = pi + 1
        If Mid$(pattern, k, 1) = "^" Then neg = True: k = k + 1
        Do While k < ae - 1
            If Mid$(pattern, k, 1) = "\" Then
                If classEscape(Mid$(pattern, k + 1, 1), ch) Then hit = True
                k = k + 2
            ElseIf Mid$(pattern, k + 1, 1) = "-" And k + 2 < ae - 1 Then
                lo = Mid$(pattern, k, 1)
                hi = Mid$(pattern, k + 2, 1)
                If AscW(ch) >= AscW(lo) And AscW(ch) <= AscW(hi) Then hit = True
                k = k + 3
            Else
                If Mid$(pattern, k, 1) = ch Then hit = True
                k = k + 1
            End If
        Loop
        atomMatches = hit Xor neg
    Else
        atomMatches = (c = ch)
    End If
End Function

Private Function matchHere(ByVal text As String, ByVal ti As Long, ByVal pattern As String, ByVal pi As Long) As Boolean
    Dim ae As Long, q As String, k As Long
    If pi > Len(pattern) Then
        matchHere = (ti > Len(text))
        Exit Function
    End If
    ae = atomEnd(pattern, pi)
    q = Mid$(pattern, ae, 1)
    If q = "*" Or q = "+" Or q = "?" Then
        ' greedy: count how many characters the atom can take, then back off
        k = 0
        Do While ti + k <= Len(text)
            If Not atomMatches(pattern, pi, ae, Mid$(text, ti + k, 1)) Then Exit Do
            k = k + 1
            If q = "?" Then Exit Do
        Loop
        Do While k >= 0
            If q = "+" And k = 0 Then Exit Do
            If matchHere(text, ti + k, pattern, ae + 1) Then
                matchHere = True
                Exit Function
            End If
            k = k - 1
        Loop
        matchHere = False
        Exit Function
    End If
    If ti > Len(text) Then
        matchHere = False
        Exit Function
    End If
    If Not atomMatches(pattern, pi, ae, Mid$(text, ti, 1)) Then
        matchHere = False
        Exit Function
    End If
    matchHere = matchHere(text, ti + 1, pattern, ae)
End Function
