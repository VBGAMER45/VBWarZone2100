Attribute VB_Name = "urlhelpers_c"
' src/urlhelpers.cpp, src/urlhelpers.h, Windows: urlHasAcceptableProtocol, openURLInBrowser
' (UrlCanonicalizeW, ShellExecuteExW "open", then WzPlatform's wzBackendAttemptOpenURL), urlEncode
' (curl_easy_escape: each UTF-8 byte but A-Z, a-z, 0-9, "-", ".", "_" and "~" as %XX), openFolderInDefaultFileManager
' (ShellExecuteExW "explore"). win_utf8ToUtf16 has no port: VB strings are UTF-16 already.
' TEST-ONLY: with WZ_TEST_NO_DESKTOP_EFFECTS set (PlatformApi's wzTestNoDesktopEffects), nothing is opened: the
' call is logged and succeeds.
Option Explicit

Private Type SHELLEXECUTEINFOW
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As Long
    lpFile As Long
    lpParameters As Long
    lpDirectory As Long
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As Long
    hkeyClass As Long
    dwHotKey As Long
    hIconOrMonitor As Long
    hProcess As Long
End Type

Private Declare Function ShellExecuteExW Lib "shell32" (lpExecInfo As SHELLEXECUTEINFOW) As Long
Private Declare Function UrlCanonicalizeW Lib "shlwapi" (ByVal pszUrl As Long, ByVal pszCanonicalized As Long, pcchCanonicalized As Long, ByVal dwFlags As Long) As Long
Private Declare Function CoInitializeEx Lib "ole32" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long
Private Declare Sub CoUninitialize Lib "ole32" ()

Private Const SEE_MASK_NOASYNC As Long = &H100
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOW As Long = 5
Private Const URL_UNESCAPE As Long = &H10000000
Private Const URL_ESCAPE_UNSAFE As Long = &H20000000
Private Const URL_ESCAPE_AS_UTF8 As Long = &H40000
Private Const COINIT_APARTMENTTHREADED As Long = &H2
Private Const COINIT_DISABLE_OLE1DDE As Long = &H4
Private Const INTERNET_MAX_URL_LENGTH As Long = 2083        ' _INTERNET_MAX_URL_LENGTH

Public Function urlHasAcceptableProtocol(ByVal url As String) As Boolean
    Dim validLinkPrefixes As Variant, i As Long, bValidLinkPrefix As Boolean
    validLinkPrefixes = Array("http://", "https://", "ms-windows-store://")
    For i = 0 To UBound(validLinkPrefixes)
        ' strncasecmp
        If StrComp(Left$(url, Len(validLinkPrefixes(i))), validLinkPrefixes(i), vbTextCompare) = 0 Then
            bValidLinkPrefix = True
            Exit For
        End If
    Next
    urlHasAcceptableProtocol = bValidLinkPrefix
End Function

Public Function openURLInBrowser(ByVal url As String) As Boolean
    Dim wUrlCanonicalized As String, canonicalizedLength As Long, dwFlags As Long, hr As Long, coInitResult As Long
    Dim ShExecInfo As SHELLEXECUTEINFOW, verb As String, bShellExecuteFailure As Boolean, dwError As Long
    Dim succeededOpeningUrl As Boolean

    If Len(url) = 0 Then Exit Function
    If Not urlHasAcceptableProtocol(url) Then
        debug_ 137, LOG_ERROR, "openURLInBrowser", "Rejecting attempt to open link with URL: %s", url
        Exit Function
    End If

    ' Desktop Windows (wUrl: the URL as UTF-16 with its terminating null)
    If Len(url) + 1 > INTERNET_MAX_URL_LENGTH Then
        debug_ 155, LOG_ERROR, "openURLInBrowser", "URL length (%zu) exceeds maximum supported length: %s" & vbLf, Len(url) + 1, url
        Exit Function
    End If

    wUrlCanonicalized = String$(INTERNET_MAX_URL_LENGTH + 1, vbNullChar)
    canonicalizedLength = Len(wUrlCanonicalized)
    dwFlags = URL_UNESCAPE Or URL_ESCAPE_UNSAFE
    ' (IsWindows7OrGreater())
    dwFlags = dwFlags Or URL_ESCAPE_AS_UTF8
    hr = UrlCanonicalizeW(StrPtr(url), StrPtr(wUrlCanonicalized), canonicalizedLength, dwFlags)
    If hr < 0 Then
        debug_ 173, LOG_ERROR, "openURLInBrowser", "Failed to canonicalize URL ('%s') with error: %s", url, "HRESULT 0x" & Hex$(hr)
        Exit Function
    End If
    wUrlCanonicalized = Left$(wUrlCanonicalized, canonicalizedLength)

    If wzTestNoDesktopEffects() Then
        debug_ 0, LOG_INFO, "openURLInBrowser", "(VB) TEST-ONLY: ShellExecuteEx not called: open %s", wUrlCanonicalized
        openURLInBrowser = True
        Exit Function
    End If

    coInitResult = CoInitializeEx(0, COINIT_APARTMENTTHREADED Or COINIT_DISABLE_OLE1DDE)
    verb = "open"
    ShExecInfo.cbSize = LenB(ShExecInfo)
    ShExecInfo.fMask = SEE_MASK_NOASYNC
    ShExecInfo.lpVerb = StrPtr(verb)
    ShExecInfo.lpFile = StrPtr(wUrlCanonicalized)
    ShExecInfo.nShow = SW_SHOWNORMAL
    If ShellExecuteExW(ShExecInfo) = 0 Then
        dwError = Err.LastDllError
        debug_ 187, LOG_ERROR, "openURLInBrowser", "ShellExecuteEx failed with error: %lu", dwError
        bShellExecuteFailure = True
    End If
    If coInitResult = 0 Or coInitResult = 1 Then        ' S_OK, S_FALSE
        CoUninitialize
    End If
    succeededOpeningUrl = Not bShellExecuteFailure

    If succeededOpeningUrl Then
        openURLInBrowser = succeededOpeningUrl
        Exit Function
    End If
    succeededOpeningUrl = wzBackendAttemptOpenURL(url)
    openURLInBrowser = succeededOpeningUrl
End Function

' std::string urlEncode(const char* urlFragment): curl_easy_escape (every byte of the UTF-8 text but the unreserved
' characters as %XX, upper case)
Public Function urlEncode(ByVal urlFragment As String) As String
    Dim b() As Byte, n As Long, i As Long, c As Long, result As String
    n = utf8_encode(urlFragment, b)
    For i = 0 To n - 1
        c = b(i)
        Select Case c
        Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126     ' 0-9 A-Z a-z - . _ ~
            result = result & Chr$(c)
        Case Else
            result = result & "%" & Right$("0" & Hex$(c), 2)
        End Select
    Next
    urlEncode = result
End Function

Public Function openFolderInDefaultFileManager(ByVal path As String) As Boolean
    Dim coInitResult As Long, ShExecInfo As SHELLEXECUTEINFOW, verb As String, bShellExecuteFailure As Boolean, dwError As Long

    If wzTestNoDesktopEffects() Then
        debug_ 0, LOG_INFO, "openFolderInDefaultFileManager", "(VB) TEST-ONLY: ShellExecuteEx not called: explore %s", path
        openFolderInDefaultFileManager = True
        Exit Function
    End If

    coInitResult = CoInitializeEx(0, COINIT_APARTMENTTHREADED Or COINIT_DISABLE_OLE1DDE)
    verb = "explore"
    ShExecInfo.cbSize = LenB(ShExecInfo)
    ShExecInfo.fMask = SEE_MASK_NOASYNC
    ShExecInfo.hwnd = 0
    ShExecInfo.lpVerb = StrPtr(verb)
    ShExecInfo.lpFile = StrPtr(path)
    ShExecInfo.nShow = SW_SHOW
    If ShellExecuteExW(ShExecInfo) = 0 Then
        dwError = Err.LastDllError
        debug_ 275, LOG_ERROR, "openFolderInDefaultFileManager", "ShellExecuteEx failed with error: %lu", dwError
        bShellExecuteFailure = True
    End If
    If coInitResult = 0 Or coInitResult = 1 Then        ' S_OK, S_FALSE
        CoUninitialize
    End If
    openFolderInDefaultFileManager = Not bShellExecuteFailure
End Function
