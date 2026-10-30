Attribute VB_Name = "wzpaths_c"
' lib/framework/wzpaths.cpp. Ported: splitAtAnyDelimiter,
' getWZInstallPrefix (WZ_BINDIR is "bin", as in the Windows build).
' Missing: WzPathInfo.
Option Explicit

Private Const WZ_BINDIR As String = "bin"

' C: std::vector<std::string> splitAtAnyDelimiter(s, delimiters); returns the count, the parts in v(0 ..)
Private Function splitAtAnyDelimiter(ByVal s As String, ByVal delimiters As String, v() As String) As Long
    Dim pos As Long, end_ As Long, n As Long

    pos = 1
    Do While pos <= Len(s)
        end_ = pos
        Do While end_ <= Len(s)
            If InStr(delimiters, Mid$(s, end_, 1)) > 0 Then Exit Do
            end_ = end_ + 1
        Loop
        If n = 0 Then ReDim v(0 To 0) Else ReDim Preserve v(0 To n)
        v(n) = Mid$(s, pos, end_ - pos)
        n = n + 1
        pos = end_
        If pos <= Len(s) Then
            pos = pos + 1
        End If
    Loop

    splitAtAnyDelimiter = n
End Function

Public Function getWZInstallPrefix() As String
    Dim dirSeparator As String, prefixDir As String, binDirComponentsCount As Long
    Dim binDir As String, binDirComponents() As String, n As Long, i As Long, lastSlash As Long

    dirSeparator = PHYSFS_getDirSeparator()
    If Not (Len(dirSeparator) > 0) Then
        ASSERT_FAILURE "dirSeparator.length() > 0", "lib/framework/wzpaths.cpp:49", "getWZInstallPrefix", "PHYSFS_getDirSeparator returned an empty string"
    End If

    ' Construct the install PREFIX path
    prefixDir = PHYSFS_getBaseDir()
    Do While Len(prefixDir) > 0
        If Right$(prefixDir, Len(dirSeparator)) <> dirSeparator Then Exit Do
        prefixDir = Left$(prefixDir, Len(prefixDir) - Len(dirSeparator)) ' Remove trailing path separators
    Loop
    binDirComponentsCount = 1
    ' Trim off as many path components as are in WZ_BINDIR
    binDir = WZ_BINDIR
    n = splitAtAnyDelimiter(binDir, "/", binDirComponents)
    binDirComponentsCount = 0
    For i = 0 To n - 1
        If Len(binDirComponents(i)) > 0 And binDirComponents(i) <> "." Then binDirComponentsCount = binDirComponentsCount + 1
    Next
    If Not (binDirComponentsCount > 0) Then
        ASSERT_FAILURE "binDirComponentsCount > 0", "lib/framework/wzpaths.cpp:63", "getWZInstallPrefix", "WZ_BINDIR unexpectedly has zero components?: ""%s""", WZ_BINDIR
    End If
    For i = 0 To binDirComponentsCount - 1
        lastSlash = InStrRev(prefixDir, dirSeparator)
        If lastSlash > 0 Then
            prefixDir = Left$(prefixDir, lastSlash - 1) ' Trim off the last path component
        End If
    Next

    getWZInstallPrefix = prefixDir
End Function
