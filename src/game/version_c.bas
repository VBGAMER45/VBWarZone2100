Attribute VB_Name = "version_c"
' src/version.cpp, on the constants build.py generates into autorevision_vb
' (Warzone: build_tools/autorevision.h). Deviation: the config folder is
' "WarzoneVB" where Warzone's is "Warzone 2100", so the port never touches a
' real installation's settings. TagVer is version(0..2) and qualifier.
' Translations (_()) are not ported: the English text is used.
Option Explicit

Private Const TAGVER_MAX_QUALIF_LEN As Long = 16
Public Const MAX_STR_LENGTH As Long = 256                ' stringdef.h

' C: optional<TagVer> version_extractVersionNumberFromTag(const std::string& tag)
Public Function version_extractVersionNumberFromTag(ByVal tag As String, version() As Long, qualifier As String) As Boolean
    Dim p As Long, i As Long, n As Long, v As Double, digits As Long

    ReDim version(0 To 2)
    qualifier = ""
    p = 1
    ' Remove "v/" or "v" prefix (as in "v3.2.2"), if present
    If Mid$(tag, p, 1) = "v" Then p = p + 1 ' skip
    If Mid$(tag, p, 1) = "/" Then p = p + 1 ' skip

    For i = 0 To 2
        If i > 0 Then p = p + 1                 ' parser.get(): skip any separator
        ' parser >> uint16_t: leading white space, then digits
        Do While Mid$(tag, p, 1) = " " Or Mid$(tag, p, 1) = vbTab
            p = p + 1
        Loop
        v = 0
        digits = 0
        Do While p <= Len(tag)
            If Mid$(tag, p, 1) < "0" Or Mid$(tag, p, 1) > "9" Then Exit Do
            v = v * 10 + (AscW(Mid$(tag, p, 1)) - 48)
            digits = digits + 1
            p = p + 1
        Loop
        If digits = 0 Or v > 65535 Then
            Exit Function
        End If
        version(i) = v
    Next

    If p <= Len(tag) Then
        ' it has "-rc/beta1.." suffix
        If Mid$(tag, p, 1) = "-" Then p = p + 1 ' skip
        n = Len(tag) - p + 1
        qualifier = Mid$(tag, p, TAGVER_MAX_QUALIF_LEN - 1)
        ' must terminate now
        If n > TAGVER_MAX_QUALIF_LEN - 1 Then
            Exit Function
        End If
    End If
    version_extractVersionNumberFromTag = True
End Function

' Obtain the versioned application-data / config writedir folder name
'  If on a tag, this is "Warzone 2100" / "warzone2100"
'  If not on a tag, this is "Warzone 2100 <branch>" / "warzone2100-<branch>"
'  If no branch is defined, this is "Warzone 2100 <VCS_EXTRA>" / "warzone2100-<VCS_EXTRA>"
Public Function version_getVersionedAppDirFolderName() As String
    Dim versionedWriteDirFolderName As String, versionedWriteDirFolderNameSuffix As String

    versionedWriteDirFolderName = "WarzoneVB"           ' (VB) Warzone: "Warzone 2100"

    If Len(VCS_TAG) > 0 Then
        ' release builds now use a stable config directory name, with no suffix
        versionedWriteDirFolderNameSuffix = ""
    ElseIf Len(VCS_BRANCH) > 0 Then
        versionedWriteDirFolderNameSuffix = VCS_BRANCH
    Else
        ' not a branch or a tag, so we are detached most likely: VCS_EXTRA without its spaces
        versionedWriteDirFolderNameSuffix = Replace(VCS_EXTRA, " ", "")
    End If

    If Len(versionedWriteDirFolderNameSuffix) > 0 Then
        versionedWriteDirFolderName = versionedWriteDirFolderName & " " & versionedWriteDirFolderNameSuffix
    End If
    version_getVersionedAppDirFolderName = versionedWriteDirFolderName
End Function

' Obtain the versioned application-data / config "mods" folder path
'  If on a tag, this is "mods/<tag version>/"
'  If not on a tag, this is "mods/" (as non-tagged builds have versioned parent config folders)
Public Function version_getVersionedModsFolderPath(Optional ByVal subFolders As String = "") As String
    Dim versionedModsFolderPath As String
    If Len(VCS_TAG) > 0 Then
        versionedModsFolderPath = "mods/" & VCS_TAG
    Else
        versionedModsFolderPath = "mods"
    End If

    If Len(subFolders) > 0 Then
        versionedModsFolderPath = versionedModsFolderPath & "/" & subFolders
    End If
    version_getVersionedModsFolderPath = versionedModsFolderPath
End Function

' Composes a nicely formatted version string: a tag (no hash), a branch (with the hash) or
' a detached state (with the hash).
Public Function version_getVersionString() As String
    If Len(VCS_TAG) > 0 Then
        version_getVersionString = VCS_TAG
    ElseIf Len(VCS_BRANCH) > 0 Then
        version_getVersionString = VCS_BRANCH & " " & VCS_SHORT_HASH
    Else
        ' not a branch or a tag, so we are detached most likely.
        version_getVersionString = VCS_EXTRA
    End If
End Function

Public Function version_getLatestTag() As String
    If Len(VCS_TAG) > 0 Then
        version_getLatestTag = VCS_TAG
        Exit Function
    End If
    ' this is non empty for sure!
    version_getLatestTag = VCS_MOST_RECENT_TAGGED_VERSION
End Function

' warzone2100@TAG_NAME (for tagged builds), warzone2100@FULL_COMMIT_HASH (for other builds)
Public Function version_getBuildIdentifierReleaseString() As String
    If Len(VCS_TAG) > 0 Then
        version_getBuildIdentifierReleaseString = "warzone2100@" & VCS_TAG
    Else
        version_getBuildIdentifierReleaseString = "warzone2100@" & VCS_FULL_HASH
    End If
End Function

' "release" (release tags), "preview" (tags with a qualifier, like "4.0.0-beta1"), "development" (other builds)
Public Function version_getBuildIdentifierReleaseEnvironment() As String
    Dim version() As Long, qualifier As String
    If Len(VCS_TAG) > 0 Then
        If version_extractVersionNumberFromTag(VCS_TAG, version, qualifier) And Len(qualifier) <> 0 Then
            version_getBuildIdentifierReleaseEnvironment = "preview"
        Else
            version_getBuildIdentifierReleaseEnvironment = "release"
        End If
    Else
        version_getBuildIdentifierReleaseEnvironment = "development"
    End If
End Function

Public Function version_getFormattedVersionString(Optional ByVal translated As Boolean = True) As String
    Dim wc_state As String, build_type As String
    ' Compose the working copy state string
    If VCS_WC_MODIFIED <> 0 Then wc_state = " (modified locally)"
    ' Compose the build type string: release (DEBUG undefined)
    build_type = ""

    ' Construct the version string: "Version: <version name/number>, <working copy state>, Built: <BUILD DATE><BUILD TYPE>"
    version_getFormattedVersionString = Left$(va("Version: %s,%s Built: %s%s", version_getVersionString(), wc_state, getCompileDate(), build_type), MAX_STR_LENGTH - 1)
End Function

' Returns a user-agent string, following typical formats & best-practices for HTTP(S) requests
Public Function version_getHTTPUserAgentString() As String
    Dim userAgentString As String, includeCompatible As Boolean
    userAgentString = "Warzone2100/"
    If Len(VCS_TAG) > 0 Then
        userAgentString = userAgentString & VCS_TAG
    Else
        ' use VCS_MOST_RECENT_TAGGED_VERSION
        userAgentString = userAgentString & VCS_MOST_RECENT_TAGGED_VERSION
        includeCompatible = True
    End If

    ' following best-practices, include a comment with details & descriptive url
    userAgentString = userAgentString & " ("
    If includeCompatible Then
        userAgentString = userAgentString & "compatible; "
    End If
    userAgentString = userAgentString & "+https://github.com/Warzone2100/warzone2100)"
    version_getHTTPUserAgentString = userAgentString
End Function
