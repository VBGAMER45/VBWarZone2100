Attribute VB_Name = "modinfo_c"
' src/modinfo.cpp, modinfo.h, whole; with src/wzjsonlocalizedstring.cpp's from_json and getLocalizedString
' (WzJsonLocalizedString_from_json, WzJsonLocalizedString_getLocalizedString: only modinfo.cpp reads such
' strings in the port). modinfo.h's classes: WZModInfo and its only subclass WzCampaignModInfo are one VB class,
' WzCampaignModInfo.cls (WZModInfo's fields first; WZModInfo::getModCompatibility is its method, which calls
' isWZModVersionCompatible here); struct WzCampaignTweakOption is WzCampaignTweakOption.cls; modinfo.h's enums
' are campaign_title_c's (WzModType, WzModCompatibilityResult, WzCampaignUniverse, WzCampaignTweakOption_Type).
' A from_json(j, v) overload is <Type>_from_json(j) returning v; they throw as the C++ does (nlohmann's
' type_error::create(302, ...) is VB error vbObjectError + 302 from "nlohmann::json" with its what() text:
' "[json.exception.type_error.302] ..." (JSON_DIAGNOSTICS is 0); parse_error::create(302, 0, ...) likewise,
' "[json.exception.parse_error.302] parse error: ..."). optional<WzCampaignModInfo> is the object or Nothing.
' (VB) A WzJsonLocalizedString field holds its getLocalizedString(): from_json reads the mappings and the
' string is picked at once, with the language of that moment (translations are not ported: dgettext gives its
' input back). A TagVer is version(0 .. 2) (version_c).
Option Explicit

'--------------------------------------------------------------------------
' (VB) TagVer's comparisons (version.h)

' bool TagVer::operator < (const TagVer &other) const: std::lexicographical_compare of the three numbers
Private Function TagVer_less(a() As Long, b() As Long) As Boolean
    Dim i As Long
    For i = 0 To 2
        If a(i) < b(i) Then
            TagVer_less = True
            Exit Function
        ElseIf b(i) < a(i) Then
            Exit Function
        End If
    Next
End Function

' bool TagVer::operator == (const TagVer &other) const: memcmp of the three numbers
Private Function TagVer_equal(a() As Long, b() As Long) As Boolean
    TagVer_equal = (a(0) = b(0)) And (a(1) = b(1)) And (a(2) = b(2))
End Function

'--------------------------------------------------------------------------

' static WzModCompatibilityResult isWZModVersionCompatible(const WZModInfo& modInfo): (VB) Public, for
' WzCampaignModInfo.getModCompatibility
Public Function isWZModVersionCompatible(ByVal modInfo As WzCampaignModInfo) As WzModCompatibilityResult
    Dim tagHasValue As Boolean, latestTaggedVersion() As Long, tagQualifier As String
    Dim minVerHas As Boolean, minVerResult() As Long, minQualifier As String
    Dim maxVerHas As Boolean, maxVerTestedResult() As Long, maxQualifier As String
    Dim buildVersionIdentifier As String, isOnReleaseBuild As Boolean

    tagHasValue = version_extractVersionNumberFromTag(version_getLatestTag(), latestTaggedVersion, tagQualifier)
    If Not (tagHasValue) Then ASSERT_FAILURE "tagResult.has_value()", "src/modinfo.cpp:34", "isWZModVersionCompatible", "No extractable latest tag?? - Please try re-downloading the latest official source bundle"
    If Not tagHasValue Then ReDim latestTaggedVersion(0 To 2)  ' const TagVer latestTaggedVersion = tagResult.value_or(TagVer());

    ' minVersionSupported incompatibility check
    minVerHas = version_extractVersionNumberFromTag(modInfo.minVersionSupported, minVerResult, minQualifier)
    If minVerHas Then
        If TagVer_less(latestTaggedVersion, minVerResult) Then
            isWZModVersionCompatible = WzModCompatibilityResult_NOT_COMPATIBLE
            Exit Function
        End If
        ' otherwise, fall-through and do additional checks
    Else
        ' input minVersionSupported didn't seem to be a version string...
        isWZModVersionCompatible = WzModCompatibilityResult_NOT_COMPATIBLE
        Exit Function
    End If

    ' additional checks
    maxVerHas = version_extractVersionNumberFromTag(modInfo.maxVersionTested, maxVerTestedResult, maxQualifier)
    buildVersionIdentifier = version_getBuildIdentifierReleaseEnvironment()
    isOnReleaseBuild = (buildVersionIdentifier = "release" Or buildVersionIdentifier = "preview")
    If isOnReleaseBuild Then
        If TagVer_equal(latestTaggedVersion, minVerResult) Then
            isWZModVersionCompatible = WzModCompatibilityResult_OK
            Exit Function
        End If
    End If

    ' latestTaggedVersion is >= minVersionSupported - check maxVersionTested
    If Not maxVerHas Then
        ' no maxVersionTested? - _might_ be compatible
        isWZModVersionCompatible = WzModCompatibilityResult_MAYBE_COMPATIBLE
        Exit Function
    End If

    ' check maxTestedVersion
    If TagVer_less(latestTaggedVersion, maxVerTestedResult) Then
        isWZModVersionCompatible = WzModCompatibilityResult_OK
    ElseIf TagVer_equal(latestTaggedVersion, maxVerTestedResult) Then
        ' if latestTaggedVersion == maxVersionTested on a RELEASE build, then mark as compatible
        ' if not on a release build, it's likely that changes have been made since the latestTaggedVersion, so mark as maybe_compatible
        If isOnReleaseBuild Then
            isWZModVersionCompatible = WzModCompatibilityResult_OK
        Else
            isWZModVersionCompatible = WzModCompatibilityResult_MAYBE_COMPATIBLE
        End If
    Else
        isWZModVersionCompatible = WzModCompatibilityResult_MAYBE_COMPATIBLE
    End If
End Function

' void from_json(const nlohmann::json& j, WzModType& v)
Private Function WzModType_from_json(ByVal j As nlohmann_json) As WzModType
    Dim str_ As String
    str_ = j.get_string()
    If str_ = "alternateCampaign" Then
        WzModType_from_json = WzModType_AlternateCampaign
        Exit Function
    ElseIf str_ = "campaignBalanceMod" Then
        WzModType_from_json = WzModType_CampaignBalanceMod
        Exit Function
    End If
    Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type value is unknown: """ & str_ & """"
End Function

Private Sub loadBaseModInfo(ByVal j As nlohmann_json, ByVal v As WzCampaignModInfo)
    Dim it As nlohmann_json
    If Not j.is_object() Then
        Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type must be an object, but is " & j.type_name()
    End If

    v.name_ = j.at("name").get_string()
    v.author = j.at("author").get_string()
    v.type_ = WzModType_from_json(j.at("type"))
    v.description = WzJsonLocalizedString_from_json(j.at("description"))
    v.license = j.at("license").get_string()
    Set it = j.find("version")
    If Not it Is Nothing Then
        v.version = it.get_string()
        v.hasVersion = True
    End If
    Set it = j.find("updatesUrl")
    If Not it Is Nothing Then
        v.updatesUrl = it.get_string()
        v.hasUpdatesUrl = True
    End If
    v.minVersionSupported = j.at("minVersionSupported").get_string()
    v.maxVersionTested = j.at("maxVersionTested").get_string()
End Sub

' void from_json(const nlohmann::json& j, DIFFICULTY_LEVEL& v)
Private Function DIFFICULTY_LEVEL_from_json(ByVal j As nlohmann_json) As DIFFICULTY_LEVEL
    Dim str_ As String
    str_ = j.get_string()
    If str_ = "super easy" Then
        DIFFICULTY_LEVEL_from_json = DL_SUPER_EASY
        Exit Function
    ElseIf str_ = "easy" Then
        DIFFICULTY_LEVEL_from_json = DL_EASY
        Exit Function
    ElseIf str_ = "normal" Then
        DIFFICULTY_LEVEL_from_json = DL_NORMAL
        Exit Function
    ElseIf str_ = "hard" Then
        DIFFICULTY_LEVEL_from_json = DL_HARD
        Exit Function
    ElseIf str_ = "insane" Then
        DIFFICULTY_LEVEL_from_json = DL_INSANE
        Exit Function
    End If
    Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] difficultyLevel value is unknown: """ & str_ & """"
End Function

' void from_json(const nlohmann::json& j, WzCampaignUniverse& v)
Private Function WzCampaignUniverse_from_json(ByVal j As nlohmann_json) As WzCampaignUniverse
    Dim str_ As String
    str_ = j.get_string()
    If str_ = "wz2100" Then
        WzCampaignUniverse_from_json = WzCampaignUniverse_WZ2100
        Exit Function
    ElseIf str_ = "wz2100extended" Then
        WzCampaignUniverse_from_json = WzCampaignUniverse_WZ2100Extended
        Exit Function
    ElseIf str_ = "unique" Then
        WzCampaignUniverse_from_json = WzCampaignUniverse_Unique
        Exit Function
    End If
    Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] universe value is unknown: """ & str_ & """"
End Function

' void from_json(const nlohmann::json& j, WzCampaignTweakOption::Type& v)
Private Function WzCampaignTweakOption_Type_from_json(ByVal j As nlohmann_json) As WzCampaignTweakOption_Type
    Dim str_ As String
    str_ = j.get_string()
    If str_ = "bool" Then
        WzCampaignTweakOption_Type_from_json = WzCampaignTweakOption_Type_Bool
        Exit Function
    End If
    Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type value is unknown: """ & str_ & """"
End Function

Private Function builtinTweakOptionFromJSON(ByVal j As nlohmann_json) As WzCampaignTweakOption
    Dim v As WzCampaignTweakOption, it As nlohmann_json
    Set v = New WzCampaignTweakOption
    v.uniqueIdentifier = j.at("id").get_string()
    v.type_ = WzCampaignTweakOption_Type_Bool ' always a bool
    v.enabled = j.at("default").get_bool()
    Set it = j.find("userEditable")
    If Not it Is Nothing Then
        v.userEditable = j.at("userEditable").get_bool()
    Else
        v.userEditable = True
    End If
    Set builtinTweakOptionFromJSON = v
End Function

Private Function customTweakOptionFromJSON(ByVal j As nlohmann_json) As WzCampaignTweakOption
    Dim v As WzCampaignTweakOption, it As nlohmann_json
    Set v = New WzCampaignTweakOption
    v.uniqueIdentifier = j.at("id").get_string()
    v.type_ = WzCampaignTweakOption_Type_from_json(j.at("type"))
    v.enabled = j.at("default").get_bool()
    Set it = j.find("userEditable")
    If Not it Is Nothing Then
        v.userEditable = j.at("userEditable").get_bool()
    Else
        v.userEditable = True
    End If
    v.displayName = WzJsonLocalizedString_from_json(j.at("displayName"))
    v.description = WzJsonLocalizedString_from_json(j.at("description"))
    Set customTweakOptionFromJSON = v
End Function

' (VB) the try around jsonObj.value().get<CAMPAIGN_FILE>(): False and the exception's what() when it throws
Private Function CAMPAIGN_FILE_get_try(ByVal j As nlohmann_json, c As CAMPAIGN_FILE, what As String) As Boolean
    On Error GoTo threw
    CAMPAIGN_FILE_from_json j, c
    CAMPAIGN_FILE_get_try = True
    Exit Function
threw:
    what = Err.Description
    CAMPAIGN_FILE_get_try = False
End Function

Private Function loadCampaignModInfo(ByVal j As nlohmann_json, ByVal baseDir As String) As WzCampaignModInfo
    Dim v As WzCampaignModInfo, campaigns_ As nlohmann_json, campaignJSONList() As String, nList As Long, k As Long
    Dim jsonName As String, fullJSONPath As String, jsonObj As nlohmann_json, c As CAMPAIGN_FILE, blank As CAMPAIGN_FILE
    Dim campaignFiles() As CAMPAIGN_FILE, nFiles As Long, what As String, difficultyLevels As nlohmann_json, i As Long
    Dim lvl As DIFFICULTY_LEVEL, it As nlohmann_json, tweakOptions As nlohmann_json
    Set v = New WzCampaignModInfo

    If Not j.is_object() Then
        Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type must be an object, but is " & j.type_name()
    End If
    loadBaseModInfo j, v

    If v.type_ = WzModType_AlternateCampaign Then
        ' campaigns
        Set campaigns_ = j.at("campaigns")
        If Not campaigns_.is_array() Then
            Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] campaigns type must be an array, but is " & campaigns_.type_name()
        End If
        ' campaignJSONList = campaigns.get<std::vector<std::string>>();
        nList = campaigns_.size()
        If nList > 0 Then ReDim campaignJSONList(0 To nList - 1)
        For k = 0 To nList - 1
            campaignJSONList(k) = campaigns_.at_idx(k).get_string()
        Next
        For k = 0 To nList - 1
            jsonName = campaignJSONList(k)
            fullJSONPath = baseDir & "/campaigns/" & jsonName
            Set jsonObj = wzLoadJsonObjectFromFile(fullJSONPath)
            If jsonObj Is Nothing Then
                debug_ 260, LOG_ERROR, "loadCampaignModInfo", "Failed to find: campaigns/%s", jsonName
                GoTo continue_
            End If

            c = blank
            If Not CAMPAIGN_FILE_get_try(jsonObj, c, what) Then
                ' Failed to parse the JSON
                debug_ 270, LOG_ERROR, "loadCampaignModInfo", "JSON document from %s is invalid: %s", jsonName, what
                GoTo continue_
            End If

            ' v.campaignFiles.push_back(c);
            ReDim Preserve campaignFiles(0 To nFiles)
            campaignFiles(nFiles) = c
            nFiles = nFiles + 1
continue_:
        Next
        v.setCampaignFiles campaignFiles, nFiles
        If v.campaignFiles_size = 0 Then
            Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.parse_error.302] parse error: ""campaigns"" must list one or more valid campaign JSON files"
        End If
    End If

    ' difficultyLevels
    Set difficultyLevels = j.at("difficultyLevels")
    If difficultyLevels.is_array() Then
        For i = 0 To difficultyLevels.size() - 1
            lvl = DIFFICULTY_LEVEL_from_json(difficultyLevels.at_idx(i))
            If Not v.hasDifficultyLevel(lvl) Then           ' v.difficultyLevels.insert(...)
                v.difficultyLevels.Add lvl, "d" & lvl
            End If
        Next
    ElseIf difficultyLevels.is_string() Then
        If difficultyLevels.get_string() = "default" Then
            Set v.difficultyLevels = New Collection         ' v.difficultyLevels.clear();
        Else
            Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] difficultyLevels is a string, but is not ""default"""
        End If
    Else
        Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] difficultyLevels is not an array or string, but is " & difficultyLevels.type_name()
    End If

    ' universe
    v.universe = WzCampaignUniverse_from_json(j.at("universe"))

    ' camTweakOptions (optional)
    Set v.camTweakOptions = New Collection                  ' v.camTweakOptions.clear();
    Set it = j.find("camTweakOptions")
    If Not it Is Nothing Then
        Set tweakOptions = it
        If Not tweakOptions.is_array() Then
            Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] camTweakOptions type must be an array, but is " & tweakOptions.type_name()
        End If

        ' (v.camTweakOptions.reserve(tweakOptions.size());)
        For i = 0 To tweakOptions.size() - 1
            v.camTweakOptions.Add builtinTweakOptionFromJSON(tweakOptions.at_idx(i))
        Next
    End If

    ' customTweakOptions (optional)
    Set v.customTweakOptions = New Collection               ' v.customTweakOptions.clear();
    Set it = j.find("customTweakOptions")
    If Not it Is Nothing Then
        Set tweakOptions = it
        If Not tweakOptions.is_array() Then
            Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] customTweakOptions type must be an array, but is " & tweakOptions.type_name()
        End If

        ' (v.customTweakOptions.reserve(tweakOptions.size());)
        For i = 0 To tweakOptions.size() - 1
            v.customTweakOptions.Add customTweakOptionFromJSON(tweakOptions.at_idx(i))
        Next
    End If

    Set loadCampaignModInfo = v
End Function

' (VB) the try around loadCampaignModInfo: False and the exception's what() when it throws
Private Function loadCampaignModInfo_try(ByVal j As nlohmann_json, ByVal baseDir As String, result As WzCampaignModInfo, what As String) As Boolean
    On Error GoTo threw
    Set result = loadCampaignModInfo(j, baseDir)
    loadCampaignModInfo_try = True
    Exit Function
threw:
    what = Err.Description
    loadCampaignModInfo_try = False
End Function

' optional<WzCampaignModInfo> loadCampaignModInfoFromFile(const std::string& filePath, const std::string& baseDir):
' (VB) Nothing for nullopt
Public Function loadCampaignModInfoFromFile(ByVal filePath As String, ByVal baseDir As String) As WzCampaignModInfo
    Dim jsonObj As nlohmann_json, result As WzCampaignModInfo, what As String
    Set jsonObj = wzLoadJsonObjectFromFile(filePath)
    If jsonObj Is Nothing Then
        Exit Function                                   ' return nullopt;
    End If

    ' from_json
    If Not loadCampaignModInfo_try(jsonObj, baseDir, result, what) Then
        ' Failed to parse the JSON
        debug_ 364, LOG_ERROR, "loadCampaignModInfoFromFile", "JSON document from %s is invalid: %s", filePath, what
        Exit Function                                   ' return nullopt;
    End If
    Set loadCampaignModInfoFromFile = result
End Function

'--------------------------------------------------------------------------
' src/wzjsonlocalizedstring.cpp

' (VB) WzJsonLocalizedString::mappings' find_if (findMapping): the index, n for mappings.end()
Private Function WzJsonLocalizedString_findMapping(codes() As String, ByVal n As Long, ByVal languageCode As String) As Long
    Dim i As Long
    If Len(languageCode) = 0 Then
        WzJsonLocalizedString_findMapping = n
        Exit Function
    End If
    For i = 0 To n - 1
        If codes(i) = languageCode Then
            WzJsonLocalizedString_findMapping = i
            Exit Function
        End If
    Next
    WzJsonLocalizedString_findMapping = n
End Function

' const char* WzJsonLocalizedString::getLocalizedString() const (getLocalizedString(PACKAGE)): (VB) the mappings
' codes(i) -> strs(i), i < n
Private Function WzJsonLocalizedString_getLocalizedString(codes() As String, strs() As String, ByVal n As Long) As String
    Dim currentLanguage As String, it As Long

    currentLanguage = getLanguage()                     ' (pLangCode) ? pLangCode : std::string()

    it = WzJsonLocalizedString_findMapping(codes, n, currentLanguage)
    If it = n Then
        ' get "en" value
        it = WzJsonLocalizedString_findMapping(codes, n, "en")
        If it <> n Then
            ' check gettext message catalog for a built-in translation of the "en" value
            ' if there is no translation, this just returns the input string (the "en" value)
            If Len(strs(it)) <> 0 Then
                WzJsonLocalizedString_getLocalizedString = strs(it)     ' dgettext(domainname, it->str.c_str()): not ported
                Exit Function
            End If
        Else
            ' just use the first entry, if available
            it = 0
        End If
    End If

    WzJsonLocalizedString_getLocalizedString = strs(it)
End Function

' void from_json(const nlohmann::json& j, WzJsonLocalizedString& v): (VB) then v.getLocalizedString()
Public Function WzJsonLocalizedString_from_json(ByVal j As nlohmann_json) As String
    Dim codes() As String, strs() As String, n As Long, i As Long
    If j.is_object() Then
        If j.size() > 0 Then
            ReDim codes(0 To j.size() - 1)
            ReDim strs(0 To j.size() - 1)
        End If
        For i = 0 To j.size() - 1                       ' for (auto it : j.items())
            codes(n) = j.key_at(i)
            strs(n) = j.value_at(i).get_string()
            n = n + 1
        Next
        If n = 0 Then
            Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type is an empty object"
        End If
    ElseIf j.is_string() Then
        ReDim codes(0 To 0)
        ReDim strs(0 To 0)
        codes(0) = "en"
        strs(0) = j.get_string()
        n = 1
    Else
        Err.Raise vbObjectError + 302, "nlohmann::json", "[json.exception.type_error.302] type must be an object or string, but is " & j.type_name()
    End If
    WzJsonLocalizedString_from_json = WzJsonLocalizedString_getLocalizedString(codes, strs, n)
End Function
