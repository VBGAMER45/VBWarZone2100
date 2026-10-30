Attribute VB_Name = "campaigninfo_c"
' src/campaigninfo.cpp, campaigninfo.h: the campaign's number, name and tweak options. The tweak options
' (std::unordered_map<std::string, nlohmann::json>) are a json object of them; getCamTweakOptions returns a
' copy of it. readCampaignFiles' std::vector<CAMPAIGN_FILE> is an array and its count.
' campaigninfo_c_init sets the statics' initialisers (exemain_vb calls it).
Option Explicit

Public Type CAMPAIGN_FILE
    name_ As String                     ' WzString name
    level As String
    video As String
    captions As String
End Type

' Which campaign are we dealing with?
Private camNumber As Long               ' static uint32_t camNumber = 1;
Private campaignName As String          ' static optional<std::string> campaignName = nullopt;
Private hasCampaignName As Boolean

' cam tweak options
Private camTweakOptions As nlohmann_json   ' std::unordered_map<std::string, nlohmann::json> (Nothing: empty)

' readCampaignFiles' static std::vector<CAMPAIGN_FILE> result
Private campaignFilesResult() As CAMPAIGN_FILE
Private campaignFilesResult_size As Long

' (VB) the statics' initialisers
Public Sub campaigninfo_c_init()
    camNumber = 1
End Sub

Public Sub setCampaignNumber(ByVal number As Long)
    If Not (ULT(number, 4)) Then ASSERT_FAILURE "number < 4", "src/campaigninfo.cpp:45", "setCampaignNumber", "Campaign Number too high!"
    camNumber = number

    If Not bMultiPlayer And Not bInTutorial Then
        If Not hasCampaignName Then
            campaignName = getCampaignNameFromCampaignNumber(number)
            hasCampaignName = True
        End If
    Else
        clearCampaignName
    End If
End Sub

Public Function getCampaignNumber() As Long
    getCampaignNumber = camNumber
End Function

Public Sub setCampaignName(ByVal newCampaignName As String)
    campaignName = newCampaignName
    hasCampaignName = True
End Sub

Public Sub clearCampaignName()
    campaignName = vbNullString
    hasCampaignName = False
End Sub

' campaignName.value_or(std::string())
Public Function getCampaignName() As String
    If hasCampaignName Then getCampaignName = campaignName Else getCampaignName = ""
End Function

' Fallback for old campaign saves *ONLY*
Public Function getCampaignNameFromCampaignNumber(ByVal campaignNum As Long) As String
    Dim list() As CAMPAIGN_FILE, n As Long
    n = readCampaignFiles(list)

    If campaignNum > 0 And campaignNum <= n Then
        getCampaignNameFromCampaignNumber = list(campaignNum - 1).name_
        Exit Function
    End If
    getCampaignNameFromCampaignNumber = ""
End Function

' std::vector<CAMPAIGN_FILE> readCampaignFiles(): (VB) the files in list(0 .. n - 1), n returned
Public Function readCampaignFiles(list() As CAMPAIGN_FILE) As Long
    Dim files() As String, nFiles As Long, k As Long, c As CAMPAIGN_FILE, filename As String
    Dim jsonObj As nlohmann_json, ok As Boolean, errDescription As String, empty_ As CAMPAIGN_FILE

    If campaignFilesResult_size = 0 Then
        ' WZ_PHYSFS_enumerateFiles("campaigns", [&](const char *i) -> bool {...}): the files, then the lambda's body
        If WZ_PHYSFS_enumerateFiles("campaigns", files, nFiles) Then
            For k = 0 To nFiles - 1
                c = empty_
                filename = "campaigns/"
                filename = filename & files(k)
                If Right$(filename, 5) <> ".json" Then
                    GoTo continue_                  ' return true; // continue;
                End If

                Set jsonObj = wzLoadJsonObjectFromFile(filename)
                If jsonObj Is Nothing Then
                    GoTo continue_                  ' return true; // continue;
                End If

                ' try { c = jsonObj.value().get<CAMPAIGN_FILE>(); }
                ok = CAMPAIGN_FILE_from_json_try(jsonObj, c, errDescription)
                If Not ok Then
                    ' Failed to parse the JSON
                    debug_ 121, LOG_ERROR, "readCampaignFiles", "JSON document from %s is invalid: %s", filename, errDescription
                    GoTo continue_                  ' return true; // continue;
                End If

                If campaignFilesResult_size = 0 Then
                    ReDim campaignFilesResult(0 To 0)
                Else
                    ReDim Preserve campaignFilesResult(0 To campaignFilesResult_size)
                End If
                campaignFilesResult(campaignFilesResult_size) = c   ' result.push_back(c);
                campaignFilesResult_size = campaignFilesResult_size + 1
continue_:
            Next
        End If
    End If

    If campaignFilesResult_size > 0 Then
        ReDim list(0 To campaignFilesResult_size - 1)
        For k = 0 To campaignFilesResult_size - 1
            list(k) = campaignFilesResult(k)
        Next
    Else
        Erase list
    End If
    readCampaignFiles = campaignFilesResult_size
End Function

' (VB) the try block around jsonObj.value().get<CAMPAIGN_FILE>(): False and the exception's what() when it throws
Private Function CAMPAIGN_FILE_from_json_try(ByVal j As nlohmann_json, v As CAMPAIGN_FILE, errDescription As String) As Boolean
    On Error GoTo threw
    CAMPAIGN_FILE_from_json j, v
    CAMPAIGN_FILE_from_json_try = True
    Exit Function
threw:
    errDescription = Err.Description
    CAMPAIGN_FILE_from_json_try = False
End Function

' void from_json(const nlohmann::json& j, CAMPAIGN_FILE& v)
Public Sub CAMPAIGN_FILE_from_json(ByVal j As nlohmann_json, v As CAMPAIGN_FILE)
    Dim it As nlohmann_json
    ' required properties
    v.name_ = j.at("name").get_string()
    v.level = j.at("level").get_string()

    ' optional properties
    Set it = j.find("video")
    If Not it Is Nothing Then
        v.video = it.get_string()
    End If
    Set it = j.find("captions")
    If Not it Is Nothing Then
        v.captions = it.get_string()
    End If
End Sub

' void setCamTweakOptions(std::unordered_map<std::string, nlohmann::json> options): (VB) a json object of them, copied
Public Sub setCamTweakOptions(ByVal options As nlohmann_json)
    Set camTweakOptions = options.clone()
End Sub

Public Sub clearCamTweakOptions()
    Set camTweakOptions = Nothing
End Sub

' const std::unordered_map<std::string, nlohmann::json>& getCamTweakOptions(): (VB) a json object of them (a copy)
Public Function getCamTweakOptions() As nlohmann_json
    If camTweakOptions Is Nothing Then
        Set getCamTweakOptions = json_object()
    Else
        Set getCamTweakOptions = camTweakOptions.clone()
    End If
End Function

Public Function getCamTweakOptionsValue(ByVal identifier As String, ByVal defaultValue As nlohmann_json) As nlohmann_json
    Dim it As nlohmann_json
    If Not camTweakOptions Is Nothing Then Set it = camTweakOptions.find(identifier)
    If it Is Nothing Then
        Set getCamTweakOptionsValue = defaultValue
        Exit Function
    End If
    Set getCamTweakOptionsValue = it.clone()
End Function

Public Function getCamTweakOption_AutosavesOnly() As Boolean
    Dim val As nlohmann_json
    Set val = getCamTweakOptionsValue("autosavesOnly", json_value(False))
    If Not val.is_boolean() Then
        getCamTweakOption_AutosavesOnly = False
        Exit Function
    End If
    getCamTweakOption_AutosavesOnly = val.get_bool()
End Function

Public Function getCamTweakOption_PS1Modifiers() As Boolean
    Dim val As nlohmann_json
    Set val = getCamTweakOptionsValue("ps1Modifiers", json_value(False))
    If Not val.is_boolean() Then
        getCamTweakOption_PS1Modifiers = False
        Exit Function
    End If
    getCamTweakOption_PS1Modifiers = val.get_bool()
End Function

Public Function getCamTweakOption_FastExp() As Boolean
    Dim val As nlohmann_json
    Set val = getCamTweakOptionsValue("fastExp", json_value(False))
    If Not val.is_boolean() Then
        getCamTweakOption_FastExp = False
        Exit Function
    End If
    getCamTweakOption_FastExp = val.get_bool()
End Function
