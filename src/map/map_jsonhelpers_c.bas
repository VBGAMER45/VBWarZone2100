Attribute VB_Name = "map_jsonhelpers_c"
' lib/wzmaplib/src/map_jsonhelpers.cpp, map_jsonhelpers.h: loading and saving
' a map's JSON files through an IOProvider, with WzFramework's
' nlohmann_json. optional<nlohmann::json> is the value or Nothing.
Option Explicit

Private Const MaxJsonFileSize As Long = 20971520    ' 20 * 1024 * 1024

' try { mRoot = nlohmann::json::parse(data.begin(), data.end() - 1); }: false
' with e.what() when it throws (isStdException false: catch (...))
Private Function TryParse(data() As Byte, ByVal size As Long, mRoot As nlohmann_json, what As String, isStdException As Boolean) As Boolean
    On Error GoTo catch_
    Set mRoot = json_parse_bytes(data, 0, size)
    TryParse = True
    Exit Function
catch_:
    isStdException = (Err.Source = "nlohmann::json")
    what = Err.Description
End Function

Public Function loadJsonObjectFromFile(ByVal filename As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As nlohmann_json
    Dim path As String, data() As Byte, dataSize As Long, loadFileResult As LoadFullFileResult
    Dim mRoot As nlohmann_json, what As String, isStdException As Boolean, nul As Long
    path = filename
    loadFileResult = mapIO.loadFullFile(filename, data, dataSize, MaxJsonFileSize, True)
    Select Case loadFileResult
    Case LoadFullFileResult_SUCCESS
    Case LoadFullFileResult_FAILURE_OPEN
        Exit Function
    Case LoadFullFileResult_FAILURE_READ
        ' file exists but can't be read
        printLog_ pCustomLogger, 43, LogLevel_Error, "loadJsonObjectFromFile", "Failed to read file: %s", filename
        Exit Function
    Case LoadFullFileResult_FAILURE_EXCEEDS_MAXFILESIZE
        printLog_ pCustomLogger, 46, LogLevel_Error, "loadJsonObjectFromFile", "File is too large: %s", filename
        Exit Function
    End Select
    If dataSize = 0 Then
        printLog_ pCustomLogger, 51, LogLevel_Error, "loadJsonObjectFromFile", "Empty file: %s", path
        Exit Function
    End If
    If data(dataSize - 1) <> 0 Then
        ReDim Preserve data(0 To dataSize)
        data(dataSize) = 0 ' always ensure data is null-terminated
        dataSize = dataSize + 1
    End If

    ' parse JSON
    If Not TryParse(data, dataSize - 1, mRoot, what, isStdException) Then
        If isStdException Then
            printLog_ pCustomLogger, 65, LogLevel_Error, "loadJsonObjectFromFile", "JSON document from %s is invalid: %s", path, what
        Else
            printLog_ pCustomLogger, 69, LogLevel_Error, "loadJsonObjectFromFile", "Unexpected exception parsing JSON %s", path
        End If
        Exit Function
    End If

    If mRoot.is_null() Then
        printLog_ pCustomLogger, 74, LogLevel_Error, "loadJsonObjectFromFile", "JSON document from %s is null", path
        Exit Function
    End If
    If Not mRoot.is_object() Then
        nul = 0                                 ' data.data() as a C string: up to its first NUL
        Do While data(nul) <> 0
            nul = nul + 1
        Loop
        printLog_ pCustomLogger, 79, LogLevel_Error, "loadJsonObjectFromFile", "JSON document from %s is not an object. Read: \n%s", path, utf8_decode(data, 0, nul)
        Exit Function
    End If

    Erase data
    Set loadJsonObjectFromFile = mRoot
End Function

Public Function saveOrderedJsonObjectToFile(ByVal jsonObj As nlohmann_json, ByVal filename As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As Boolean
    Dim jsonStr As String, b() As Byte, n As Long
    jsonStr = jsonObj.dump(-1, " ", False, error_handler_t_ignore)
    ' (the SIZE_MAX > UINT32_MAX check: not on 32-bit Windows)
    n = utf8_encode(jsonStr, b)
    saveOrderedJsonObjectToFile = mapIO.writeFullFile(filename, b, n)
End Function
