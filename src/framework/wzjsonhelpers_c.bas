Attribute VB_Name = "wzjsonhelpers_c"
' src/wzjsonhelpers.cpp (in WzFramework, see docs/CONVENTIONS.md):
' wzLoadJsonObjectFromFile. optional<nlohmann::json> is the value or Nothing.
Option Explicit

Public Function wzLoadJsonObjectFromFile(ByVal filePath As String, Optional ByVal quiet As Boolean = False) As nlohmann_json
    Dim data() As Byte, size As Long, mRoot As nlohmann_json

    If PHYSFS_exists(filePath) = 0 Then
        Exit Function
    End If
    If Not loadFileToBufferVector(filePath, data, size, False, True) Then
        Exit Function
    End If

    If size = 0 Then
        If Not quiet Then debug_ 42, LOG_ERROR, "wzLoadJsonObjectFromFile", "Empty file: %s", filePath
        Exit Function
    End If
    If data(size - 1) <> 0 Then
        ReDim Preserve data(0 To size)
        data(size) = 0 ' always ensure data is null-terminated
        size = size + 1
    End If

    ' parse JSON
    On Error GoTo failed
    Set mRoot = json_parse_bytes(data, 0, size - 1)
    On Error GoTo 0
    If mRoot.is_null() Then
        If Not quiet Then debug_ 65, LOG_ERROR, "wzLoadJsonObjectFromFile", "JSON document from %s is null", filePath
        Exit Function
    End If
    If Not mRoot.is_object() Then
        If Not quiet Then debug_ 70, LOG_ERROR, "wzLoadJsonObjectFromFile", "JSON document from %s is not an object. Read: \n%s", filePath, utf8_decode(data, 0, size - 1)
        Exit Function
    End If
    Erase data

    Set wzLoadJsonObjectFromFile = mRoot
    Exit Function

failed:
    If Err.Source = "nlohmann::json" Then
        If Not quiet Then debug_ 56, LOG_ERROR, "wzLoadJsonObjectFromFile", "JSON document from %s is invalid: %s", filePath, Err.Description
    Else
        If Not quiet Then debug_ 60, LOG_ERROR, "wzLoadJsonObjectFromFile", "Unexpected exception parsing JSON %s", filePath
    End If
End Function
