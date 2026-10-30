VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "json_variant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/framework/wzconfig.h and lib/framework/wzconfig.cpp: class json_variant, a JSON value with
' conversions (WzConfig::value, json_getValue). Made by json_variant_new
' (its constructors). The conversions' bool *ok is an Optional Variant that
' receives the result (C: nullptr when it is not given).
Option Explicit

Private mObj As nlohmann_json

Private Sub Class_Initialize()
    Set mObj = json_value(Empty)                ' json_variant(): null
End Sub

' json_variant(const nlohmann::json& jObj): a copy
Friend Sub init(ByVal jObj As nlohmann_json)
    Set mObj = jObj.clone()
End Sub

Public Function jsonValue() As nlohmann_json
    Set jsonValue = mObj
End Function

Public Function toInt(Optional ok As Variant) As Long
    Dim result As String, value As Long, errText As String
    If mObj.is_number() Then
        toInt = json_variant_toType(Me, JVT_INT, 0, ok)
    ElseIf mObj.is_boolean() Then
        toInt = -CLng(json_variant_toType(Me, JVT_BOOL, False, ok))
    ElseIf mObj.is_string() Then
        result = json_variant_toType(Me, JVT_STRING, "", ok)
        If std_stoi(result, value, errText) Then
            toInt = value
        Else
            debug_ 604, LOG_WARNING, "toInt", "Failed to convert string '%s' to int because of error: %s", result, errText
            toInt = 0
        End If
    ElseIf mObj.is_null() Then
        toInt = 0
    Else
        toInt = 0
    End If
End Function

' unsigned int: the bits
Public Function toUInt(Optional ok As Variant) As Long
    Dim result As String, ulongValue As Long, errText As String
    If mObj.is_number() Then
        toUInt = json_variant_toType(Me, JVT_UINT, 0, ok)
    ElseIf mObj.is_boolean() Then
        toUInt = -CLng(json_variant_toType(Me, JVT_BOOL, False, ok))
    ElseIf mObj.is_string() Then
        result = json_variant_toType(Me, JVT_STRING, "", ok)
        If std_stoul(result, ulongValue, errText) Then
            ' (ulongValue > std::numeric_limits<unsigned int>::max() cannot be: unsigned long is 32 bits on Windows)
            toUInt = ulongValue
        Else
            debug_ 642, LOG_WARNING, "toUInt", "Failed to convert string '%s' to unsigned int because of error: %s", result, errText
            toUInt = 0
        End If
    ElseIf mObj.is_null() Then
        toUInt = 0
    Else
        toUInt = 0
    End If
End Function

' uint64_t, as a Double (exact to 2^53)
Public Function toUint64(Optional ok As Variant) As Double
    Dim result As String, ulonglongValue As Double, errText As String
    If mObj.is_number() Then
        toUint64 = json_variant_toType(Me, JVT_UINT64, 0, ok)
    ElseIf mObj.is_boolean() Then
        toUint64 = -CLng(json_variant_toType(Me, JVT_BOOL, False, ok))
    ElseIf mObj.is_string() Then
        result = json_variant_toType(Me, JVT_STRING, "", ok)
        If std_stoull(result, ulonglongValue, errText) Then
            toUint64 = ulonglongValue
        Else
            debug_ 682, LOG_WARNING, "toUint64", "Failed to convert string '%s' to uint64_t because of error: %s", result, errText
            toUint64 = 0
        End If
    ElseIf mObj.is_null() Then
        toUint64 = 0
    Else
        toUint64 = 0
    End If
End Function

Public Function toBool() As Boolean
    Dim result As String, d As Double
    If mObj.is_boolean() Then
        toBool = json_variant_toType(Me, JVT_BOOL, False)
    ElseIf mObj.is_number() Then
        d = json_variant_toType(Me, JVT_DOUBLE, 0)
        toBool = (d <> 0)
    ElseIf mObj.is_string() Then
        result = json_variant_toType(Me, JVT_STRING, "")
        result = LCase$(result)             ' std::tolower(c, std::locale::classic()) (only compared with "0" and "false" below)
        toBool = Len(result) > 0 And result <> "0" And result <> "false"
    Else
        debug_ 717, LOG_INFO, "toBool", "Expected a json value that can be converted to a boolean, instead received: %s", mObj.type_name()
        toBool = False
    End If
End Function

Public Function toDouble(Optional ok As Variant) As Double
    Dim result As String, value As Double, errText As String
    If mObj.is_number_float() Then
        toDouble = json_variant_toType(Me, JVT_DOUBLE, 0#, ok)
    ElseIf mObj.is_number_unsigned() Then
        toDouble = json_variant_toType(Me, JVT_NUMBER_UNSIGNED, 0, ok)
    ElseIf mObj.is_number_integer() Then
        toDouble = json_variant_toType(Me, JVT_NUMBER_INTEGER, 0, ok)
    ElseIf mObj.is_boolean() Then
        toDouble = -CLng(json_variant_toType(Me, JVT_BOOL, False, ok))
    ElseIf mObj.is_string() Then
        result = json_variant_toType(Me, JVT_STRING, "", ok)
        If std_stod(result, value, errText) Then
            toDouble = value
        Else
            debug_ 749, LOG_WARNING, "toDouble", "Failed to convert string '%s' to double because of error: %s", result, errText
            toDouble = 0
        End If
    ElseIf mObj.is_null() Then
        toDouble = 0
    Else
        toDouble = 0
    End If
End Function

Public Function toFloat(Optional ok As Variant) As Single
    Dim result As String, value As Single, errText As String
    If mObj.is_number_float() Then
        toFloat = json_variant_toType(Me, JVT_FLOAT, 0!, ok)
    ElseIf mObj.is_number_unsigned() Then
        toFloat = CSng(json_variant_toType(Me, JVT_NUMBER_UNSIGNED, 0, ok))
    ElseIf mObj.is_number_integer() Then
        toFloat = CSng(json_variant_toType(Me, JVT_NUMBER_INTEGER, 0, ok))
    ElseIf mObj.is_boolean() Then
        toFloat = -CLng(json_variant_toType(Me, JVT_BOOL, False, ok))
    ElseIf mObj.is_string() Then
        result = json_variant_toType(Me, JVT_STRING, "", ok)
        If std_stof(result, value, errText) Then
            toFloat = value
        Else
            debug_ 790, LOG_WARNING, "toFloat", "Failed to convert string '%s' to float because of error: %s", result, errText
            toFloat = 0
        End If
    ElseIf mObj.is_null() Then
        toFloat = 0
    Else
        toFloat = 0
    End If
End Function

Public Function toWzString(Optional ok As Variant) As String
    If mObj.is_string() Then
        toWzString = json_variant_toType(Me, JVT_WZSTRING, "", ok)
    ElseIf mObj.is_boolean() Then
        If json_variant_toType(Me, JVT_BOOL, False, ok) Then toWzString = "true" Else toWzString = "false"
    ElseIf mObj.is_number_unsigned() Then
        toWzString = dump_integer(json_variant_toType(Me, JVT_NUMBER_UNSIGNED, 0, ok))     ' WzString::number
    ElseIf mObj.is_number_integer() Then
        toWzString = dump_integer(json_variant_toType(Me, JVT_NUMBER_INTEGER, 0, ok))
    ElseIf mObj.is_number_float() Then
        toWzString = va("%.15g", CDbl(json_variant_toType(Me, JVT_NUMBER_FLOAT, 0#, ok))) ' WzString::number(double): precision digits10
    ElseIf mObj.is_null() Then
        toWzString = ""
    Else
        If Not IsMissing(ok) Then ok = False
        toWzString = ""
    End If
End Function

' C: std::vector<WzString> toWzStringList(): the strings in result(0 ..), and their count
Public Function toWzStringList(result() As String) As Long
    Dim selfAsString As String, str_ As String, ok As Variant, i As Long, n As Long, v As json_variant
    If Not mObj.is_array() Then
        ' attempt to convert the non-array to a string, and return as a vector of 1 entry (if the string is non-empty)
        selfAsString = toWzString()
        If Len(selfAsString) > 0 Then
            ReDim result(0 To 0)
            result(0) = selfAsString
            n = 1
        End If
        toWzStringList = n
        Exit Function
    End If

    ok = False
    For i = 0 To mObj.size() - 1
        ' Use json_variant to support conversion from other value types to string
        Set v = json_variant_new(mObj.value_at(i))
        str_ = v.toWzString(ok)
        If Not ok Then
            debug_ 867, LOG_WARNING, "toWzStringList", "Encountered a JSON value in the array that cannot be converted to a string (type: %s)", mObj.value_at(i).type_name()
            Exit For
        End If
        If n = 0 Then ReDim result(0 To mObj.size() - 1)
        result(n) = str_
        n = n + 1
    Next
    toWzStringList = n
End Function
