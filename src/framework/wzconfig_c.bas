Attribute VB_Name = "wzconfig_c"
' lib/framework/wzconfig.cpp: what is not a member of WzConfig or
' json_variant (WzConfig.cls, json_variant.cls): jsonMerge,
' json_variant_toType, json_getValue; WzConfig_new and json_variant_new
' (the constructors). json_variant_toType's TYPE is a JVT_ code; the
' warning names it with typeid(TYPE).name() as Microsoft's compiler does.
Option Explicit

' json_variant_toType's TYPE
Public Const JVT_INT As Long = 0
Public Const JVT_UINT As Long = 1
Public Const JVT_UINT64 As Long = 2
Public Const JVT_BOOL As Long = 3
Public Const JVT_DOUBLE As Long = 4
Public Const JVT_FLOAT As Long = 5
Public Const JVT_STRING As Long = 6
Public Const JVT_WZSTRING As Long = 7
Public Const JVT_NUMBER_UNSIGNED As Long = 8
Public Const JVT_NUMBER_INTEGER As Long = 9
Public Const JVT_NUMBER_FLOAT As Long = 10

' WzConfig(name, warning)
Public Function WzConfig_new(ByVal name_ As String, ByVal warning As WzConfig_warning) As WzConfig
    Dim c As WzConfig
    Set c = New WzConfig
    c.construct name_, warning
    Set WzConfig_new = c
End Function

' json_variant(int), (const char *), (std::string), (WzString), (const nlohmann::json&); () is Empty
Public Function json_variant_new(Optional ByVal v As Variant) As json_variant
    Dim jv As json_variant
    Set jv = New json_variant
    If IsMissing(v) Then
        ' json_variant(): null
    ElseIf IsObject(v) Then
        jv.init v
    Else
        jv.init json_value(v)
    End If
    Set json_variant_new = jv
End Function

' C: static; Public for WzConfig's constructor (WzConfig.cls)
Public Function jsonMerge(ByVal original As nlohmann_json, ByVal override_ As nlohmann_json) As nlohmann_json
    Dim i As Long, key As String, it_original As nlohmann_json, v As nlohmann_json
    Set original = original.clone()             ' (by value)
    For i = 0 To override_.size() - 1
        key = override_.key_at(i)
        Set v = override_.value_at(i)
        Set it_original = original.find(key)
        If v.is_object() And Not it_original Is Nothing Then
            original.set_item_ref key, jsonMerge(it_original, v)
        ElseIf v.is_null() Then
            original.erase_ key
        Else
            original.set_item key, v
        End If
    Next
    Set jsonMerge = original
End Function

' Microsoft's typeid(TYPE).name()
Private Function TypeIdName(ByVal typ As Long) As String
    Select Case typ
    Case JVT_INT: TypeIdName = "int"
    Case JVT_UINT: TypeIdName = "unsigned int"
    Case JVT_UINT64, JVT_NUMBER_UNSIGNED: TypeIdName = "unsigned __int64"
    Case JVT_NUMBER_INTEGER: TypeIdName = "__int64"
    Case JVT_BOOL: TypeIdName = "bool"
    Case JVT_DOUBLE, JVT_NUMBER_FLOAT: TypeIdName = "double"
    Case JVT_FLOAT: TypeIdName = "float"
    Case JVT_STRING: TypeIdName = "class std::basic_string<char,struct std::char_traits<char>,class std::allocator<char> >"
    Case JVT_WZSTRING: TypeIdName = "class WzString"
    End Select
End Function

' template<TYPE> json_variant_toType(json, ok, defaultValue): get<TYPE>(), or the default (and a warning)
Public Function json_variant_toType(ByVal json As json_variant, ByVal typ As Long, ByVal defaultValue As Variant, Optional ok As Variant) As Variant
    Dim result As Variant, j As nlohmann_json
    result = defaultValue
    Set j = json.jsonValue()
    On Error GoTo failed
    Select Case typ
    Case JVT_INT: result = j.get_int()
    Case JVT_UINT: result = j.get_uint()
    Case JVT_UINT64, JVT_NUMBER_UNSIGNED: result = j.get_uint64()
    Case JVT_NUMBER_INTEGER: result = j.get_int64()
    Case JVT_BOOL: result = j.get_bool()
    Case JVT_DOUBLE, JVT_NUMBER_FLOAT: result = j.get_double()
    Case JVT_FLOAT: result = j.get_float()
    Case JVT_STRING, JVT_WZSTRING: result = j.get_string()
    End Select
    If Not IsMissing(ok) Then
        ok = True
    End If
    json_variant_toType = result
    Exit Function
failed:
    If Err.Source = "nlohmann::json" Then
        debug_ 573, LOG_WARNING, "json_variant_toType", "Failed to convert json_variant to %s because of error: %s", TypeIdName(typ), Err.Description
        If Not IsMissing(ok) Then
            ok = False
        End If
        result = defaultValue
    Else
        debug_ 581, LOG_FATAL, "json_variant_toType", "Unexpected exception encountered: json_variant::toType<%s>", TypeIdName(typ)
    End If
    json_variant_toType = result
End Function

' Convenience methods

' json_getValue(json, key, defaultValue = json_variant())
Public Function json_getValue(ByVal json As nlohmann_json, ByVal key As String, Optional ByVal defaultValue As json_variant = Nothing) As json_variant
    Dim it As nlohmann_json
    Set it = json.find(key)
    If it Is Nothing Then
        If defaultValue Is Nothing Then Set json_getValue = json_variant_new() Else Set json_getValue = defaultValue
    Else
        Set json_getValue = json_variant_new(it)
    End If
End Function

' json_getValue(json, size_type idx, defaultValue = json_variant())
Public Function json_getValue_idx(ByVal json As nlohmann_json, ByVal idx As Long, Optional ByVal defaultValue As json_variant = Nothing) As json_variant
    Dim v As nlohmann_json
    On Error GoTo failed
    Set v = json.at_idx(idx)
    On Error GoTo 0
    Set json_getValue_idx = json_variant_new(v)
    Exit Function
failed:
    If Err.Source <> "nlohmann::json" Then
        debug_ 899, LOG_FATAL, "json_getValue", "Encountered unexpected exception in json_getValue"
    End If
    If defaultValue Is Nothing Then Set json_getValue_idx = json_variant_new() Else Set json_getValue_idx = defaultValue
End Function
