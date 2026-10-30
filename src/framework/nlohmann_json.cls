VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nlohmann_json"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' nlohmann::json (json.hpp 3.11.3) and nlohmann::ordered_json: one JSON
' value. The other projects get values from json_parse, json_object,
' json_array and json_value (FrameworkApi) and through these methods.
'
' A C++ json is a value; this is a reference. Where the C++ copies a value
' (json a = b; obj[key] = v; push_back(v)), the port copies: clone, assign,
' set_item and push_back copy their argument; set_item_ref and push_back_ref
' take it as it is (for a temporary or a std::move).
'
' The methods keep the C++ names; overloads by argument type get a suffix:
' at(key) / at_idx(i), item(key) / item_idx(i) (operator[], which inserts
' as the non-const operator does), get<T>() is get_bool, get_string,
' get_int (int), get_uint (unsigned: the uint32_t bits), get_int64,
' get_uint64 (as Double), get_double, get_float; erase(key) is erase_
' (Erase is a VB statement); value(key, default) is
' value_int, value_uint, value_bool, value_string, value_double,
' value_float. Iterating: size(), key_at(i) and value_at(i), in the
' library's order (json: keys sorted by their bytes; ordered_json: as
' inserted). Errors are raised as the library throws them: Err.Number is
' vbObjectError + the exception id, Err.Description its what().
Option Explicit

Private m_type As value_t
Private m_ordered As Boolean            ' nlohmann::ordered_json: objects keep their insertion order
Private m_bool As Boolean
Private m_num As Variant                ' numbers: Double, or Decimal for an integer beyond 2^53
Private m_str As String
Private m_count As Long                 ' array and object size
Private m_keys() As String              ' object keys
Private m_items() As nlohmann_json      ' array elements, object values

'--------------------------------------------------------------------------
' Friend: for json_hpp (the parser, the factories) and the other instances

Friend Sub init_value(ByVal t As value_t, ByVal ordered As Boolean)
    m_type = t
    m_ordered = ordered
    m_bool = False
    m_num = Empty
    m_str = vbNullString
    m_count = 0
    Erase m_keys
    Erase m_items
    Select Case t
    Case value_t_number_integer, value_t_number_unsigned, value_t_number_float
        m_num = 0#
    End Select
End Sub

Friend Property Get f_ordered() As Boolean
    f_ordered = m_ordered
End Property

Friend Property Let f_bool(ByVal v As Boolean)
    m_bool = v
End Property

Friend Property Get f_bool() As Boolean
    f_bool = m_bool
End Property

Friend Property Let f_num(ByVal v As Variant)
    m_num = v
End Property

Friend Property Get f_num() As Variant
    f_num = m_num
End Property

Friend Property Let f_str(ByVal v As String)
    m_str = v
End Property

Friend Property Get f_str() As String
    f_str = m_str
End Property

Friend Property Get f_count() As Long
    f_count = m_count
End Property

Friend Property Get f_key(ByVal i As Long) As String
    f_key = m_keys(i)
End Property

Friend Property Get f_item(ByVal i As Long) As nlohmann_json
    Set f_item = m_items(i)
End Property

' The index of key's entry, inserted (as null) when missing: object::operator[]
Friend Function object_slot(ByVal key As String) As Long
    Dim i As Long, j As Long, v As nlohmann_json
    If FindKey(key, i) Then
        object_slot = i
        Exit Function
    End If
    GrowTo m_count + 1, True
    For j = m_count To i + 1 Step -1
        m_keys(j) = m_keys(j - 1)
        Set m_items(j) = m_items(j - 1)
    Next
    Set v = New nlohmann_json
    v.init_value value_t_null, m_ordered
    m_keys(i) = key
    Set m_items(i) = v
    m_count = m_count + 1
    object_slot = i
End Function

Friend Sub set_slot(ByVal i As Long, ByVal v As nlohmann_json)
    Set m_items(i) = v
End Sub

Friend Sub append_item(ByVal v As nlohmann_json)
    GrowTo m_count + 1, False
    Set m_items(m_count) = v
    m_count = m_count + 1
End Sub

' *this = src, deep
Friend Sub copy_from(ByVal src As nlohmann_json)
    Dim i As Long, n As Long, t As value_t, keys() As String, items() As nlohmann_json
    t = src.type_
    n = src.f_count
    If t = value_t_object Or t = value_t_array Then
        If n > 0 Then
            ReDim items(0 To n - 1)
            If t = value_t_object Then ReDim keys(0 To n - 1)
            For i = 0 To n - 1
                Set items(i) = src.f_item(i).clone()
                If t = value_t_object Then keys(i) = src.f_key(i)
            Next
        End If
    End If
    init_value t, src.f_ordered
    m_bool = src.f_bool
    m_num = src.f_num
    m_str = src.f_str
    m_count = n
    If n > 0 Then
        m_items = items
        If t = value_t_object Then m_keys = keys
    End If
End Sub

' serializer::dump
Friend Sub dump_to(ByVal pretty_print As Boolean, ByVal ensure_ascii As Boolean, ByVal indent_step As Long, ByVal current_indent As Long, ByVal indent_char As String, ByVal error_handler As error_handler_t)
    Dim i As Long, new_indent As Long
    Select Case m_type
    Case value_t_object
        If m_count = 0 Then
            json_hpp.sb_add "{}"
            Exit Sub
        End If
        If pretty_print Then
            json_hpp.sb_add "{" & vbLf
            new_indent = current_indent + indent_step
            For i = 0 To m_count - 1
                json_hpp.sb_add String$(new_indent, indent_char) & """"
                json_hpp.dump_escaped m_keys(i), ensure_ascii, error_handler
                json_hpp.sb_add """: "
                m_items(i).dump_to True, ensure_ascii, indent_step, new_indent, indent_char, error_handler
                If i < m_count - 1 Then json_hpp.sb_add "," & vbLf
            Next
            json_hpp.sb_add vbLf & String$(current_indent, indent_char) & "}"
        Else
            json_hpp.sb_add "{"
            For i = 0 To m_count - 1
                json_hpp.sb_add """"
                json_hpp.dump_escaped m_keys(i), ensure_ascii, error_handler
                json_hpp.sb_add """:"
                m_items(i).dump_to False, ensure_ascii, indent_step, current_indent, indent_char, error_handler
                If i < m_count - 1 Then json_hpp.sb_add ","
            Next
            json_hpp.sb_add "}"
        End If
    Case value_t_array
        If m_count = 0 Then
            json_hpp.sb_add "[]"
            Exit Sub
        End If
        If pretty_print Then
            json_hpp.sb_add "[" & vbLf
            new_indent = current_indent + indent_step
            For i = 0 To m_count - 1
                json_hpp.sb_add String$(new_indent, indent_char)
                m_items(i).dump_to True, ensure_ascii, indent_step, new_indent, indent_char, error_handler
                If i < m_count - 1 Then json_hpp.sb_add "," & vbLf
            Next
            json_hpp.sb_add vbLf & String$(current_indent, indent_char) & "]"
        Else
            json_hpp.sb_add "["
            For i = 0 To m_count - 1
                m_items(i).dump_to False, ensure_ascii, indent_step, current_indent, indent_char, error_handler
                If i < m_count - 1 Then json_hpp.sb_add ","
            Next
            json_hpp.sb_add "]"
        End If
    Case value_t_string
        json_hpp.sb_add """"
        json_hpp.dump_escaped m_str, ensure_ascii, error_handler
        json_hpp.sb_add """"
    Case value_t_boolean
        If m_bool Then json_hpp.sb_add "true" Else json_hpp.sb_add "false"
    Case value_t_number_integer, value_t_number_unsigned
        json_hpp.sb_add json_hpp.dump_integer(m_num)
    Case value_t_number_float
        json_hpp.sb_add json_hpp.dump_float(m_num)
    Case value_t_discarded
        json_hpp.sb_add "<discarded>"
    Case Else
        json_hpp.sb_add "null"
    End Select
End Sub

'--------------------------------------------------------------------------
' Helpers

' The key's index (True), or where it would go (False). json: binary search
' in the sorted keys; ordered_json: in order.
' (VB) Keys compare as UTF-16 (StrComp binary); std::string compares UTF-8
' bytes: the same order except for characters beyond U+FFFF.
Private Function FindKey(ByVal key As String, pos As Long) As Boolean
    Dim lo As Long, hi As Long, mid_ As Long, c As Long
    If m_ordered Then
        For lo = 0 To m_count - 1
            If m_keys(lo) = key Then
                pos = lo
                FindKey = True
                Exit Function
            End If
        Next
        pos = m_count
        Exit Function
    End If
    lo = 0
    hi = m_count - 1
    Do While lo <= hi
        mid_ = (lo + hi) \ 2
        c = StrComp(m_keys(mid_), key, vbBinaryCompare)
        If c = 0 Then
            pos = mid_
            FindKey = True
            Exit Function
        ElseIf c < 0 Then
            lo = mid_ + 1
        Else
            hi = mid_ - 1
        End If
    Loop
    pos = lo
End Function

Private Sub GrowTo(ByVal n As Long, ByVal withKeys As Boolean)
    Dim cap As Long
    If m_count = 0 Then
        cap = 4
        If cap < n Then cap = n
        ReDim m_items(0 To cap - 1)
        If withKeys Then ReDim m_keys(0 To cap - 1)
    ElseIf n > UBound(m_items) + 1 Then
        cap = 2 * (UBound(m_items) + 1)
        If cap < n Then cap = n
        ReDim Preserve m_items(0 To cap - 1)
        If withKeys Then ReDim Preserve m_keys(0 To cap - 1)
    End If
End Sub

Private Function NewNull() As nlohmann_json
    Set NewNull = New nlohmann_json
    NewNull.init_value value_t_null, m_ordered
End Function

Private Sub RaiseTypeError(ByVal id As Long, ByVal msg As String)
    json_hpp.json_throw "type_error", id, msg
End Sub

Private Sub RaiseNotNumber()
    RaiseTypeError 302, "type must be number, but is " & type_name()
End Sub

' static_cast<int32_t>(integer value): the low 32 bits. A Double stays in
' Double arithmetic (exact to 2^53; CDec would keep 15 digits).
Private Function WrapToLong(ByVal v As Variant) As Long
    Dim w As Variant, d As Variant
    If v >= -2147483648# And v <= 2147483647# Then
        WrapToLong = CLng(v)
        Exit Function
    End If
    If VarType(v) = vbDecimal Then d = CDec(4294967296#) Else d = 4294967296#
    w = v - Int(v / d) * d
    If w < 0 Then w = w + d
    If w >= d Then w = w - d
    If w >= 2147483648# Then w = w - d
    WrapToLong = CLng(w)
End Function

'--------------------------------------------------------------------------
' Type

Public Property Get type_() As value_t
    type_ = m_type
End Property

Public Function type_name() As String
    Select Case m_type
    Case value_t_null: type_name = "null"
    Case value_t_object: type_name = "object"
    Case value_t_array: type_name = "array"
    Case value_t_string: type_name = "string"
    Case value_t_boolean: type_name = "boolean"
    Case value_t_binary: type_name = "binary"
    Case value_t_discarded: type_name = "discarded"
    Case Else: type_name = "number"
    End Select
End Function

Public Function is_ordered() As Boolean
    is_ordered = m_ordered
End Function

Public Function is_null() As Boolean
    is_null = (m_type = value_t_null)
End Function

Public Function is_boolean() As Boolean
    is_boolean = (m_type = value_t_boolean)
End Function

Public Function is_number() As Boolean
    is_number = (m_type = value_t_number_integer Or m_type = value_t_number_unsigned Or m_type = value_t_number_float)
End Function

Public Function is_number_integer() As Boolean
    is_number_integer = (m_type = value_t_number_integer Or m_type = value_t_number_unsigned)
End Function

Public Function is_number_unsigned() As Boolean
    is_number_unsigned = (m_type = value_t_number_unsigned)
End Function

Public Function is_number_float() As Boolean
    is_number_float = (m_type = value_t_number_float)
End Function

Public Function is_string() As Boolean
    is_string = (m_type = value_t_string)
End Function

Public Function is_array() As Boolean
    is_array = (m_type = value_t_array)
End Function

Public Function is_object() As Boolean
    is_object = (m_type = value_t_object)
End Function

Public Function is_primitive() As Boolean
    is_primitive = is_null() Or is_string() Or is_boolean() Or is_number() Or m_type = value_t_binary
End Function

Public Function is_structured() As Boolean
    is_structured = is_array() Or is_object()
End Function

Public Function is_discarded() As Boolean
    is_discarded = (m_type = value_t_discarded)
End Function

'--------------------------------------------------------------------------
' Capacity and lookup

' null: 0; array, object: their size; others: 1
Public Function size() As Long
    Select Case m_type
    Case value_t_null: size = 0
    Case value_t_array, value_t_object: size = m_count
    Case Else: size = 1
    End Select
End Function

' C: empty()
Public Function empty_() As Boolean
    Select Case m_type
    Case value_t_null: empty_ = True
    Case value_t_array, value_t_object: empty_ = (m_count = 0)
    Case Else: empty_ = False
    End Select
End Function

Public Function contains(ByVal key As String) As Boolean
    Dim i As Long
    If m_type = value_t_object Then contains = FindKey(key, i)
End Function

Public Function count(ByVal key As String) As Long
    If contains(key) Then count = 1
End Function

' find(key): the value, or Nothing for end()
Public Function find(ByVal key As String) As nlohmann_json
    Dim i As Long
    If m_type = value_t_object Then
        If FindKey(key, i) Then Set find = m_items(i)
    End If
End Function

Public Function at(ByVal key As String) As nlohmann_json
    Dim i As Long
    If m_type = value_t_object Then
        If Not FindKey(key, i) Then
            json_hpp.json_throw "out_of_range", 403, "key '" & key & "' not found"
        End If
        Set at = m_items(i)
        Exit Function
    End If
    RaiseTypeError 304, "cannot use at() with " & type_name()
End Function

Public Function at_idx(ByVal idx As Long) As nlohmann_json
    If m_type = value_t_array Then
        If idx < 0 Or idx >= m_count Then
            json_hpp.json_throw "out_of_range", 401, "array index " & json_hpp.size_t_string(idx) & " is out of range"
        End If
        Set at_idx = m_items(idx)
        Exit Function
    End If
    RaiseTypeError 304, "cannot use at() with " & type_name()
End Function

' operator[](key): a null value becomes an object; a missing key is inserted as null
Public Function item(ByVal key As String) As nlohmann_json
    If m_type = value_t_null Then init_value value_t_object, m_ordered
    If m_type = value_t_object Then
        Set item = m_items(object_slot(key))
        Exit Function
    End If
    RaiseTypeError 305, "cannot use operator[] with a string argument with " & type_name()
End Function

' operator[](idx): a null value becomes an array; an array grows with nulls to reach idx
Public Function item_idx(ByVal idx As Long) As nlohmann_json
    If m_type = value_t_null Then init_value value_t_array, m_ordered
    If m_type = value_t_array Then
        Do While idx >= m_count
            append_item NewNull()
        Loop
        Set item_idx = m_items(idx)
        Exit Function
    End If
    RaiseTypeError 305, "cannot use operator[] with a numeric argument with " & type_name()
End Function

' Iteration: the i-th key of an object
Public Function key_at(ByVal i As Long) As String
    If m_type <> value_t_object Then
        json_hpp.json_throw "invalid_iterator", 207, "cannot use key() for non-object iterators"
    End If
    key_at = m_keys(i)
End Function

' Iteration: the i-th value (an array's element, an object's value; a primitive is its one value)
Public Function value_at(ByVal i As Long) As nlohmann_json
    If m_type = value_t_object Or m_type = value_t_array Then
        Set value_at = m_items(i)
    Else
        Set value_at = Me
    End If
End Function

Public Function front() As nlohmann_json
    If m_type = value_t_null Then json_hpp.json_throw "invalid_iterator", 214, "cannot get value"
    Set front = value_at(0)
End Function

Public Function back() As nlohmann_json
    If m_type = value_t_null Then json_hpp.json_throw "invalid_iterator", 214, "cannot get value"
    Set back = value_at(size() - 1)
End Function

'--------------------------------------------------------------------------
' get<T>()

Public Function get_bool() As Boolean
    If m_type <> value_t_boolean Then RaiseTypeError 302, "type must be boolean, but is " & type_name()
    get_bool = m_bool
End Function

Public Function get_string() As String
    If m_type <> value_t_string Then RaiseTypeError 302, "type must be string, but is " & type_name()
    get_string = m_str
End Function

' get<int>(): a number (truncated, wrapped as static_cast does) or a boolean
Public Function get_int() As Long
    Select Case m_type
    Case value_t_number_integer, value_t_number_unsigned
        get_int = WrapToLong(m_num)
    Case value_t_number_float
        get_int = ToInt(m_num)
    Case value_t_boolean
        get_int = -CLng(m_bool)
    Case Else
        RaiseNotNumber
    End Select
End Function

' get<unsigned>(): the uint32_t's bits
Public Function get_uint() As Long
    Select Case m_type
    Case value_t_number_integer, value_t_number_unsigned
        get_uint = WrapToLong(m_num)
    Case value_t_number_float
        If m_num > -9.2E+18 And m_num < 9.2E+18 Then get_uint = WrapToLong(Fix(m_num)) Else get_uint = 0
    Case value_t_boolean
        get_uint = -CLng(m_bool)
    Case Else
        RaiseNotNumber
    End Select
End Function

' get<int64_t>() (no boolean: get_arithmetic_value); as a Double (exact to 2^53)
Public Function get_int64() As Double
    Select Case m_type
    Case value_t_number_integer, value_t_number_unsigned
        get_int64 = CDbl(m_num)
    Case value_t_number_float
        get_int64 = Fix(m_num)
    Case Else
        RaiseNotNumber
    End Select
End Function

' get<uint64_t>() (no boolean); as a Double (exact to 2^53)
Public Function get_uint64() As Double
    Select Case m_type
    Case value_t_number_integer, value_t_number_unsigned
        If m_num < 0 Then get_uint64 = CDbl(m_num) + 1.84467440737096E+19 Else get_uint64 = CDbl(m_num)
    Case value_t_number_float
        get_uint64 = Fix(m_num)
    Case Else
        RaiseNotNumber
    End Select
End Function

' get<double>() (no boolean)
Public Function get_double() As Double
    Select Case m_type
    Case value_t_number_integer, value_t_number_unsigned, value_t_number_float
        get_double = CDbl(m_num)
    Case Else
        RaiseNotNumber
    End Select
End Function

' get<float>(): a number or a boolean
Public Function get_float() As Single
    Select Case m_type
    Case value_t_number_integer, value_t_number_unsigned, value_t_number_float
        get_float = CSng(m_num)
    Case value_t_boolean
        get_float = -CSng(m_bool)
    Case Else
        RaiseNotNumber
    End Select
End Function

'--------------------------------------------------------------------------
' value(key, default)

Private Function ValueOf(ByVal key As String) As nlohmann_json
    If m_type <> value_t_object Then RaiseTypeError 306, "cannot use value() with " & type_name()
    Set ValueOf = find(key)
End Function

Public Function value_int(ByVal key As String, ByVal default_value As Long) As Long
    Dim v As nlohmann_json
    Set v = ValueOf(key)
    If v Is Nothing Then value_int = default_value Else value_int = v.get_int()
End Function

Public Function value_uint(ByVal key As String, ByVal default_value As Long) As Long
    Dim v As nlohmann_json
    Set v = ValueOf(key)
    If v Is Nothing Then value_uint = default_value Else value_uint = v.get_uint()
End Function

Public Function value_bool(ByVal key As String, ByVal default_value As Boolean) As Boolean
    Dim v As nlohmann_json
    Set v = ValueOf(key)
    If v Is Nothing Then value_bool = default_value Else value_bool = v.get_bool()
End Function

Public Function value_string(ByVal key As String, ByVal default_value As String) As String
    Dim v As nlohmann_json
    Set v = ValueOf(key)
    If v Is Nothing Then value_string = default_value Else value_string = v.get_string()
End Function

Public Function value_double(ByVal key As String, ByVal default_value As Double) As Double
    Dim v As nlohmann_json
    Set v = ValueOf(key)
    If v Is Nothing Then value_double = default_value Else value_double = v.get_double()
End Function

Public Function value_float(ByVal key As String, ByVal default_value As Single) As Single
    Dim v As nlohmann_json
    Set v = ValueOf(key)
    If v Is Nothing Then value_float = default_value Else value_float = v.get_float()
End Function

'--------------------------------------------------------------------------
' Modifiers

' (*this)[key] = v: v is copied into the key's value (references to that value see it)
Public Sub set_item(ByVal key As String, ByVal v As nlohmann_json)
    item(key).assign v
End Sub

' (*this)[key] = std::move(v): v itself becomes the key's value
Public Sub set_item_ref(ByVal key As String, ByVal v As nlohmann_json)
    Dim i As Long
    If m_type = value_t_null Then init_value value_t_object, m_ordered
    If m_type <> value_t_object Then RaiseTypeError 305, "cannot use operator[] with a string argument with " & type_name()
    i = object_slot(key)
    Set m_items(i) = v
End Sub

' (*this)[idx] = v
Public Sub set_item_idx(ByVal idx As Long, ByVal v As nlohmann_json)
    item_idx(idx).assign v
End Sub

' push_back(v): a copy of v; a null value becomes an array
Public Sub push_back(ByVal v As nlohmann_json)
    push_back_ref v.clone()
End Sub

' push_back(std::move(v))
Public Sub push_back_ref(ByVal v As nlohmann_json)
    If Not (m_type = value_t_null Or m_type = value_t_array) Then
        RaiseTypeError 308, "cannot use push_back() with " & type_name()
    End If
    If m_type = value_t_null Then init_value value_t_array, m_ordered
    append_item v
End Sub

' erase(key): the number of values removed (0 or 1)
Public Function erase_(ByVal key As String) As Long
    Dim i As Long, j As Long
    If m_type <> value_t_object Then RaiseTypeError 307, "cannot use erase() with " & type_name()
    If Not FindKey(key, i) Then Exit Function
    For j = i To m_count - 2
        m_keys(j) = m_keys(j + 1)
        Set m_items(j) = m_items(j + 1)
    Next
    m_count = m_count - 1
    m_keys(m_count) = vbNullString
    Set m_items(m_count) = Nothing
    erase_ = 1
End Function

' erase(idx) of an array
Public Sub erase_idx(ByVal idx As Long)
    Dim j As Long
    If m_type <> value_t_array Then RaiseTypeError 307, "cannot use erase() with " & type_name()
    If idx < 0 Or idx >= m_count Then
        json_hpp.json_throw "out_of_range", 401, "array index " & json_hpp.size_t_string(idx) & " is out of range"
    End If
    For j = idx To m_count - 2
        Set m_items(j) = m_items(j + 1)
    Next
    m_count = m_count - 1
    Set m_items(m_count) = Nothing
End Sub

' clear(): numbers become 0, booleans false, strings "", arrays and objects empty
Public Sub clear()
    Select Case m_type
    Case value_t_number_integer, value_t_number_unsigned, value_t_number_float
        m_num = 0#
    Case value_t_boolean
        m_bool = False
    Case value_t_string
        m_str = ""
    Case value_t_array, value_t_object
        m_count = 0
        Erase m_keys
        Erase m_items
    End Select
End Sub

' *this = v (a copy)
Public Sub assign(ByVal v As nlohmann_json)
    If v Is Me Then Exit Sub
    copy_from v
End Sub

' json copy = *this
Public Function clone() As nlohmann_json
    Dim c As nlohmann_json
    Set c = New nlohmann_json
    c.copy_from Me
    Set clone = c
End Function

' operator==
Public Function equals(ByVal v As nlohmann_json) As Boolean
    Dim i As Long, t As value_t
    t = v.type_
    If m_type = t Then
        Select Case m_type
        Case value_t_array, value_t_object
            If m_count <> v.f_count Then Exit Function
            For i = 0 To m_count - 1
                If m_type = value_t_object Then
                    If m_keys(i) <> v.f_key(i) Then Exit Function
                End If
                If Not m_items(i).equals(v.f_item(i)) Then Exit Function
            Next
            equals = True
        Case value_t_string
            equals = (m_str = v.f_str)
        Case value_t_boolean
            equals = (m_bool = v.f_bool)
        Case value_t_number_integer, value_t_number_unsigned, value_t_number_float
            equals = (m_num = v.f_num)
        Case Else
            equals = True
        End Select
    ElseIf is_number() And v.is_number() Then
        equals = (CDbl(m_num) = CDbl(v.f_num))
    End If
End Function

' dump(indent = -1, indent_char = ' ', ensure_ascii = false, error_handler = strict)
Public Function dump(Optional ByVal indent As Long = -1, Optional ByVal indent_char As String = " ", Optional ByVal ensure_ascii As Boolean = False, Optional ByVal error_handler As error_handler_t = error_handler_t_strict) As String
    json_hpp.sb_reset
    If indent >= 0 Then
        dump_to True, ensure_ascii, indent, 0, indent_char, error_handler
    Else
        dump_to False, ensure_ascii, 0, 0, indent_char, error_handler
    End If
    dump = json_hpp.sb_result()
End Function
