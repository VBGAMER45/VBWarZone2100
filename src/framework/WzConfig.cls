VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/framework/wzconfig.cpp: class WzConfig, a JSON file read (and written)
' a group or array at a time. Made by WzConfig_new(name, warning) (the
' constructor); the destructor's save runs when the last reference goes
' (Class_Terminate), as when the C++ object goes out of scope. pCurrentObj
' and the stacks hold references into mRoot, mArray and the new objects.
' Vectors are returned in their fields (ByRef x, y, z). string() is string_
' (String is a VB type name), set() is set_.
Option Explicit

Private mRoot As nlohmann_json
Private pCurrentObj As nlohmann_json   ' the "current" json object
Private mArray As nlohmann_json
Private mName As String
Private mObjStack() As nlohmann_json
Private mObjStackSize As Long
Private mNewObjStack() As nlohmann_json ' stores newly-created objects (before they are added to root)
Private mNewObjStackSize As Long
Private mObjNameStack() As String
Private mObjNameStackSize As Long
Private mFilename As String
Private mStatus As Boolean
Private mWarning As WzConfig_warning
Private mConstructed As Boolean         ' (VB) WzConfig_new ran

Private Sub Class_Initialize()
    Set mRoot = json_object()
    ReDim mObjStack(0 To 7)
    ReDim mNewObjStack(0 To 7)
    ReDim mObjNameStack(0 To 7)
End Sub

'--------------------------------------------------------------------------
' The lists (std::list): push_back, back, pop_back

Private Sub PushObj(ByVal v As nlohmann_json)
    If mObjStackSize > UBound(mObjStack) Then ReDim Preserve mObjStack(0 To 2 * mObjStackSize - 1)
    Set mObjStack(mObjStackSize) = v
    mObjStackSize = mObjStackSize + 1
End Sub

Private Function PopObj() As nlohmann_json
    mObjStackSize = mObjStackSize - 1
    Set PopObj = mObjStack(mObjStackSize)
    Set mObjStack(mObjStackSize) = Nothing
End Function

Private Sub PushNewObj(ByVal v As nlohmann_json)
    If mNewObjStackSize > UBound(mNewObjStack) Then ReDim Preserve mNewObjStack(0 To 2 * mNewObjStackSize - 1)
    Set mNewObjStack(mNewObjStackSize) = v
    mNewObjStackSize = mNewObjStackSize + 1
End Sub

Private Sub PopNewObj()
    mNewObjStackSize = mNewObjStackSize - 1
    Set mNewObjStack(mNewObjStackSize) = Nothing
End Sub

' !mNewObjStack.empty() && &mNewObjStack.back() == p
Private Function IsLatestNewObj(ByVal p As nlohmann_json) As Boolean
    If mNewObjStackSize > 0 Then IsLatestNewObj = (mNewObjStack(mNewObjStackSize - 1) Is p)
End Function

Private Sub PushName(ByVal s As String)
    If mObjNameStackSize > UBound(mObjNameStack) Then ReDim Preserve mObjNameStack(0 To 2 * mObjNameStackSize - 1)
    mObjNameStack(mObjNameStackSize) = s
    mObjNameStackSize = mObjNameStackSize + 1
End Sub

Private Function PopName() As String
    mObjNameStackSize = mObjNameStackSize - 1
    PopName = mObjNameStack(mObjNameStackSize)
End Function

'--------------------------------------------------------------------------

Private Sub Class_Terminate()
    Dim jsonString As String, b() As Byte, n As Long, s As String
    If Not mConstructed Then Exit Sub
    If mWarning = WzConfig_ReadAndWrite Then
        If Not (mObjStackSize = 0) Then
            ASSERT_FAILURE "mObjStack.empty()", "lib/framework/wzconfig.cpp:37", "~WzConfig", "Some json groups have not been closed, stack size %zu.", mObjStackSize
        End If
        jsonString = mRoot.dump(4) & vbLf            ' stream << mRoot.dump(4) << std::endl
        n = utf8_encode(jsonString, b)
        ' (jsonString.size() <= UDWORD max: a String never exceeds it)
        If n = 0 Then ReDim b(0 To 0)
        saveFile mFilename, b, n
    End If
    If mWarning = WzConfig_ReadAndWrite Then s = "Saving" Else s = "Closing"
    debug_ 46, LOG_SAVE, "~WzConfig", "%s %s", s, mFilename
End Sub

' WzConfig::WzConfig(name, warning): WzConfig_new
Friend Sub construct(ByVal name_ As String, ByVal warning As WzConfig_warning)
    Dim size As Long, data() As Byte, folders() As String, nFolders As Long, k As Long
    Dim str_ As String, tmpJson As nlohmann_json, ok As Boolean

    Set mArray = json_array()
    mConstructed = True

    mFilename = name_
    mStatus = True
    mWarning = warning
    Set pCurrentObj = mRoot

    If PHYSFS_exists(name_) = 0 Then
        If warning = WzConfig_ReadOnly Then
            mStatus = False
            Exit Sub
        ElseIf warning = WzConfig_ReadOnlyAndRequired Then
            debug_ 91, LOG_FATAL, "WzConfig", "Missing required file %s", name_
            exit_ 3                                 ' abort()
        ElseIf warning = WzConfig_ReadAndWrite Then
            Exit Sub
        End If
    End If
    If Not loadFile(name_, data, size) Then
        mStatus = False
        debug_ 102, LOG_FATAL, "WzConfig", "Could not open ""%s""", name_
        Exit Sub
    End If

    On Error GoTo parse_failed
    Set mRoot = json_parse_bytes(data, 0, size)
    On Error GoTo 0
parsed:
    On Error GoTo 0
    Set pCurrentObj = mRoot
    If Not (Not mRoot.is_null()) Then
        ASSERT_FAILURE "!mRoot.is_null()", "lib/framework/wzconfig.cpp:117", "WzConfig", "JSON document from %s is null", name_
    End If
    If Not mRoot.is_object() Then
        ASSERT_FAILURE "mRoot.is_object()", "lib/framework/wzconfig.cpp:120", "WzConfig", "JSON document from %s is not an object. Read: \n%s", name_, utf8_decode(data, 0, size)
        Set mRoot = json_object()
        mStatus = False
        Exit Sub
    End If
    Erase data

    ok = WZ_PHYSFS_enumerateFolders("diffs", folders, nFolders)
    For k = 0 To nFolders - 1
        str_ = "diffs/" & folders(k) & "/" & name_
        If PHYSFS_exists(str_) = 0 Then
            GoTo continue_
        End If
        size = 0
        Erase data
        If Not loadFile(str_, data, size) Then
            debug_ 137, LOG_FATAL, "WzConfig", "jsondiff file ""%s"" could not be opened!", name_
        End If
        Set tmpJson = json_value(Empty)
        On Error GoTo diff_parse_failed
        Set tmpJson = json_parse_bytes(data, 0, size)
        On Error GoTo 0
diff_parsed:
        On Error GoTo 0
        If Not (Not tmpJson.is_null()) Then
            ASSERT_FAILURE "!tmpJson.is_null()", "lib/framework/wzconfig.cpp:149", "WzConfig", "JSON diff from %s is null", name_
        End If
        If tmpJson.is_object() Then
            Set mRoot = jsonMerge(mRoot, tmpJson)
        Else
            ASSERT_FAILURE "tmpJson.is_object()", "lib/framework/wzconfig.cpp:156", "WzConfig", "JSON diff from %s is not an object. Read: \n%s", name_, utf8_decode(data, 0, size)
        End If
        debug_ 159, LOG_INFO, "WzConfig", "jsondiff ""%s"" loaded and merged", str_
continue_:
    Next
    debug_ 162, LOG_SAVE, "WzConfig", "Opening %s", name_
    Set pCurrentObj = mRoot
    Exit Sub

parse_failed:
    If Err.Source = "nlohmann::json" Then
        ASSERT_FAILURE "false", "lib/framework/wzconfig.cpp:111", "WzConfig", "JSON document from %s is invalid: %s", name_, Err.Description
    Else
        debug_ 114, LOG_FATAL, "WzConfig", "Unexpected exception parsing JSON %s", name_
    End If
    Resume parsed

diff_parse_failed:
    If Err.Source = "nlohmann::json" Then
        ASSERT_FAILURE "false", "lib/framework/wzconfig.cpp:144", "WzConfig", "JSON diff from %s is invalid: %s", name_, Err.Description
    Else
        debug_ 147, LOG_FATAL, "WzConfig", "Unexpected exception parsing JSON diff from %s", name_
    End If
    Resume diff_parsed
End Sub

Public Function isAtDocumentRoot() As Boolean
    isAtDocumentRoot = (pCurrentObj Is mRoot)
End Function

Public Function compactStringRepresentation(Optional ByVal ensure_ascii As Boolean = False) As String
    ' Use the most compact representation of the JSON
    compactStringRepresentation = mRoot.dump(-1, " ", ensure_ascii)
End Function

' C: std::vector<WzString> childGroups(): the keys whose values are objects, in keys(0 ..); the count
Public Function childGroups(keys() As String) As Long
    Dim i As Long, n As Long
    If Not (Not pCurrentObj Is Nothing) Then
        ASSERT_FAILURE "pCurrentObj != nullptr", "lib/framework/wzconfig.cpp:179", "childGroups", "current object is nullptr"
        Exit Function
    End If
    If Not pCurrentObj.is_object() Then
        ASSERT_FAILURE "pCurrentObj->is_object()", "lib/framework/wzconfig.cpp:180", "childGroups", "current pointer is not an object"
        Exit Function
    End If
    For i = 0 To pCurrentObj.size() - 1
        If pCurrentObj.value_at(i).is_object() Then
            If n = 0 Then ReDim keys(0 To pCurrentObj.size() - 1)
            keys(n) = pCurrentObj.key_at(i)
            n = n + 1
        End If
    Next
    childGroups = n
End Function

' C: std::vector<WzString> childKeys(): the keys in keys(0 ..); the count
Public Function childKeys(keys() As String) As Long
    Dim i As Long, n As Long
    If Not (Not pCurrentObj Is Nothing) Then
        ASSERT_FAILURE "pCurrentObj != nullptr", "lib/framework/wzconfig.cpp:194", "childKeys", "current object is nullptr"
        Exit Function
    End If
    If Not pCurrentObj.is_object() Then
        ASSERT_FAILURE "pCurrentObj->is_object()", "lib/framework/wzconfig.cpp:195", "childKeys", "current pointer is not an object"
        Exit Function
    End If
    n = pCurrentObj.size()
    If n > 0 Then ReDim keys(0 To n - 1)
    For i = 0 To n - 1
        keys(i) = pCurrentObj.key_at(i)
    Next
    childKeys = n
End Function

Public Function contains(ByVal key As String) As Boolean
    contains = Not (pCurrentObj.find(key) Is Nothing)
End Function

Public Function value(ByVal key As String, Optional ByVal defaultValue As json_variant = Nothing) As json_variant
    Set value = json_getValue(pCurrentObj, key, defaultValue)
End Function

' nlohmann::json json(key, defaultValue = json()): a copy
Public Function json(ByVal key As String, Optional ByVal defaultValue As nlohmann_json = Nothing) As nlohmann_json
    Dim it As nlohmann_json
    Set it = pCurrentObj.find(key)
    If it Is Nothing Then
        If defaultValue Is Nothing Then Set json = json_value(Empty) Else Set json = defaultValue.clone()
    Else
        Set json = it.clone()
    End If
End Function

' nlohmann::json& currentJsonValue()
Public Function currentJsonValue() As nlohmann_json
    If Not (Not pCurrentObj Is Nothing) Then
        ASSERT_FAILURE "pCurrentObj", "lib/framework/wzconfig.cpp:229", "currentJsonValue", "pCurrentObj is null"
    End If
    Set currentJsonValue = pCurrentObj
End Function

' C: WzString string(key, defaultValue = WzString())
Public Function string_(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    Dim it As nlohmann_json, ok As Variant, stringValue As String
    Set it = pCurrentObj.find(key)
    If it Is Nothing Then
        string_ = defaultValue
    Else
        ' Use json_variant to support conversion from other value types to string
        ok = False
        stringValue = json_variant_new(it).toWzString(ok)
        If Not ok Then
            ASSERT_FAILURE "false", "lib/framework/wzconfig.cpp:247", "string", "Failed to convert value of key ""%s"" to string", key
            string_ = defaultValue
            Exit Function
        End If
        string_ = stringValue
    End If
End Function

Public Sub setVector3f(ByVal name_ As String, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim a As nlohmann_json
    Set a = json_array()
    a.push_back_ref json_value(x)
    a.push_back_ref json_value(y)
    a.push_back_ref json_value(z)
    pCurrentObj.set_item_ref name_, a
End Sub

' Vector3f vector3f(name): its fields in x, y, z (0 when the key is missing)
Public Sub vector3f(ByVal name_ As String, x As Single, y As Single, z As Single)
    Dim it As nlohmann_json, v As nlohmann_json
    x = 0
    y = 0
    z = 0
    Set it = pCurrentObj.find(name_)
    If it Is Nothing Then
        Exit Sub
    End If
    Set v = it.clone()
    If Not (v.size() = 3) Then
        ASSERT_FAILURE "v.size() == 3", "lib/framework/wzconfig.cpp:268", "vector3f", "%s: Bad list of %s", mFilename, name_
    End If
    On Error GoTo failed
    x = v.item_idx(0).get_float()
    y = v.item_idx(1).get_float()
    z = v.item_idx(2).get_float()
    Exit Sub
failed:
    If Err.Source = "nlohmann::json" Then
        ASSERT_FAILURE "false", "lib/framework/wzconfig.cpp:275", "vector3f", "%s: Bad list of %s; exception: %s", mFilename, name_, Err.Description
    Else
        debug_ 278, LOG_FATAL, "vector3f", "Unexpected exception encountered in vector3f"
    End If
End Sub

Public Sub setVector3i(ByVal name_ As String, ByVal x As Long, ByVal y As Long, ByVal z As Long)
    Dim a As nlohmann_json
    Set a = json_array()
    a.push_back_ref json_value(x)
    a.push_back_ref json_value(y)
    a.push_back_ref json_value(z)
    pCurrentObj.set_item_ref name_, a
End Sub

' Vector3i vector3i(name)
Public Sub vector3i(ByVal name_ As String, x As Long, y As Long, z As Long)
    Dim it As nlohmann_json, v As nlohmann_json
    x = 0
    y = 0
    z = 0
    Set it = pCurrentObj.find(name_)
    If it Is Nothing Then
        Exit Sub
    End If
    Set v = it.clone()
    If Not (v.size() = 3) Then
        ASSERT_FAILURE "v.size() == 3", "lib/framework/wzconfig.cpp:297", "vector3i", "%s: Bad list of %s", mFilename, name_
    End If
    On Error GoTo failed
    x = v.item_idx(0).get_int()
    y = v.item_idx(1).get_int()
    z = v.item_idx(2).get_int()
    Exit Sub
failed:
    If Err.Source = "nlohmann::json" Then
        ASSERT_FAILURE "false", "lib/framework/wzconfig.cpp:304", "vector3i", "%s: Bad list of %s; exception: %s", mFilename, name_, Err.Description
    Else
        debug_ 307, LOG_FATAL, "vector3i", "Unexpected exception encountered in vector3f"
    End If
End Sub

Public Sub setVector2i(ByVal name_ As String, ByVal x As Long, ByVal y As Long)
    Dim a As nlohmann_json
    Set a = json_array()
    a.push_back_ref json_value(x)
    a.push_back_ref json_value(y)
    pCurrentObj.set_item_ref name_, a
End Sub

' Vector2i vector2i(name)
Public Sub vector2i(ByVal name_ As String, x As Long, y As Long)
    Dim it As nlohmann_json, v As nlohmann_json
    x = 0
    y = 0
    Set it = pCurrentObj.find(name_)
    If it Is Nothing Then
        Exit Sub
    End If
    Set v = it.clone()
    If Not (v.size() = 2) Then
        ASSERT_FAILURE "v.size() == 2", "lib/framework/wzconfig.cpp:326", "vector2i", "Bad list of %s", name_
    End If
    On Error GoTo failed
    x = v.item_idx(0).get_int()
    y = v.item_idx(1).get_int()
    Exit Sub
failed:
    If Err.Source = "nlohmann::json" Then
        ASSERT_FAILURE "false", "lib/framework/wzconfig.cpp:332", "vector2i", "%s: Bad list of %s; exception: %s", mFilename, name_, Err.Description
    Else
        debug_ 335, LOG_FATAL, "vector2i", "Unexpected exception encountered in vector3f"
    End If
End Sub

Public Function beginGroup(ByVal prefix As String) As Boolean
    Dim it As nlohmann_json
    PushName mName
    PushObj pCurrentObj
    mName = prefix
    If mWarning = WzConfig_ReadAndWrite Then
        PushNewObj json_object()
        Set pCurrentObj = mNewObjStack(mNewObjStackSize - 1)
    Else
        Set it = pCurrentObj.find(prefix)
        ' Check if mObj contains the prefix
        If it Is Nothing Then ' handled in this way for backwards compatibility
            PushNewObj json_object()
            Set pCurrentObj = mNewObjStack(mNewObjStackSize - 1)
            beginGroup = False
            Exit Function
        End If
        If Not it.is_object() Then
            ASSERT_FAILURE "it.value().is_object()", "lib/framework/wzconfig.cpp:360", "beginGroup", "%s: beginGroup() on non-object key ""%s""", mFilename, prefix
        End If
        Set pCurrentObj = it
    End If
    beginGroup = True
End Function

Public Sub endGroup()
    Dim pLatestObj As nlohmann_json
    If Not (mObjStackSize > 0) Then
        ASSERT_FAILURE "!mObjStack.empty()", "lib/framework/wzconfig.cpp:368", "endGroup", "An endGroup() too much!"
    End If
    If mWarning = WzConfig_ReadAndWrite Then
        Set pLatestObj = pCurrentObj
        Set pCurrentObj = PopObj()
        If IsLatestNewObj(pLatestObj) Then
            pCurrentObj.set_item_ref mName, pLatestObj  ' a copy whose source is dropped next: taken as it is
            PopNewObj
        Else
            ' this object is not a "new" object, and so is already source from the current obj
            debug_ 382, LOG_INFO, "endGroup", "Got it"
        End If
        mName = PopName()
    Else
        If IsLatestNewObj(pCurrentObj) Then
            PopNewObj
        End If
        mName = PopName()
        Set pCurrentObj = PopObj()
    End If
End Sub

' This function is intended to be paired with endArray().
'
' [Creation Case]:
'
' By calling beginArray with a name, a new array value at the desired key (name) is created in the current object.
' The current object is then set to a **new JSON object**, which can be further manipulated.
'
' Calling nextArrayItem() inserts the current object into the current array, and then sets the current object
' to a new JSON object.
'
' Calling endArray() then finalizes and sets the array value in the (original / parent) object for key "name".
'
' [Read-only Case]:
'
' Similarly, beginArray() let's you traverse the array of objects found at key "name". (With some backwards-compatibility quirks.)
'
' The current object is set to the first object in the array. Subsequent calls to nextArrayItem() will increment which array object
' the current object points to.
'
Public Sub beginArray(ByVal name_ As String)
    Dim it As nlohmann_json
    If Not mArray.empty_() Then
        ASSERT_FAILURE "mArray.empty()", "lib/framework/wzconfig.cpp:421", "beginArray", "beginArray() cannot be nested"
    End If
    PushName mName
    PushObj pCurrentObj
    mName = name_
    If mWarning = WzConfig_ReadAndWrite Then
        PushNewObj json_object()
        Set pCurrentObj = mNewObjStack(mNewObjStackSize - 1)
    Else
        Set it = pCurrentObj.find(name_)
        ' Check if mObj contains the name
        If it Is Nothing Then ' handled in this way for backwards compatibility
            Exit Sub
        End If
        If Not it.is_array() Then
            ASSERT_FAILURE "it.value().is_array()", "lib/framework/wzconfig.cpp:438", "beginArray", "%s: beginArray() on non-array key ""%s""", mFilename, name_
        End If
        Set mArray = it.clone()
        If mArray.size() = 0 Then
            Exit Sub
        End If
        If Not mArray.front().is_object() Then
            ASSERT_FAILURE "mArray.front().is_object()", "lib/framework/wzconfig.cpp:444", "beginArray", "%s: beginArray() on non-object array ""%s""", mFilename, name_
        End If
        Set pCurrentObj = mArray.front()
    End If
End Sub

Public Sub nextArrayItem()
    If mWarning = WzConfig_ReadAndWrite Then
        mArray.push_back pCurrentObj
        If IsLatestNewObj(pCurrentObj) Then
            PopNewObj
        End If
        PushNewObj json_object()
        Set pCurrentObj = mNewObjStack(mNewObjStackSize - 1)
    Else
        If Not mArray.empty_() Then
            mArray.erase_idx 0
            If Not mArray.empty_() Then
                Set pCurrentObj = mArray.front()
            Else
                PushNewObj json_object()
                Set pCurrentObj = mNewObjStack(mNewObjStackSize - 1)
            End If
        End If
    End If
End Sub

Public Function remainingArrayItems() As Long
    remainingArrayItems = mArray.size()
End Function

Public Sub endArray()
    If mWarning = WzConfig_ReadAndWrite Then
        If Not pCurrentObj.empty_() Then
            mArray.push_back pCurrentObj
        End If
        If IsLatestNewObj(pCurrentObj) Then
            PopNewObj
        End If
        Set pCurrentObj = PopObj()
        If Not mArray.empty_() Then
            If Not (Not pCurrentObj Is Nothing) Then
                ASSERT_FAILURE "pCurrentObj != nullptr", "lib/framework/wzconfig.cpp:500", "endArray", "pCurrentObj is null"
            End If
            pCurrentObj.set_item_ref mName, mArray      ' a copy; mArray is replaced below: taken as it is
        End If
        mName = PopName()
    Else
        If IsLatestNewObj(pCurrentObj) Then
            PopNewObj
        End If
        mName = PopName()
        Set pCurrentObj = PopObj()
    End If
    Set mArray = json_array()
End Sub

' setValue(key, json&&) and setValue(key, const json&): a copy of value
Public Sub setValue(ByVal key As String, ByVal value_ As nlohmann_json)
    If Not (Not pCurrentObj Is Nothing) Then
        ASSERT_FAILURE "pCurrentObj != nullptr", "lib/framework/wzconfig.cpp:528", "setValue", "pCurrentObj is null"
    End If
    pCurrentObj.set_item key, value_
End Sub

Public Sub set_(ByVal key As String, ByVal value_ As nlohmann_json)
    setValue key, value_
End Sub

Public Function fileName() As String
    fileName = mFilename
End Function

Public Function isWritable() As Boolean
    isWritable = (mWarning = WzConfig_ReadAndWrite And mStatus)
End Function

Public Function group() As String
    group = mName
End Function

Public Function status() As Boolean
    status = mStatus
End Function
