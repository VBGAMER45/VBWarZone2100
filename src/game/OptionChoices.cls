VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.h: template<typename T> struct OptionChoices, for every T
' (VB: the values are Variants; Empty is nullopt). choices is std::vector<OptionChoice<T>>: count
' and the per-choice properties; addChoice is choices.push_back({displayName, helpDescription,
' value, disabled, icon}).
Option Explicit

Private choices_() As OptionChoice
Private count_ As Long
Public currentIdx As Long               ' size_t = 0

Public Property Get count() As Long
    count = count_
End Property

Public Sub addChoice(ByVal displayName As String, ByVal helpDescription As String, ByVal value As Variant, Optional ByVal disabled As Boolean = False, Optional ByVal icon As Long = 0)
    If count_ = 0 Then
        ReDim choices_(0 To 7)
    ElseIf count_ > UBound(choices_) Then
        ReDim Preserve choices_(0 To count_ * 2 - 1)
    End If
    choices_(count_).displayName = displayName
    choices_(count_).helpDescription = helpDescription
    choices_(count_).value = value
    choices_(count_).disabled = disabled
    choices_(count_).icon = icon
    count_ = count_ + 1
End Sub

' (VB) choices.clear()
Public Sub clearChoices()
    count_ = 0
    Erase choices_
End Sub

Public Property Get displayName(ByVal i As Long) As String
    displayName = choices_(i).displayName
End Property

Public Property Let displayName(ByVal i As Long, ByVal v As String)
    choices_(i).displayName = v
End Property

Public Property Get helpDescription(ByVal i As Long) As String
    helpDescription = choices_(i).helpDescription
End Property

Public Property Let helpDescription(ByVal i As Long, ByVal v As String)
    choices_(i).helpDescription = v
End Property

Public Property Get value(ByVal i As Long) As Variant
    value = choices_(i).value
End Property

Public Property Let value(ByVal i As Long, ByVal v As Variant)
    choices_(i).value = v
End Property

Public Property Get disabled(ByVal i As Long) As Boolean
    disabled = choices_(i).disabled
End Property

Public Property Let disabled(ByVal i As Long, ByVal v As Boolean)
    choices_(i).disabled = v
End Property

' AtlasImageDef *icon: a handle (0 = nullptr)
Public Property Get icon(ByVal i As Long) As Long
    icon = choices_(i).icon
End Property

Public Property Let icon(ByVal i As Long, ByVal v As Long)
    choices_(i).icon = v
End Property

' bool setCurrentIdxForValue(const T& value)
Public Function setCurrentIdxForValue(ByVal value As Variant) As Boolean
    Dim i As Long
    For i = 0 To count_ - 1
        If OptionValue_equal(choices_(i).value, value) Then
            currentIdx = i
            setCurrentIdxForValue = True
            Exit Function
        End If
    Next
    setCurrentIdxForValue = False
End Function

' (VB) choices == other.choices (OptionChoice::operator==: the value, displayName, helpDescription and disabled)
Public Function choicesEqual(ByVal other As OptionChoices) As Boolean
    Dim i As Long
    If other.count <> count_ Then Exit Function
    For i = 0 To count_ - 1
        If Not OptionValue_equal(choices_(i).value, other.value(i)) Then Exit Function
        If StrComp(choices_(i).displayName, other.displayName(i), vbBinaryCompare) <> 0 Then Exit Function
        If StrComp(choices_(i).helpDescription, other.helpDescription(i), vbBinaryCompare) <> 0 Then Exit Function
        If choices_(i).disabled <> other.disabled(i) Then Exit Function
    Next
    choicesEqual = True
End Function
