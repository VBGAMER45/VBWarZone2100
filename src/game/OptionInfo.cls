VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.cpp: class OptionInfo, an option title and other information
' (OptionInfo_new(optionId, displayName, helpDescription), optionsform_c). The chained setters
' return the object itself (C++'s OptionInfo&); an OptionAvailabilityCondition is a WzFunc called
' with the OptionInfo, returning an OptionInfo_AvailabilityResult.
Option Explicit

Public optionId As String
Private displayName As String
Private helpDescription As String
Private availabilityConditions As Collection    ' std::vector<OptionAvailabilityCondition>
Private bRequiresRestart As Boolean             ' = false

' OptionInfo(const WzString& optionId, const WzString& displayName, const WzString& helpDescription)
Friend Sub construct(ByVal optionId_ As String, ByVal displayName_ As String, ByVal helpDescription_ As String)
    optionId = optionId_
    displayName = displayName_
    helpDescription = helpDescription_
    Set availabilityConditions = New Collection
    bRequiresRestart = False
End Sub

Public Function addAvailabilityCondition(ByVal condition As WzFunc) As OptionInfo
    availabilityConditions.Add condition
    Set addAvailabilityCondition = Me
End Function

Public Function setRequiresRestart(ByVal val_ As Boolean) As OptionInfo
    bRequiresRestart = val_
    Set setRequiresRestart = Me
End Function

Public Function getTranslatedDisplayName() As String
    If Len(displayName) = 0 Then
        getTranslatedDisplayName = ""
        Exit Function
    End If
    getTranslatedDisplayName = displayName      ' gettext (translations not ported)
End Function

Public Function getTranslatedHelpDescription() As String
    If Len(helpDescription) = 0 Then
        getTranslatedHelpDescription = ""
        Exit Function
    End If
    getTranslatedHelpDescription = helpDescription
End Function

Public Function isAvailable() As Boolean
    Dim condition As WzFunc, result As OptionInfo_AvailabilityResult, self As OptionInfo
    Set self = Me                    ' (VB) WzFunc_call takes a copy of Me (CONVENTIONS, "Widgets")
    For Each condition In availabilityConditions
        Set result = WzFunc_call(condition, self)
        If Not result.available_ Then
            isAvailable = False
            Exit Function
        End If
    Next
    isAvailable = True
End Function

' std::vector<AvailabilityResult> getAvailabilityResults() const: the count, the results in result(0 .. count - 1)
Public Function getAvailabilityResults(result() As OptionInfo_AvailabilityResult) As Long
    Dim condition As WzFunc, n As Long, self As OptionInfo
    Set self = Me
    n = 0
    For Each condition In availabilityConditions
        If n = 0 Then ReDim result(0 To 0) Else ReDim Preserve result(0 To n)
        Set result(n) = WzFunc_call(condition, self)
        n = n + 1
    Next
    getAvailabilityResults = n
End Function

Public Function requiresRestart() As Boolean
    requiresRestart = bRequiresRestart
End Function
