VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.h: the fields of class OptionsForm
' (VB) data of widget class OptionsForm
Option Explicit

Public optionDetailsMap As Object        ' std::unordered_map<WzString, OptionDetails>: optionId -> index in optionDetails
Private optionDetails_() As OptionsForm_OptionDetails
Public optionDetailsCount As Long
Public sectionDetailsMap As Object       ' std::unordered_map<WzString, SectionDetails>: sectionId -> index in sectionDetails
Private sectionDetails_() As OptionsForm_SectionDetails
Public sectionDetailsCount As Long
Private orderedSectionIds_() As String   ' std::vector<WzString>
Public orderedSectionIdsCount As Long
Public lastConfiguredTranslationLocale As String
Public optionsList As WIDGET             ' std::shared_ptr<ScrollableListWidget>
Public has_mouseOverFormPosThisFrame As Boolean ' optional<Vector2i> mouseOverFormPosThisFrame = nullopt
Public mouseOverFormPosThisFrameX As Long
Public mouseOverFormPosThisFrameY As Long
Public priorMouseOverRowWidget As WIDGET
Public capturedEditingRowWidget As WIDGET
Public currentHelpPopoverWidget As WIDGET ' std::shared_ptr<PopoverWidget>
Public maxListWidth As Long              ' = 0
Public refreshOptionsOnScreenSizeChange As Boolean ' = false
Public isInRefreshOptionsInternal As Boolean ' = false
Public weakSerial As Long               ' (VB) its std::weak_ptr serial (widgweak_vb)

Friend Property Get optionDetails(ByVal i As Long) As OptionsForm_OptionDetails
    optionDetails = optionDetails_(i)
End Property

Friend Property Let optionDetails(ByVal i As Long, v As OptionsForm_OptionDetails)
    optionDetails_(i) = v
End Property

' (VB) push_back: the new element's index
Friend Function optionDetails_push_back(v As OptionsForm_OptionDetails) As Long
    If optionDetailsCount = 0 Then
        ReDim optionDetails_(0 To 7)
    ElseIf optionDetailsCount > UBound(optionDetails_) Then
        ReDim Preserve optionDetails_(0 To optionDetailsCount * 2 - 1)
    End If
    optionDetails_(optionDetailsCount) = v
    optionDetails_push_back = optionDetailsCount
    optionDetailsCount = optionDetailsCount + 1
End Function

Friend Property Get sectionDetails(ByVal i As Long) As OptionsForm_SectionDetails
    sectionDetails = sectionDetails_(i)
End Property

Friend Property Let sectionDetails(ByVal i As Long, v As OptionsForm_SectionDetails)
    sectionDetails_(i) = v
End Property

Friend Function sectionDetails_push_back(v As OptionsForm_SectionDetails) As Long
    If sectionDetailsCount = 0 Then
        ReDim sectionDetails_(0 To 7)
    ElseIf sectionDetailsCount > UBound(sectionDetails_) Then
        ReDim Preserve sectionDetails_(0 To sectionDetailsCount * 2 - 1)
    End If
    sectionDetails_(sectionDetailsCount) = v
    sectionDetails_push_back = sectionDetailsCount
    sectionDetailsCount = sectionDetailsCount + 1
End Function

Public Property Get orderedSectionIds(ByVal i As Long) As String
    orderedSectionIds = orderedSectionIds_(i)
End Property

Public Sub orderedSectionIds_push_back(ByVal s As String)
    If orderedSectionIdsCount = 0 Then
        ReDim orderedSectionIds_(0 To 7)
    ElseIf orderedSectionIdsCount > UBound(orderedSectionIds_) Then
        ReDim Preserve orderedSectionIds_(0 To orderedSectionIdsCount * 2 - 1)
    End If
    orderedSectionIds_(orderedSectionIdsCount) = s
    orderedSectionIdsCount = orderedSectionIdsCount + 1
End Sub
