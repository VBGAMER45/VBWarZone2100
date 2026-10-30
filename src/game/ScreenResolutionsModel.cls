VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScreenResolutionsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/windowoptions.cpp: class ScreenResolutionsModel (its static functions:
' windowoptions_c's ScreenResolutionsModel_<name>). modes is std::vector<optional<screeninfo>>:
' modes_(i) with has_(i); an iterator is an index (count is end()).
Option Explicit

Private modes_() As screeninfo
Private has_() As Boolean
Private count_ As Long

' ScreenResolutionsModel(): modes(ScreenResolutionsModel::loadModes())
Friend Sub construct()
    loadModes
End Sub

Public Property Get count() As Long
    count = count_
End Property

' (VB) ScreenResolutionsModel::resolutionString(modes[i])
Public Function resolutionStringAt(ByVal i As Long) As String
    resolutionStringAt = ScreenResolutionsModel_resolutionString(has_(i), modes_(i))
End Function

' const_iterator findResolutionClosestToCurrent() const: an index
Public Function findResolutionClosestToCurrent() As Long
    Dim currentResolution As screeninfo, has_current As Boolean, closest As Long
    has_current = ScreenResolutionsModel_getCurrentResolution(currentResolution)
    ' std::lower_bound(modes.begin(), modes.end(), currentResolution, compareLess)
    closest = 0
    Do While closest < count_
        If Not ScreenResolutionsModel_compareLess(has_(closest), modes_(closest), has_current, currentResolution) Then Exit Do
        closest = closest + 1
    Loop
    If closest <> count_ Then
        If ScreenResolutionsModel_compareEq(has_(closest), modes_(closest), has_current, currentResolution) Then
            findResolutionClosestToCurrent = closest
            Exit Function
        End If
    End If

    If closest <> 0 Then
        closest = closest - 1   ' If current resolution doesn't exist, round down to next-highest one.
    End If

    findResolutionClosestToCurrent = closest
End Function

Public Function selectAt(ByVal index As Long) As Boolean
    Dim selectedResolution As screeninfo, has_selected As Boolean, currentResolution As screeninfo, has_current As Boolean
    If Not (index < count_) Then ASSERT_FAILURE "index < modes.size()", "src/titleui/options/windowoptions.cpp:82", "selectAt", "Invalid mode index passed to ScreenResolutionsModel::selectAt": selectAt = False: Exit Function

    selectedResolution = modes_(index)
    has_selected = has_(index)

    has_current = ScreenResolutionsModel_getCurrentResolution(currentResolution)
    ' Attempt to change the resolution
    If Not wzChangeFullscreenDisplayMode(has_selected, selectedResolution) Then
        debug_ 90, LOG_WARNING, "selectAt", "Failed to change fullscreen mode from: %s to: %s", ScreenResolutionsModel_resolutionString(has_current, currentResolution), ScreenResolutionsModel_resolutionString(has_selected, selectedResolution)
        selectAt = False
        Exit Function
    End If

    ' NOTE: Settings are saved by wzChangeFullscreenDisplayMode on success

    selectAt = True
End Function

' static std::vector<optional<screeninfo>> loadModes()
Private Sub loadModes()
    Dim i As Long, j As Long, tmp As screeninfo, tmpHas As Boolean, n As Long
    ' Get resolutions, sorted with duplicates removed.
    count_ = wzAvailableResolutions(modes_, has_)
    ' std::sort(modes.begin(), modes.end(), ScreenResolutionsModel::compareLess) (an insertion sort)
    For i = 1 To count_ - 1
        tmp = modes_(i)
        tmpHas = has_(i)
        j = i - 1
        Do While j >= 0
            If Not ScreenResolutionsModel_compareLess(tmpHas, tmp, has_(j), modes_(j)) Then Exit Do
            modes_(j + 1) = modes_(j)
            has_(j + 1) = has_(j)
            j = j - 1
        Loop
        modes_(j + 1) = tmp
        has_(j + 1) = tmpHas
    Next
    ' modes.erase(std::unique(modes.begin(), modes.end(), ScreenResolutionsModel::compareEq), modes.end())
    If count_ > 0 Then
        n = 1
        For i = 1 To count_ - 1
            If Not ScreenResolutionsModel_compareEq(has_(n - 1), modes_(n - 1), has_(i), modes_(i)) Then
                modes_(n) = modes_(i)
                has_(n) = has_(i)
                n = n + 1
            End If
        Next
        count_ = n
    End If
End Sub
