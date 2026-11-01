VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGuideTopic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/screens/guidescreen.cpp: struct WzGuideTopic (a std::shared_ptr<WzGuideTopic> is this object; guidescreen_c).
' A WzJsonLocalizedString holds its getLocalizedString(), picked when the topic is loaded (modinfo_c's
' convention: translations are not ported); a std::vector is a Collection (1-based) or an array and its count.
Option Explicit

Public identifier As String
Public has_sortKey As Boolean                ' optional<std::string> sortKey
Public sortKey As String
Public version As String
Public displayName As String                 ' WzJsonLocalizedString
Public has_applicableModes As Boolean        ' optional<std::unordered_set<ActivitySink::GameMode>> applicableModes
Public applicableModes As Long               ' (VB) the set: bit 2^mode for each mode in it
Public applicableModesCount As Long          ' (VB) its size
Private m_visuals() As WzGuideVisual         ' std::vector<WzGuideVisual> visuals
Public visualsCount As Long
Public contents As Collection                ' std::vector<WzJsonLocalizedString>
Public linkedTopicIdentifiers As Collection  ' std::vector<std::string>

Private Sub Class_Initialize()
    Set contents = New Collection
    Set linkedTopicIdentifiers = New Collection
End Sub

' visuals[i] (i from 0)
Friend Function visual(ByVal i As Long) As WzGuideVisual
    visual = m_visuals(i)
End Function

' visuals.push_back(v)
Friend Sub addVisual(v As WzGuideVisual)
    ReDim Preserve m_visuals(0 To visualsCount)
    m_visuals(visualsCount) = v
    visualsCount = visualsCount + 1
End Sub
