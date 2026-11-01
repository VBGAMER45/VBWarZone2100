VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzWrappedGuideTopic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/screens/guidescreen.cpp: class WzWrappedGuideTopic (a std::shared_ptr<WzWrappedGuideTopic> is this object;
' guidescreen_c). The std::weak_ptr parent is the parent's identifier, locked through the registry's lookup
' (WzGuideTopicsRegistry_getTopic): a parent that is gone from the registry has expired.
' optional<ActivityDBProtocol::GuideTopicPrefs> prefs: has_prefs and the struct's two fields.
Option Explicit

Public topic As WzGuideTopic                 ' std::shared_ptr<WzGuideTopic>
Public parentId As String                    ' std::weak_ptr<WzWrappedGuideTopic> parent ("": none)
Public children As Collection                ' std::vector<std::shared_ptr<WzWrappedGuideTopic>>
Public has_prefs As Boolean                  ' prefs: loaded on-demand, used by WzGuideTopicsRegistry only
Public prefs_has_lastReadVersion As Boolean  ' GuideTopicPrefs::lastReadVersion (optional<std::string>)
Public prefs_lastReadVersion As String
Public prefs_bookmarked As Boolean           ' GuideTopicPrefs::bookmarked = false

Private Sub Class_Initialize()
    Set children = New Collection
End Sub
