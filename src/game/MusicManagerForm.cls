VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicManagerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class MusicManagerForm
' src/musicmanager.cpp: the fields of class MusicManagerForm (musicmanager_c); weakSerial: (VB) its std::weak_ptrs
' (widgweak_vb); the initialisers are its _construct's.
Option Explicit

Public weakSerial As Long
Public trackList As Collection             ' std::vector<std::shared_ptr<const WZ_TRACK>>
Public selectedTrack As WZ_TRACK           ' std::shared_ptr<const WZ_TRACK>
Public table As WIDGET                     ' std::shared_ptr<ScrollableTableWidget>
Public trackDetailsBox As WIDGET           ' std::shared_ptr<TrackDetailsForm>
Public musicManagerAudioEventSink As MusicManager_CDAudioEventSink ' std::shared_ptr<MusicManager_CDAudioEventSink>
Public minDesiredTableWidth As Long        ' int32_t = 0
Public ingame_ As Boolean                  ' bool ingame (VB: _ as the global ingame is stubs_game's)
Public tablePaddingX As Long               ' const int32_t = 6
