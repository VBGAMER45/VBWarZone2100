VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrackDetailsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class TrackDetailsForm
' src/musicmanager.cpp: the fields of class TrackDetailsForm (musicmanager_c); the initialisers are its _construct's.
Option Explicit

Public pAlbumCoverTexture As gfx_api_texture ' gfx_api::texture* = nullptr
Public album_cover_path As String

Public psNowPlaying As WIDGET              ' std::shared_ptr<W_LABEL>
Public psSelectedTrackName As WIDGET       ' std::shared_ptr<W_LABEL>
Public psSelectedTrackAuthorName As WIDGET ' std::shared_ptr<W_LABEL>
Public psSelectedTrackAlbumName As WIDGET  ' std::shared_ptr<W_LABEL>
Public psSelectedTrackAlbumDate As WIDGET  ' std::shared_ptr<W_LABEL>
Public descriptionContainer As WIDGET      ' std::shared_ptr<ScrollableListWidget>

Public outerHorizontalPadding As Long      ' int32_t = 20
