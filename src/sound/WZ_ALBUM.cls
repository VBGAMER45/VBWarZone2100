VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZ_ALBUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' cdaudio.h: struct WZ_ALBUM, a music album (music/albums/<dir>/album.json). PublicNotCreatable: the game's
' music manager reads them through their tracks. date is date_ (a VB keyword); std::vector<std::shared_ptr
' <WZ_TRACK>> tracks is a Collection of WZ_TRACKs.
Option Explicit

Public title As String
Public author As String
Public date_ As String
Public description As String
Public album_cover_filename As String
Public tracks As Collection

Private Sub Class_Initialize()
    Set tracks = New Collection
End Sub
