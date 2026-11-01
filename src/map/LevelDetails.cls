VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LevelDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map_package.h: struct WzMap::LevelDetails, a map's level details (a .lev file's
' or a level.json's). C++ copies it as a value: copyFrom makes a copy (MapApi.LevelDetails_new for a new one).
' optional<std::string> generator is generator with hasGenerator; the vector of additional authors is
' additionalAuthors(0 .. additionalAuthorsCount - 1).
Option Explicit

Public name_ As String                      ' the level / map name
Public type_ As MapType
Public players As Long                      ' uint8_t: the max number of players the map is designed for
Public tileset As MAP_TILESET               ' the base level data set used for terrain and tiles
Public mapFolderPath As String              ' the path to the folder containing the map
Public createdDate As String
Public uploadedDate As String
Public author As String
Private m_additionalAuthors() As String
Private m_additionalAuthorsCount As Long
Public license As String
Public generator As String                  ' the program / script / library used to generate (or export) the map
Public hasGenerator As Boolean

Private Sub Class_Initialize()
    tileset = MAP_TILESET_ARIZONA
End Sub

Public Function additionalAuthorsCount() As Long
    additionalAuthorsCount = m_additionalAuthorsCount
End Function

Public Function additionalAuthor(ByVal i As Long) As String
    additionalAuthor = m_additionalAuthors(i)
End Function

' additionalAuthors.push_back(a)
Public Sub additionalAuthors_push_back(ByVal a As String)
    ReDim Preserve m_additionalAuthors(0 To m_additionalAuthorsCount)
    m_additionalAuthors(m_additionalAuthorsCount) = a
    m_additionalAuthorsCount = m_additionalAuthorsCount + 1
End Sub

' (VB) *this = other
Public Sub copyFrom(ByVal other As LevelDetails)
    Dim i As Long
    name_ = other.name_
    type_ = other.type_
    players = other.players
    tileset = other.tileset
    mapFolderPath = other.mapFolderPath
    createdDate = other.createdDate
    uploadedDate = other.uploadedDate
    author = other.author
    Erase m_additionalAuthors
    m_additionalAuthorsCount = 0
    For i = 0 To other.additionalAuthorsCount() - 1
        additionalAuthors_push_back other.additionalAuthor(i)
    Next
    license = other.license
    generator = other.generator
    hasGenerator = other.hasGenerator
End Sub

' Returns the ".gam" file path
' For an old-fashioned map, this should be a file that exists
' For a newer map, this points to a "virtual" .gam file that does not exist (but is in the parent folder of the map folder, with the same map name .gam)
Public Function gamFilePath() As String
    Dim gamFilePath_ As String
    gamFilePath_ = gamFilePathFromMapFolderPath(mapFolderPath, "/")
    If Len(gamFilePath_) > 0 Then
        gamFilePath = gamFilePath_
    Else
        ' not sure if this can happen, but return something reasonable at least...
        If Len(name_) > 0 Then
            gamFilePath = name_ & ".gam"
        Else
            gamFilePath = "map.gam"
        End If
    End If
End Function
