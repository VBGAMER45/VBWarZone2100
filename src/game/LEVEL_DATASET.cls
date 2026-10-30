VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LEVEL_DATASET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/levels.h: struct LEVEL_DATASET, a level (or a dataset) of the level
' descriptions; the LEVEL_DATASET* pointers are references to these objects.
' type is typ; std::array<std::string, LEVEL_MAXFILES> apDataFiles is
' apDataFiles(i); game and dataDir are game_ and dataDir_ (the global game
' and main.cpp's datadir); the char* realFileName and customMountPoint are
' vbNullString for nullptr, so they are properties: a Public String field
' assigned from outside its class stores vbNullString as "" (VB's implicit
' Property Let copies it); realFileHash (a Sha256) is a Friend property.
' Class_Initialize gives the C++ member initialisers.
Option Explicit

Public typ As Long                      ' LEVEL_TYPE: type of map
Public players As Long                  ' SWORD: number of players for the map
Public game_ As Long                     ' SWORD: index of WRF/WDG that loads the scenario file
Public pName As String                  ' title for the level
Public dataDir_ As Long                  ' searchPathMode: title for the level
Private apDataFiles_(0 To LEVEL_MAXFILES - 1) As String ' the WRF/GAM files for the level
                                        ' in load order
Public psBaseData As LEVEL_DATASET      ' LEVEL_DATASET that must be loaded for this level to load
Public psChange As LEVEL_DATASET        ' LEVEL_DATASET used when changing to this level from another
Private realFileName_ As String         ' Filename of the file containing the level, or NULL if the level is built in.
Private realFileHash_ As Sha256         ' Use levGetFileHash() to read this value. SHA-256 hash of the file containing the level, or 0x00x32 if the level is built in or not yet calculated.
Private customMountPoint_ As String     ' A custom mount point (to be used for "flattened" map packages, or NULL for the default.

Private Sub Class_Initialize()
    typ = LEVEL_TYPE_LDS_COMPLETE
    players = 0
    game_ = 0
    dataDir_ = mod_clean
    realFileName_ = vbNullString
    customMountPoint_ = vbNullString
End Sub

Public Property Get realFileName() As String
    realFileName = realFileName_
End Property

Public Property Let realFileName(ByVal v As String)
    realFileName_ = v
End Property

Public Property Get customMountPoint() As String
    customMountPoint = customMountPoint_
End Property

Public Property Let customMountPoint(ByVal v As String)
    customMountPoint_ = v
End Property

Public Property Get apDataFiles(ByVal i As Long) As String
    apDataFiles = apDataFiles_(i)
End Property

Public Property Let apDataFiles(ByVal i As Long, ByVal v As String)
    apDataFiles_(i) = v
End Property

Friend Property Get realFileHash() As Sha256
    realFileHash = realFileHash_
End Property

Friend Property Let realFileHash(v As Sha256)
    realFileHash_ = v
End Property

Public Sub reset()
    Dim i As Long
    typ = 0
    players = 0
    game_ = 0
    pName = ""
    dataDir_ = mod_clean
    For i = 0 To LEVEL_MAXFILES - 1
        apDataFiles_(i) = ""
    Next
    Set psBaseData = Nothing
    Set psChange = Nothing
    realFileName_ = vbNullString
    Sha256_setZero realFileHash_
    customMountPoint_ = vbNullString
End Sub
