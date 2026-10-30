VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzLobbyMultiOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/multilobbyoptions.cpp: the fields of class WzLobbyMultiOptionButton
' (VB) data of widget class WzLobbyMultiOptionButton
Option Explicit

Public mode As Long                     ' MultiOptionButtonMode mode = MultiOptionButtonMode::HOST_CONFIGURATION;
Public voteTotal As Long                ' int voteTotal = 0;
Public showSelected As Boolean
Public showTotalVotes As Boolean
Public localVoteIndicator As Boolean
Public showLocalVoteIndicatorOnHighlight As Boolean
Public voteIndicatorColor As Long       ' PIELIGHT voteIndicatorColor = pal_RGBA(13, 255, 0, 255);
Public voteIndicatorDimensions As Long  ' int voteIndicatorDimensions = 12;
Public voteTotalText As WzCachedText
