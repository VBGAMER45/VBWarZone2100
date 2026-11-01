Attribute VB_Name = "netjoin_c"
' lib/netplay/netjoin.cpp, netjoin.h: struct JoinConnectionDescription's connection type names (the struct
' itself, host, port and type, is the game's: it only crosses into WzNet as those three values).
' JoinConnectionType: TCP_DIRECT (0) only (no GameNetworkingSockets backend).
Option Explicit

Public Const JoinConnectionType_TCP_DIRECT As Long = 0

Public Function JoinConnectionDescription_connectiontype_to_string(ByVal typ As Long) As String
    Select Case typ
    Case JoinConnectionType_TCP_DIRECT: JoinConnectionDescription_connectiontype_to_string = "tcp"
    Case Else: JoinConnectionDescription_connectiontype_to_string = ""   ' silence compiler warning
    End Select
End Function

' optional<JoinConnectionType> connectiontype_from_string(const std::string& str): has_value() is the result
Public Function JoinConnectionDescription_connectiontype_from_string(ByVal str_ As String, typ As Long) As Boolean
    If str_ = "tcp" Then
        typ = JoinConnectionType_TCP_DIRECT
        JoinConnectionDescription_connectiontype_from_string = True
        Exit Function
    End If
    JoinConnectionDescription_connectiontype_from_string = False
End Function
