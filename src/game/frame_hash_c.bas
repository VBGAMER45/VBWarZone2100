Attribute VB_Name = "frame_hash_c"
' lib/framework/frame.cpp's findHashOfFile, in the exe: SHA-256 is src/shared/sodium's, which the exe and WzNet
' compile and WzFramework does not (the levels, the mods and the map transfer hash their files with it).
Option Explicit

' Sha256 findHashOfFile(char const *realFileName) (VB: the result in hash)
Public Sub findHashOfFile(ByVal realFileName As String, hash As Sha256)
    Dim realFileData() As Byte, realFileSize As Long
    If loadFile(realFileName, realFileData, realFileSize) Then
        hash = sha256Sum(realFileData, 0, realFileSize)
        Exit Sub
    End If
    Sha256_setZero hash
End Sub
