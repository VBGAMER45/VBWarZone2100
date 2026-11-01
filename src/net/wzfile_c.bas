Attribute VB_Name = "wzfile_c"
' lib/netplay/wzfile.cpp: physfs_file_safe_close (the class WZFile is WZFile.cls). PHYSFS_isInit is not in
' WzFramework's API: the close is not skipped after PhysFS's shutdown (a WZFile still open then is released at
' the DLL's unload, after the game closed its files).
Option Explicit

Public Function physfs_file_safe_close_impl(ByVal f As Long) As Boolean
    If f = 0 Then
        physfs_file_safe_close_impl = False
        Exit Function
    End If
    PHYSFS_close f
    physfs_file_safe_close_impl = True
End Function

Public Sub physfs_file_safe_close(ByVal f As Long)
    physfs_file_safe_close_impl f
End Sub

' WZFile(PHYSFS_file* handle, const std::string& filename, Sha256 hash, uint32_t size = 0): hashBytes(0 .. 31)
Public Function WZFile_new(ByVal handle As Long, ByVal filename As String, hashBytes() As Byte, Optional ByVal size As Long = 0) As WZFile
    Dim f As New WZFile
    f.construct handle, filename, hashBytes, size
    Set WZFile_new = f
End Function
