Attribute VB_Name = "test_wzconfig"
' WzTests: WzConfig on a data file, through WzFramework's PhysFS (the
' file's folder mounted, the write dir <output folder>\wz).
'   wzread  every group's keys: group TAB key TAB value(key).toWzString(ok)
'           (ok false: "<not ok>"; ok untouched: "<untouched>")
'   wzcopy  each group copied key by key (json(key)) into a ReadAndWrite
'           WzConfig, whose destructor writes the file: its text
Option Explicit

Private physfsReady As Boolean

Private Sub MountFolderOf(ByVal inPath As String, ByVal root As String, name_ As String)
    Dim folder As String, sp() As String, n As Long, i As Long, writeDir As String
    folder = Left$(inPath, InStrRev(inPath, "\"))
    name_ = Mid$(inPath, InStrRev(inPath, "\") + 1)
    If Not physfsReady Then
        PHYSFS_init vbNullString
        writeDir = root & "\wz\"
        If Len(Dir$(writeDir, vbDirectory)) = 0 Then MkDir writeDir
        PHYSFS_setWriteDir writeDir
        physfsReady = True
    End If
    n = PHYSFS_getSearchPath(sp)
    For i = 0 To n - 1
        PHYSFS_unmount sp(i)
    Next
    PHYSFS_mount folder, vbNullString, PHYSFS_APPEND
End Sub

Public Function RunWzConfig(ByVal mode As String, ByVal inPath As String, ByVal root As String) As String
    Dim name_ As String, ini As WzConfig, dst As WzConfig, groups() As String, keys() As String
    Dim ng As Long, nk As Long, g As Long, k As Long, s As String, ok As Variant, out As String, b() As Byte, n As Long

    MountFolderOf inPath, root, name_
    Select Case mode
    Case "wzread"
        Set ini = WzConfig_new(name_, WzConfig_ReadOnly)
        ng = ini.childGroups(groups)
        For g = 0 To ng - 1
            ini.beginGroup groups(g)
            nk = ini.childKeys(keys)
            For k = 0 To nk - 1
                ok = Empty
                s = ini.value(keys(k)).toWzString(ok)
                If IsEmpty(ok) Then
                    s = "<untouched>"
                ElseIf Not ok Then
                    s = "<not ok>"
                End If
                out = out & groups(g) & vbTab & keys(k) & vbTab & s & vbLf
            Next
            ini.endGroup
        Next
        RunWzConfig = out
    Case "wzcopy"
        Set ini = WzConfig_new(name_, WzConfig_ReadOnly)
        Set dst = WzConfig_new("copy_" & name_, WzConfig_ReadAndWrite)
        ng = ini.childGroups(groups)
        For g = 0 To ng - 1
            ini.beginGroup groups(g)
            dst.beginGroup groups(g)
            nk = ini.childKeys(keys)
            For k = 0 To nk - 1
                dst.setValue keys(k), ini.json(keys(k))
            Next
            dst.endGroup
            ini.endGroup
        Next
        Set dst = Nothing                       ' ~WzConfig: saves
        n = tests_main.ReadBytes(root & "\wz\copy_" & name_, b)
        RunWzConfig = utf8_decode(b, 0, n)
        Kill root & "\wz\copy_" & name_
    End Select
End Function
