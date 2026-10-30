Attribute VB_Name = "map_package_c"
' lib/wzmaplib/src/map_package.cpp (ported in part): trimTechLevelFromMapName.
' Missing: the level details (.lev / JSON), map packages, the rest.
Option Explicit

' "<name>-T<digit>" -> "<name>" (isdigit: the C locale's 0-9), in place
Public Sub trimTechLevelFromMapName(mapName As String)
    Dim length As Long, c As Long
    length = Len(mapName)
    If length > 2 Then
        c = AscW(Mid$(mapName, length, 1))
        If Mid$(mapName, length - 2, 1) = "-" And Mid$(mapName, length - 1, 1) = "T" And c >= 48 And c <= 57 Then
            mapName = Left$(mapName, length - 3)
        End If
    End If
End Sub
