Attribute VB_Name = "text_c"
' src/text.cpp, text.h: the game's string resource (WzFramework's strres).
Option Explicit

' The string resource object
Public psStringRes As STR_RES

' Initialise the string system
Public Function stringsInitialise() As Boolean
    Set psStringRes = strresCreate()

    stringsInitialise = Not psStringRes Is Nothing
End Function


' Shut down the string system
Public Sub stringsShutDown()
    strresDestroy psStringRes
End Sub
