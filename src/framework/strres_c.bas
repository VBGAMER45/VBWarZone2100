Attribute VB_Name = "strres_c"
' lib/framework/strres.cpp: string resources, a treap (treap_c) of strings
' by identifier, loaded from text files by the strres parser
' (strres_parser_c, strres_lexer_c). struct STR_RES * is the class STR_RES
' (Nothing: nullptr); a const char * result that may be nullptr is
' vbNullString. abort() is exit_ 3.
Option Explicit

' Initialise the string system
Public Function strresCreate() As STR_RES
    Dim psRes As STR_RES
    Set psRes = New STR_RES

    psRes.psIDTreap = treapCreate()
    If psRes.psIDTreap = 0 Then
        debug_ 60, LOG_FATAL, "strresCreate", "Out of memory"
        exit_ 3                         ' abort()
        Set strresCreate = Nothing
        Exit Function
    End If

    Set strresCreate = psRes
End Function

' Shutdown the string system
Public Sub strresDestroy(ByVal psRes As STR_RES)
    ' Release the treap and free the final memory
    treapDestroy psRes.psIDTreap
    psRes.psIDTreap = 0
End Sub

' Store a string
Public Function strresStoreString(ByVal psRes As STR_RES, ByVal pID As String, ByVal pString As String) As Boolean
    ' Make sure that this ID string hasn't been used before
    If StrPtr(treapFind(psRes.psIDTreap, pID)) <> 0 Then
        debug_ 84, LOG_FATAL, "strresStoreString", "Duplicate string for id: ""%s""", pID
        exit_ 3                         ' abort()
        strresStoreString = False
        Exit Function
    End If

    strresStoreString = treapAdd(psRes.psIDTreap, pID, pString)
End Function

Public Function strresGetString(ByVal psRes As STR_RES, ByVal ID As String) As String
    strresGetString = treapFind(psRes.psIDTreap, ID)
End Function

' Load a string resource file
Public Function strresLoad(ByVal psRes As STR_RES, ByVal fileName As String) As Boolean
    Dim retval As Boolean, physfsfile As Long

    ' lexerinput_t input: LEXINPUT_PHYSFS
    physfsfile = PHYSFS_openRead(fileName)
    debug_ 105, LOG_WZ, "strresLoad", "Reading...[directory %s] %s", PHYSFS_getRealDir(fileName) & "", fileName
    If physfsfile = 0 Then
        debug_ 108, LOG_ERROR, "strresLoad", "strresLoadFile: PHYSFS_openRead(""%s"") failed with error: %s\n", fileName, WZ_PHYSFS_getLastError()
        strresLoad = False
        Exit Function
    End If

    strres_set_extra physfsfile
    retval = (strres_parse(psRes) = 0)

    strres_lex_destroy
    PHYSFS_close physfsfile

    strresLoad = retval
End Function

' Get the ID number for a string
Public Function strresGetIDfromString(ByVal psRes As STR_RES, ByVal pString As String) As String
    strresGetIDfromString = treapFindKey(psRes.psIDTreap, pString)
End Function
