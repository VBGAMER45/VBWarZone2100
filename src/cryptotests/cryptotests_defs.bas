Attribute VB_Name = "cryptotests_defs"
' (VB) CryptoTests: what src/shared/sodium uses from WzFramework's API class (FrameworkApi.cls's code_part,
' debug_, ASSERT_FAILURE), stood in so the test exe compiles the modules without the DLL: each log or assert
' becomes a "LOG" or "ASSERT" line of the results (cryptotests_main's emit).
Option Explicit

Public Enum code_part
    LOG_ALL
    LOG_MAIN
    LOG_SOUND
    LOG_VIDEO
    LOG_WZ
    LOG_3D
    LOG_TEXTURE
    LOG_NET
    LOG_MEMORY
    LOG_WARNING
    LOG_ERROR
    LOG_NEVER
    LOG_SCRIPT
    LOG_SAVE
    LOG_SYNC
    LOG_DEATH
    LOG_LIFE
    LOG_GATEWAY
    LOG_MSG
    LOG_MAP
    LOG_SAVEGAME
    LOG_SCREEN
    LOG_INFO
    LOG_TERRAIN
    LOG_FEATURE
    LOG_FPATH
    LOG_POPUP
    LOG_CONSOLE
    LOG_ACTIVITY
    LOG_RESEARCH
    LOG_FATAL
    LOG_INPUT
    LOG_LAST
End Enum

Public Sub debug_(ByVal line_ As Long, ByVal part As code_part, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim s As String, i As Long
    s = "LOG" & vbTab & line_ & vbTab & part & vbTab & function_ & vbTab & str_
    For i = 0 To UBound(args)
        s = s & vbTab & CStr(args(i))
    Next
    testLog s
End Sub

Public Sub ASSERT_FAILURE(ByVal expr_string As String, ByVal location_description As String, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    testLog "ASSERT" & vbTab & location_description & vbTab & function_ & vbTab & str_
End Sub
