Attribute VB_Name = "nettests_stubs"
' tools/net_test.py's test exe (NetTests.exe): what WzNet's connection modules take from WzFramework (debug.h's
' code_part, debug_, debug__, ASSERT_FAILURE), here writing every log line into the results (all parts on), so
' the test can check the messages.
Option Explicit

' debug.h (FrameworkApi's enum)
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
    LOG_MOVEMENT
    LOG_ATTACK
    LOG_FOG
    LOG_SENSOR
    LOG_GUI
    LOG_MAP
    LOG_SAVE
    LOG_SYNC
    LOG_DEATH
    LOG_LIFE
    LOG_GATEWAY
    LOG_MSG
    LOG_INFO
    LOG_TERRAIN
    LOG_FEATURE
    LOG_FATAL
    LOG_INPUT
    LOG_POPUP
    LOG_CONSOLE
    LOG_ACTIVITY
    LOG_RESEARCH
    LOG_SAVEGAME
    LOG_REPAIRS
    LOG_LOBBY
    LOG_LAST
End Enum

Public asserts As Long

Public Sub debug_(ByVal line_ As Long, ByVal part As code_part, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    nettests_main.addResult "log" & vbTab & part & vbTab & function_ & ":" & line_ & vbTab & vsprintf(str_, a)
End Sub

Public Sub debug__(ByVal line_ As Long, ByVal part As code_part, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    nettests_main.addResult "log" & vbTab & part & vbTab & function_ & ":" & line_ & vbTab & vsprintf(str_, a)
End Sub

Public Sub ASSERT_FAILURE(ByVal expr_string As String, ByVal location_description As String, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    asserts = asserts + 1
    nettests_main.addResult "assert" & vbTab & expr_string & vbTab & location_description & vbTab & function_ & vbTab & vsprintf(str_, a)
End Sub
