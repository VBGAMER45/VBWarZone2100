VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapScriptLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' MapScriptTests: tools/mapscriptref.cpp's CaptureLogger, a wzmaplib
' LoggingProtocol that adds each log line to the output: "log <level>
' <function>:<line> " and the text, with a line feed when the text has none.
Option Explicit

Implements LoggingProtocol

Private Sub LoggingProtocol_printLog(ByVal level As LogLevel, ByVal function_ As String, ByVal line_ As Long, ByVal str_ As String)
    mapscripttests_main.Put_ "log " & CStr(level) & " " & function_ & ":" & CStr(line_) & " " & str_
    If Not mapscripttests_main.EndsWithLF() Then mapscripttests_main.Put_ vbLf
End Sub
