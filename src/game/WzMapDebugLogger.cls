VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMapDebugLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/map.cpp, map.h: class WzMapDebugLogger, the WzMap::LoggingProtocol
' through which wzmaplib's messages reach the game's log, under the
' library's function name and line.
Option Explicit

Implements LoggingProtocol

Private Sub LoggingProtocol_printLog(ByVal level As LogLevel, ByVal function_ As String, ByVal line_ As Long, ByVal str_ As String)
    Dim logPart As code_part
    logPart = LOG_INFO
    Select Case level
    Case LogLevel_Info_Verbose
        logPart = LOG_NEVER
    Case LogLevel_Info
        logPart = LOG_MAP
    Case LogLevel_Warning
        logPart = LOG_WARNING
    Case LogLevel_Error
        logPart = LOG_ERROR
    End Select
    ' if (enabled_debug[logPart]) _debug(line, logPart, function, "%s", str): debug_ checks the part
    debug_ line_, logPart, function_, "%s", str_
End Sub
