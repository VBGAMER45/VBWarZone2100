VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMapLoadDebugLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/init.cpp: class WzMapLoadDebugLogger, the WzMap::LoggingProtocol buildMapList gives wzmaplib while it reads
' the map packages: errors reach the log as errors only while setLogErrors(true).
Option Explicit

Implements LoggingProtocol

Private m_logErrors As Boolean

Private Sub LoggingProtocol_printLog(ByVal level As LogLevel, ByVal function_ As String, ByVal line_ As Long, ByVal str_ As String)
    Dim logPart As code_part
    logPart = LOG_MAP
    Select Case level
        Case LogLevel_Info_Verbose
            logPart = LOG_NEVER
        Case LogLevel_Info
            logPart = LOG_NEVER
        Case LogLevel_Warning
            logPart = LOG_MAP
        Case LogLevel_Error
            If m_logErrors Then logPart = LOG_ERROR Else logPart = LOG_MAP
    End Select
    ' if (enabled_debug[logPart]) _debug(line, logPart, function, "%s", str): debug_ checks the part
    debug_ line_, logPart, function_, "%s", str_
End Sub

Public Sub setLogErrors(ByVal enabled As Boolean)
    m_logErrors = enabled
End Sub
