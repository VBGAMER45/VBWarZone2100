VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoggingProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map_debug.h, lib/wzmaplib/src/map_debug.cpp:
' class WzMap::LoggingProtocol, through which wzmaplib logs. A user of the
' library implements it (Implements LoggingProtocol) and passes the object
' to Map_loadFromPath and the other functions; Nothing logs nothing. The
' body of printLog is the C++ base class's default (printed to stdout).
' LogLevel is LoggingProtocol::LogLevel; the library's LOG_INFO_VERBOSE,
' LOG_INFO, LOG_WARNING (and LOG_SYNTAX_WARNING) and LOG_ERROR are its
' members (map_internal.h).
Option Explicit

Public Enum LogLevel
    LogLevel_Info_Verbose = 0
    LogLevel_Info = 1
    LogLevel_Warning = 2
    LogLevel_Error = 3
End Enum

Public Sub printLog(ByVal level As LogLevel, ByVal function_ As String, ByVal line_ As Long, ByVal str_ As String)
    Dim outputMessage As String
    outputMessage = "[" & function_ & ":" & CStr(line_) & "] "
    outputMessage = outputMessage & str_
    fprintf STDOUT_FILE, "%s", outputMessage
End Sub
