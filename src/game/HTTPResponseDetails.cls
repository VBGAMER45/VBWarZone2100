VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTPResponseDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/urlrequest.h: class HTTPResponseDetails, with urlrequest_curl.cpp's CURLHTTPResponseDetails (the CURLcode
' that WzNet's urltransfer_c gives, with curl_easy_strerror's text) and urlrequest_private.h's
' HTTPResponseHeadersContainer (ResponseHeaderContainer: a Collection keyed by the header's name, which VB compares
' without case as CaseInsensitiveHash / CaseInsensitiveEqualFunc do).
Option Explicit

Private httpStatusCode_ As Long
Private responseHeaders As Collection           ' of Array(name, value), keyed by name
Private transportResult As Long                 ' CURLcode _curlResult
Private transportResultText As String
Private has_primaryIp As Boolean
Private primaryIp As String

Private Sub Class_Initialize()
    Set responseHeaders = New Collection
End Sub

' (VB) the constructor: CURLHTTPResponseDetails(curlResult, httpStatusCode, primaryIp, responseHeaders)
Friend Sub init(ByVal result As Long, ByVal resultText As String, ByVal httpStatusCode As Long, ByVal hasPrimaryIp As Boolean, ByVal ip As String)
    transportResult = result
    transportResultText = resultText
    httpStatusCode_ = httpStatusCode
    has_primaryIp = hasPrimaryIp
    primaryIp = ip
End Sub

' header_callback: pRequest->responseHeaders->responseHeaders[header_key] = header_value
Friend Sub setHeader(ByVal key As String, ByVal value As String)
    On Error Resume Next
    responseHeaders.Remove key
    On Error GoTo 0
    responseHeaders.Add Array(key, value), key
End Sub

Public Function curlResult() As Long
    curlResult = transportResult
End Function

' optional<std::string> getPrimaryIP() const (has_value: the result)
Public Function getPrimaryIP(out As String) As Boolean
    out = primaryIp
    getPrimaryIP = has_primaryIp
End Function

' std::string getInternalResultDescription() const (curl_easy_strerror's text)
Public Function getInternalResultDescription() As String
    getInternalResultDescription = transportResultText
End Function

Public Function httpStatusCode() As Long
    httpStatusCode = httpStatusCode_
End Function

' Permits retrieving HTTP response headers in a case-insensitive manner
Public Function hasHeader(ByVal name_ As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = responseHeaders.Item(name_)
    hasHeader = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function getHeader(ByVal name_ As String, output_value As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = responseHeaders.Item(name_)
    If Err.Number <> 0 Then
        On Error GoTo 0
        getHeader = False
        Exit Function
    End If
    On Error GoTo 0
    output_value = v(1)
    getHeader = True
End Function
