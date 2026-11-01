VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "B64ImageData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/netlobby.cpp: class B64ImageData : public WzVerifyPrompt::IvImageProviderInterface (the interface's
' only implementation: verifyprompt's image providers are these). type_ is netlobby_c's ImageType.
Option Explicit

Public b64Data As String
Public type_ As Long                            ' ImageType type = ImageType::PNG

' optional<iV_Image> GetImage() const (has_value: the result, in image; image is a new iV_Image)
Public Function GetImage(image As iV_Image) As Boolean
    Dim decodedB64Data() As Byte, n As Long, ivImage As iV_Image, resultText As String
    n = base64Decode(b64Data, decodedB64Data)
    If n = 0 Then ReDim decodedB64Data(0 To 0)
    If UBound(decodedB64Data) <> n - 1 And n > 0 Then ReDim Preserve decodedB64Data(0 To n - 1)
    Set ivImage = New iV_Image

    Select Case type_
    Case ImageType_PNG
        If n = 0 Then
            resultText = "iV_loadImage_PNG: Insufficient length for PNG header"
        Else
            resultText = iV_loadImage_PNG_mem(decodedB64Data, ivImage)
        End If
        If Len(resultText) <> 0 Then            ' !result.noError()
            debug_ 284, LOG_ERROR, "GetImage", "Failed to load image from memory buffer: %s", resultText
            GetImage = False
            Exit Function
        End If
    Case ImageType_JPEG
        If n = 0 Then
            resultText = "Invalid width and/or height"
        Else
            resultText = iV_loadImage_JPEG(decodedB64Data, ivImage)
        End If
        If Len(resultText) <> 0 Then
            debug_ 294, LOG_ERROR, "GetImage", "Failed to load image from memory buffer: %s", resultText
            GetImage = False
            Exit Function
        End If
    End Select

    Set image = ivImage
    GetImage = True
End Function
