VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFrameworkGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' What lib/framework calls in the game (src): the exe implements this
' interface (FrameworkGame.cls) and hands it to frameSetGame at startup,
' Nothing at the end. Method names without underscores (VB's Implements
' refuses them).
' The resource module's function pointers (frameresource.h: RES_BUFFERLOAD,
' RES_FILELOAD, RES_FREE, RESLOAD_CALLBACK) are numbers the game registers
' (resAddBufferLoad, resAddFileLoad, resSetLoadCallback; 0 is nullptr) and
' these methods call by number. void *pData is a Variant: Empty is
' nullptr; the game stores an object, or a number ((void *)1).
Option Explicit

' RES_BUFFERLOAD: bool (*)(const char *pBuffer, UDWORD size, void **pData)
Public Function resBufferLoad(ByVal buffLoad As Long, pBuffer() As Byte, ByVal size As Long, pData As Variant) As Boolean
End Function

' RES_FILELOAD: bool (*)(const char *pFile, void **pData)
Public Function resFileLoad(ByVal fileLoad As Long, ByVal pFile As String, pData As Variant) As Boolean
End Function

' RES_FREE: void (*)(void *pData)
Public Sub resRelease(ByVal release As Long, pData As Variant)
End Sub

' RESLOAD_CALLBACK: void (*)()
Public Sub resLoadCallback(ByVal callback As Long)
End Sub
