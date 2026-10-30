VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameworkGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The game side of WzFramework.dll's IFrameworkGame: the resource module's
' load and release functions (by their data_c numbers) and its load
' callback.
Option Explicit

Implements IFrameworkGame

Private Function IFrameworkGame_resBufferLoad(ByVal buffLoad As Long, pBuffer() As Byte, ByVal size As Long, pData As Variant) As Boolean
    IFrameworkGame_resBufferLoad = data_c.dataBufferLoad(buffLoad, pBuffer, size, pData)
End Function

Private Function IFrameworkGame_resFileLoad(ByVal fileLoad As Long, ByVal pFile As String, pData As Variant) As Boolean
    IFrameworkGame_resFileLoad = data_c.dataFileLoad(fileLoad, pFile, pData)
End Function

Private Sub IFrameworkGame_resRelease(ByVal release As Long, pData As Variant)
    data_c.dataRelease release, pData
End Sub

' resSetLoadCallback's function numbers (wrappers_c's fn_*)
Private Sub IFrameworkGame_resLoadCallback(ByVal callback As Long)
    Select Case callback
    Case fn_loadingScreenCallback: wrappers_c.loadingScreenCallback
    End Select
End Sub
