VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gfx_api_texture_array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' gfx_api.h: gfx_api::texture_array (gl_texture_array and null_texture_array
' implement it, and gfx_api_abstract_texture). upload_layer is uploadLayer.
Option Explicit

' (gfx_api::abstract_texture)
Public Sub bind()
End Sub

Public Function isArray() As Boolean
End Function

Public Sub flush()
End Sub

Public Function backendInternalValue() As Long
End Function

' gfx_api::texture_array
Public Function uploadLayer(ByVal layer As Long, ByVal mip_level As Long, ByVal image As iV_Image) As Boolean
End Function

Public Function id() As Long
End Function
