VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gfx_api_texture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' gfx_api.h: gfx_api::texture (a 2D texture; gl_texture and null_texture
' implement it, and gfx_api_abstract_texture). The image is an iV_Image (the
' port has no iV_CompressedImage: see CONVENTIONS, "The renderer"). Methods
' have no underscores (Implements refuses them): upload_sub is uploadSub,
' get_dimensions is getDimensions. The texture is freed when its last
' reference goes (C++'s delete).
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

' gfx_api::texture
Public Function upload(ByVal mip_level As Long, ByVal image As iV_Image) As Boolean
End Function

Public Function uploadSub(ByVal mip_level As Long, ByVal offset_x As Long, ByVal offset_y As Long, ByVal image As iV_Image) As Boolean
End Function

Public Function id() As Long
End Function

' texture2dDimensions get_dimensions() const
Public Sub getDimensions(width As Long, height As Long)
End Sub
