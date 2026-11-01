VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iV_BaseImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/pietypes.h: class iV_BaseImage, what a texture upload takes: iV_Image (uncompressed) and
' iV_CompressedImage implement it (Implements). PublicNotCreatable: the exe and the other DLLs pass their
' iV_Images where one is taken (VB asks the object for this interface). pixel_format() is pixelFormat and
' data_size() is dataSize (Implements refuses a member name with an underscore). A
' std::unique_ptr<iV_BaseImage> is an object (Nothing = nullptr); C++'s dynamic_cast<iV_Image*> is
' gfx_api_image_basis_priv_c's asIV_Image (a TypeOf test).
Option Explicit

Public Function width() As Long
End Function

Public Function height() As Long
End Function

Public Function pixelFormat() As pixel_format
End Function

' Get a pointer to the image data that can be read
Public Function data() As Long
End Function

Public Function dataSize() As Long
End Function

Public Function bufferRowLength() As Long
End Function

Public Function bufferImageHeight() As Long
End Function
