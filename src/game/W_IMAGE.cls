VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_IMAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_IMAGE
' src/screens/guidescreen.cpp: W_IMAGE's fields (guidescreen_c); the initialisers are W_IMAGE_construct's
' and Class_Initialize's.
Option Explicit

Public pOwnedBackgroundTexture As gfx_api_texture ' gfx_api::texture* = nullptr (released with this object: ~W_IMAGE's delete)
Public textureWidth As Long                  ' int32_t = 0
Public textureHeight As Long                 ' int32_t = 0
