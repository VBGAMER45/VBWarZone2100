VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ILightingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/pielighting.h: struct ILightingManager, its one virtual
' method. WzGfx implements it (renderingNew_LightingManager) and so does the
' exe (lighting.cpp's rendering1999::LightingManager); setLightingManager
' takes either. ComputeFrameData(const LightingData& data, LightMap& lightmap,
' const glm::mat4& worldViewProjectionMatrix): data and lightmap are always
' getCurrentLightingData() and getCurrentLightmapData(), which an
' implementation reaches through the LightingData_lights_* and LightMap_*
' procedures; the matrix is an address. The base class's SetFrameStart and
' getPointLightBuckets are pielighting_c's ILightingManager_<name>.
Option Explicit

Public Sub ComputeFrameData(ByVal worldViewProjectionMatrix As Long)
End Sub
