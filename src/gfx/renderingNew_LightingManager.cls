VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "renderingNew_LightingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/ivis_opengl/pielighting.h/.cpp: renderingNew::LightingManager, the
' lighting manager that generates a proper PointLightBuckets for per pixel
' point lights. Its ComputeFrameData is pielighting_c's
' renderingNew_LightingManager_ComputeFrameData (on the cached containers).
' Made by renderingNew_LightingManager_new.
Option Explicit

Implements ILightingManager

' cached containers to avoid frequent reallocations: std::vector<CulledLightInfo> culledLights
' (CulledLightInfo's light and clipSpaceBoundingBox as two arrays)
Private culledLights_light(0 To gfx_api_max_lights - 1) As CalculatedPointLight
Private culledLights_box(0 To gfx_api_max_lights - 1) As BoundingBox
Private culledLights_size As Long

Private Sub ILightingManager_ComputeFrameData(ByVal worldViewProjectionMatrix As Long)
    Dim m As mat4
    m = mat4_fromPtr(worldViewProjectionMatrix)
    renderingNew_LightingManager_ComputeFrameData culledLights_light, culledLights_box, culledLights_size, m
End Sub
