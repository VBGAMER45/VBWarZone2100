VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rendering1999_LightingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/lighting.h: rendering1999::LightingManager, the classic lights (they
' colour the light map's tiles). Implements WzGfx's ILightingManager; the work
' is lighting_c's rendering1999_LightingManager_ComputeFrameData. Created with
' New where C++ makes a std::make_unique<rendering1999::LightingManager>().
Option Explicit

Implements ILightingManager

Private Sub ILightingManager_ComputeFrameData(ByVal worldViewProjectionMatrix As Long)
    lighting_c.rendering1999_LightingManager_ComputeFrameData worldViewProjectionMatrix
End Sub
