VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlidePromptWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class SlidePromptWidget
' lib/widget/verifyprompt.cpp: the fields of class SlidePromptWidget (verifyprompt_c)
Option Explicit

Public promptData_image As B64ImageData ' SlideVerifyPromptData promptData
Public promptData_tile As B64ImageData
Public promptData_tile_width As Long    ' uint32_t
Public promptData_tile_height As Long
Public promptData_tile_x As Long
Public promptData_tile_y As Long
Public dragDrop As Boolean              ' = false
Public responseCallback As WzFunc       ' SlideVerifyResponseFunc responseCallback;

Public promptText As WIDGET             ' std::shared_ptr<W_LABEL>
Public pImageTexture As gfx_api_texture ' gfx_api::texture *pImageTexture = nullptr;
Public imageScalingRatio As Double      ' = 1.0
Public imageWidth As Long               ' uint32_t
Public imageHeight As Long

Public pTileTexture As gfx_api_texture  ' gfx_api::texture *pTileTexture = nullptr;
Public currTileX As Long                ' uint32_t
Public currTileY As Long

Public slider As WIDGET                 ' std::shared_ptr<VerifySlideWidget>
Public has_handlingDragTileStartMouseOffset As Boolean  ' optional<Vector2i> handlingDragTileStartMouseOffset = nullopt;
Public handlingDragTileStartMouseOffsetX As Long
Public handlingDragTileStartMouseOffsetY As Long

Public submittedResponse As Boolean     ' = false

Public weakSerial As Long               ' (VB) its std::weak_ptr<SlidePromptWidget>s: the widgweak_vb serial

Private Sub Class_Initialize()
    imageScalingRatio = 1#
End Sub
