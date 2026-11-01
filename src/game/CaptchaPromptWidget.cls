VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CaptchaPromptWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class CaptchaPromptWidget
' lib/widget/verifyprompt.cpp: the fields of class CaptchaPromptWidget (verifyprompt_c)
Option Explicit

Public promptData_image As B64ImageData ' CaptchaVerifyPromptData promptData (its image)
Public responseCallback As WzFunc       ' CaptchaVerifyResponseFunc responseCallback;

Public promptText As WIDGET             ' std::shared_ptr<W_LABEL>
Public pImageTexture As gfx_api_texture ' gfx_api::texture *pImageTexture = nullptr;
Public imageScalingRatio As Double      ' = 1.0

Public inputBox As WIDGET               ' std::shared_ptr<W_EDITBOX>
Public okayButton As WIDGET             ' std::shared_ptr<PromptActionButtonWidget>

Public submittedEntry As Boolean        ' = false

Public weakSerial As Long               ' (VB) its std::weak_ptr<CaptchaPromptWidget>s: the widgweak_vb serial

Private Sub Class_Initialize()
    imageScalingRatio = 1#
End Sub
