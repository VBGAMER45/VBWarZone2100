VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMusicManagerTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/musicmanagertitleui.cpp, musicmanagertitleui.h: class WzMusicManagerTitleUI, the Music Manager title
' screen, whole. Made by musicmanagertitleui_c's WzMusicManagerTitleUI_new, which has its lambdas and
' addOptionsBaseForm. Its std::weak_ptr is its address, registered while it lives
' (musicmanagertitleui_c's musicManagerTitleUI_lock).
Option Explicit

Implements WzTitleUI

Private parent As WzTitleUI             ' std::shared_ptr<WzTitleUI>
Private musicManagerForm As WIDGET      ' std::shared_ptr<WIDGET>
Private registeredPtr As Long

' WzMusicManagerTitleUI::WzMusicManagerTitleUI(std::shared_ptr<WzTitleUI> parent)
Friend Sub construct(ByVal parent_ As WzTitleUI)
    Set parent = parent_
    registeredPtr = ObjPtr(Me)
    musicManagerTitleUI_register registeredPtr
End Sub

' WzMusicManagerTitleUI::~WzMusicManagerTitleUI()
Private Sub Class_Terminate()
    musicManagerTitleUI_unregister registeredPtr
    ' currently a no-op
End Sub

Private Sub WzTitleUI_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
End Sub

Private Sub WzTitleUI_start()
    Dim psBotForm As WIDGET, weakSelf As Long, psQuitButton As WIDGET
    addBackdrop
    musicManagerTitleUI_addOptionsBaseForm

    Set psBotForm = widgGetFromID(psWScreen, FRONTEND_BOTFORM)
    If Not (Not psBotForm Is Nothing) Then ASSERT_FAILURE "psBotForm != nullptr", "src/titleui/musicmanagertitleui.cpp:68", "start", "Unable to find FRONTEND_BOTFORM?"

    addSideText FRONTEND_SIDETEXT, FRONTEND_SIDEX, FRONTEND_TOPFORM_WIDEY, "MUSIC MANAGER"

    weakSelf = ObjPtr(Me)

    Set psQuitButton = addMultiBut_screen(psWScreen, FRONTEND_BOTFORM, FRONTEND_QUIT, 10, 10, 30, 29, "Return To Previous Screen", IMAGE_RETURN, IMAGE_RETURN_HI, IMAGE_RETURN_HI)
    W_BUTTON_addOnClickHandler psQuitButton, WzFunc_new(FN_musicManagerTitleUI_quitClick, weakSelf)

    Set musicManagerForm = makeMusicManagerForm(False)
    psBotForm.attach musicManagerForm
    musicManagerForm.setCalcLayout WzFunc_new(FN_musicManagerTitleUI_formCalcLayout)
End Sub

Public Function getParentTitleUI() As WzTitleUI
    Set getParentTitleUI = parent
End Function

Private Function WzTitleUI_run() As Long
    Dim triggers() As WidgetTrigger
    widgRunScreen psWScreen, triggers

    widgDisplayScreen psWScreen ' show the widgets currently running

    If CancelPressed() Then
        changeTitleUI parent
    End If

    WzTitleUI_run = TITLECODE_CONTINUE
End Function
