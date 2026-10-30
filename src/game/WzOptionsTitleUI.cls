VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzOptionsTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/optionstitleui.cpp, optionstitleui.h: class WzOptionsTitleUI, the options title
' screen (the options browser of titleui/options/optionsforms.cpp), whole. Made by
' WzOptionsTitleUI_new (optionstitleui_c, which has its lambdas).
Option Explicit

Implements WzTitleUI

Private screen As W_SCREEN              ' std::shared_ptr<W_SCREEN>
Private parent As WzTitleUI             ' std::shared_ptr<WzTitleUI>
Private optionsBrowserForm As WIDGET    ' std::shared_ptr<OptionsBrowserForm>
Private bAlreadyStarted As Boolean      ' = false

' WzOptionsTitleUI::WzOptionsTitleUI(std::shared_ptr<WzTitleUI> parent)
Friend Sub construct(ByVal parent_ As WzTitleUI)
    Set parent = parent_
    Set screen = W_SCREEN_make()
End Sub

' WzOptionsTitleUI::~WzOptionsTitleUI()
Private Sub Class_Terminate()
    ' currently a no-op
End Sub

Private Sub WzTitleUI_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    screen.screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

Private Sub WzTitleUI_start()
    Dim weakSelf As Long, psQuitButton As WIDGET, newRootForm As WIDGET
    If bAlreadyStarted Then
        WzTitleUI_screenSizeDidChange screenWidth, screenHeight, screenWidth, screenHeight
        Exit Sub
    End If

    bAlreadyStarted = True

    addBackdrop screen

    weakSelf = ObjPtr(Me)

    Set psQuitButton = WzMultiButton_new()
    psQuitButton.setGeometry 0, 0, 30, 29
    psQuitButton.setTip "Return"                    ' P_("menu", "Return")
    WzMultiButton_(psQuitButton).imNormal = AtlasImage_(FrontImages, IMAGE_RETURN)
    WzMultiButton_(psQuitButton).imDown = AtlasImage_(FrontImages, IMAGE_RETURN_HI)
    WzMultiButton_(psQuitButton).doHighlight = IMAGE_RETURN_HI
    WzMultiButton_(psQuitButton).tc = MAX_PLAYERS
    WzMultiButton_(psQuitButton).alpha = 255
    W_BUTTON_addOnClickHandler psQuitButton, WzFunc_new(FN_WzOptionsTitleUI_quitClick, weakSelf)

    Set optionsBrowserForm = createOptionsBrowser(False, Margin_wrap(8, 10, 8, 10, psQuitButton))    ' Margin(8, 10).wrap(psQuitButton)

    Set newRootForm = makeTitleFormWrapper(WzFunc_new(FN_WzOptionsTitleUI_title), optionsBrowserForm)
    newRootForm.setCalcLayout WzFunc_new(FN_WzOptionsTitleUI_rootCalcLayout)
    screen.psForm.attach newRootForm, ChildZPos_Front
End Sub

Public Function getParentTitleUI() As WzTitleUI
    Set getParentTitleUI = parent
End Function

Private Function WzTitleUI_run() As Long
    Dim triggers() As WidgetTrigger
    widgRunScreen screen, triggers

    widgDisplayScreen screen    ' show the widgets currently running

    If CancelPressed() Then
        changeTitleUI parent
    End If

    WzTitleUI_run = TITLECODE_CONTINUE
End Function
