VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzGameBrowserTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/gamebrowser.cpp, gamebrowser.h: class WzGameBrowserTitleUI, the game browser title screen
' (Multiplayer > Join Game), whole. Made by gamebrowser_c's WzGameBrowserTitleUI_new, which has its lambdas. Its
' std::weak_ptr is its address (gamebrowser_c's gameBrowserTitleUIFromPtr), as WzOptionsTitleUI's.
Option Explicit

Implements WzTitleUI

Private screen As W_SCREEN              ' std::shared_ptr<W_SCREEN>
Private parent As WzTitleUI             ' std::shared_ptr<WzTitleUI>
Private gameBrowserForm As WIDGET       ' std::shared_ptr<WIDGET>
Private bAlreadyStarted As Boolean      ' = false

' WzGameBrowserTitleUI::WzGameBrowserTitleUI(std::shared_ptr<WzTitleUI> parent)
Friend Sub construct(ByVal parent_ As WzTitleUI)
    Set parent = parent_
    Set screen = W_SCREEN_make()
End Sub

' WzGameBrowserTitleUI::~WzGameBrowserTitleUI()
Private Sub Class_Terminate()
    ' currently a no-op
End Sub

Private Sub WzTitleUI_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    screen.screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

Private Sub WzTitleUI_start()
    Dim weakSelf As Long, newRootForm As WIDGET
    If bAlreadyStarted Then
        WzTitleUI_screenSizeDidChange screenWidth, screenHeight, screenWidth, screenHeight
        refreshLobbyBrowser gameBrowserForm
        Exit Sub
    End If

    bAlreadyStarted = True

    addBackdrop screen

    weakSelf = ObjPtr(Me)

    Set gameBrowserForm = makeLobbyBrowser(WzFunc_new(FN_WzGameBrowserTitleUI_back, weakSelf))

    Set newRootForm = makeTitleFormWrapper(WzFunc_new(FN_WzGameBrowserTitleUI_title), gameBrowserForm)
    newRootForm.setCalcLayout WzFunc_new(FN_WzGameBrowserTitleUI_rootCalcLayout)
    screen.psForm.attach newRootForm, ChildZPos_Front
End Sub

Public Function getParentTitleUI() As WzTitleUI
    Set getParentTitleUI = parent
End Function

Private Function WzTitleUI_run() As Long
    Dim triggers() As WidgetTrigger
    screen_disableMapPreview

    widgRunScreen screen, triggers

    widgDisplayScreen screen    ' show the widgets currently running

    If CancelPressed() Then
        changeTitleUI parent
    End If

    WzTitleUI_run = TITLECODE_CONTINUE
End Function
