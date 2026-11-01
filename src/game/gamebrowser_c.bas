Attribute VB_Name = "gamebrowser_c"
' src/titleui/gamebrowser.cpp: the factory and the lambdas of class WzGameBrowserTitleUI (WzGameBrowserTitleUI.cls),
' whole. (VB) The lambdas' std::weak_ptr<WzGameBrowserTitleUI> is the title UI's address, as optionstitleui_c's.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Const FN_WzGameBrowserTitleUI_back As Long = 16200          ' WzGameBrowserTitleUI::start, makeLobbyBrowser([weakSelf](){...})
Public Const FN_WzGameBrowserTitleUI_title As Long = 16201         ' WzGameBrowserTitleUI::start, makeTitleFormWrapper([](){ return _("GAMES"); }, ...)
Public Const FN_WzGameBrowserTitleUI_rootCalcLayout As Long = 16202 ' WzGameBrowserTitleUI::start, LAMBDA_CALCLAYOUT_SIMPLE({...})

' std::make_shared<WzGameBrowserTitleUI>(parent)
Public Function WzGameBrowserTitleUI_new(ByVal parent As WzTitleUI) As WzTitleUI
    Dim ui As WzGameBrowserTitleUI
    Set ui = New WzGameBrowserTitleUI
    ui.construct parent
    Set WzGameBrowserTitleUI_new = ui
End Function

' (VB) std::weak_ptr<WzGameBrowserTitleUI>::lock(): the title UI at the address (0: nullptr)
Private Function gameBrowserTitleUIFromPtr(ByVal p As Long) As WzGameBrowserTitleUI
    Dim o As WzGameBrowserTitleUI
    If p = 0 Then Exit Function
    CopyMemory o, p, 4                  ' a reference without AddRef...
    Set gameBrowserTitleUIFromPtr = o   ' ...this one counts
    CopyMemory o, 0&, 4                 ' and the first goes without Release
End Function

Public Function gamebrowser_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongSelf As WzGameBrowserTitleUI, psWidget As WIDGET
    Select Case f.id
    Case FN_WzGameBrowserTitleUI_back           ' [weakSelf]()
        Set strongSelf = gameBrowserTitleUIFromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/gamebrowser.cpp:65", "start", "No parent": Exit Function
        changeTitleUI strongSelf.getParentTitleUI() ' go back

    Case FN_WzGameBrowserTitleUI_title          ' [](){ return _("GAMES"); }
        gamebrowser_c_fn = "GAMES"

    Case FN_WzGameBrowserTitleUI_rootCalcLayout ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight
    End Select
End Function
