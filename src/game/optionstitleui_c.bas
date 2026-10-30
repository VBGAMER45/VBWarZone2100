Attribute VB_Name = "optionstitleui_c"
' src/titleui/optionstitleui.cpp: the factory and the lambdas of class WzOptionsTitleUI
' (WzOptionsTitleUI.cls), whole. (VB) The lambdas' std::weak_ptr<WzOptionsTitleUI> is the
' screen's address.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Const FN_WzOptionsTitleUI_quitClick As Long = 4200       ' WzOptionsTitleUI::start, psQuitButton->addOnClickHandler([weakSelf](W_BUTTON& button) {...})
Public Const FN_WzOptionsTitleUI_quitTask As Long = 4201        ' WzOptionsTitleUI::start, widgScheduleTask([weakSelf]() {...})
Public Const FN_WzOptionsTitleUI_title As Long = 4202           ' WzOptionsTitleUI::start, makeTitleFormWrapper([](){ return _("OPTIONS"); }, ...)
Public Const FN_WzOptionsTitleUI_rootCalcLayout As Long = 4203  ' WzOptionsTitleUI::start, LAMBDA_CALCLAYOUT_SIMPLE({...})

' std::make_shared<WzOptionsTitleUI>(parent)
Public Function WzOptionsTitleUI_new(ByVal parent As WzTitleUI) As WzTitleUI
    Dim ui As WzOptionsTitleUI
    Set ui = New WzOptionsTitleUI
    ui.construct parent
    Set WzOptionsTitleUI_new = ui
End Function

' (VB) std::weak_ptr<WzOptionsTitleUI>::lock(): the title UI at the address (0: nullptr)
Private Function optionsTitleUIFromPtr(ByVal p As Long) As WzOptionsTitleUI
    Dim o As WzOptionsTitleUI
    If p = 0 Then Exit Function
    CopyMemory o, p, 4                  ' a reference without AddRef...
    Set optionsTitleUIFromPtr = o       ' ...this one counts
    CopyMemory o, 0&, 4                 ' and the first goes without Release
End Function

Public Function optionstitleui_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongSelf As WzOptionsTitleUI, psWidget As WIDGET
    Select Case f.id
    Case FN_WzOptionsTitleUI_quitClick          ' [weakSelf](W_BUTTON& button)
        widgScheduleTask WzFunc_new(FN_WzOptionsTitleUI_quitTask, f.cap(0))

    Case FN_WzOptionsTitleUI_quitTask           ' [weakSelf]()
        Set strongSelf = optionsTitleUIFromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/optionstitleui.cpp:77", "start", "No parent": Exit Function
        changeTitleUI strongSelf.getParentTitleUI() ' go back

    Case FN_WzOptionsTitleUI_title              ' [](){ return _("OPTIONS"); }
        optionstitleui_c_fn = "OPTIONS"

    Case FN_WzOptionsTitleUI_rootCalcLayout     ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry 0, 0, screenWidth, screenHeight
    End Select
End Function
