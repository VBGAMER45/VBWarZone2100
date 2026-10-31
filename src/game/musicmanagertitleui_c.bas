Attribute VB_Name = "musicmanagertitleui_c"
' src/titleui/musicmanagertitleui.cpp: the factory, addOptionsBaseForm and the lambdas of class
' WzMusicManagerTitleUI (WzMusicManagerTitleUI.cls), whole. (VB) The lambdas' std::weak_ptr<WzMusicManagerTitleUI>
' is the screen's address, registered while it lives (as campaign_title_c's for the campaign selector).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

Public Const FN_musicManagerTitleUI_baseFormCalcLayout As Long = 15310 ' addOptionsBaseForm: LAMBDA_CALCLAYOUT_SIMPLE({...})
Public Const FN_musicManagerTitleUI_quitClick As Long = 15311         ' start: psQuitButton->addOnClickHandler([weakSelf](W_BUTTON& button) {...})
Public Const FN_musicManagerTitleUI_quitTask As Long = 15312          ' start: widgScheduleTask([weakSelf]() {...})
Public Const FN_musicManagerTitleUI_formCalcLayout As Long = 15313    ' start: musicManagerForm->setCalcLayout([](WIDGET *psWidget){...})

Private liveTitleUIs As Collection      ' (VB) the addresses of the WzMusicManagerTitleUIs alive

' std::make_shared<WzMusicManagerTitleUI>(parent)
Public Function WzMusicManagerTitleUI_new(ByVal parent As WzTitleUI) As WzTitleUI
    Dim ui As WzMusicManagerTitleUI
    Set ui = New WzMusicManagerTitleUI
    ui.construct parent
    Set WzMusicManagerTitleUI_new = ui
End Function

Private Function collectionHasKey(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Boolean
    On Error GoTo missing
    v = IsObject(c.Item(key))   ' (VB) not v = c.Item(key): an object item would be asked for a default member
    collectionHasKey = True
    Exit Function
missing:
    collectionHasKey = False
End Function

Public Sub musicManagerTitleUI_register(ByVal p As Long)
    If liveTitleUIs Is Nothing Then Set liveTitleUIs = New Collection
    liveTitleUIs.Add p, "p" & p
End Sub

Public Sub musicManagerTitleUI_unregister(ByVal p As Long)
    If liveTitleUIs Is Nothing Then Exit Sub
    If collectionHasKey(liveTitleUIs, "p" & p) Then liveTitleUIs.Remove "p" & p
End Sub

' std::weak_ptr<WzMusicManagerTitleUI>::lock(): the title UI, or Nothing once it is gone
Private Function musicManagerTitleUI_lock(ByVal p As Long) As WzMusicManagerTitleUI
    Dim o As WzMusicManagerTitleUI
    If p = 0 Or liveTitleUIs Is Nothing Then Exit Function
    If Not collectionHasKey(liveTitleUIs, "p" & p) Then Exit Function
    CopyMemory o, p, 4                  ' a reference without AddRef...
    Set musicManagerTitleUI_lock = o    ' ...this one counts
    CopyMemory o, 0&, 4                 ' and the first goes without Release
End Function

' static void addOptionsBaseForm()
Public Sub musicManagerTitleUI_addOptionsBaseForm()
    Dim parent As WIDGET, botForm As WIDGET
    Set parent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)

    Set botForm = IntFormAnimated_new()
    parent.attach botForm
    botForm.id = FRONTEND_BOTFORM

    botForm.setCalcLayout WzFunc_new(FN_musicManagerTitleUI_baseFormCalcLayout)
End Sub

'--------------------------------------------------------------------------
' the lambdas

Public Function musicmanagertitleui_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim strongSelf As WzMusicManagerTitleUI, psWidget As WIDGET, psParent As WIDGET, w As Long, h As Long, y0 As Long
    Select Case f.id
    Case FN_musicManagerTitleUI_baseFormCalcLayout  ' LAMBDA_CALCLAYOUT_SIMPLE: (WIDGET *psWidget)
        Set psWidget = args(0)
        psWidget.setGeometry FRONTEND_BOTFORM_WIDEX, FRONTEND_TOPFORM_WIDEY, 590, 460

    Case FN_musicManagerTitleUI_quitClick       ' [weakSelf](W_BUTTON& button)
        widgScheduleTask WzFunc_new(FN_musicManagerTitleUI_quitTask, f.cap(0))

    Case FN_musicManagerTitleUI_quitTask        ' [weakSelf]()
        Set strongSelf = musicManagerTitleUI_lock(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/musicmanagertitleui.cpp:78", "start", "No parent": Exit Function
        changeTitleUI strongSelf.getParentTitleUI() ' go back

    Case FN_musicManagerTitleUI_formCalcLayout  ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/musicmanagertitleui.cpp:87", "start", "No parent?": Exit Function
        w = psParent.width()
        h = psParent.height()
        y0 = 49
        psWidget.setGeometry 0, y0, w, h - y0 - 1
    End Select
End Function
