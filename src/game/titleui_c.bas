Attribute VB_Name = "titleui_c"
' src/titleui/titleui.cpp, titleui.h: the current title screen
' (wzTitleUICurrent) and changeTitleUI; the constructors of the title screens
' of titleui.h that are classes here: WzOldTitleUI (old.cpp: WzOldTitleUI_new)
' and WzMsgBoxTitleUI (msgbox.cpp: WzMsgBoxTitleUI_new). The title screens are
' class modules implementing the interface WzTitleUI (CONVENTIONS: a class
' hierarchy outside the simulation): VB's reference counting stands in for the
' shared_ptr, so a screen that holds its parent keeps it alive as in C++. The
' other screens have their constructors in their own modules
' (WzMultiplayerOptionsTitleUI_new in multiint_c, ...).
Option Explicit

' Pointer to the current UI. Dynamic allocation helps with the encapsulation.
Public wzTitleUICurrent As WzTitleUI

' WzOldTitleUI::WzOldTitleUI(tMode mode), std::make_shared<WzOldTitleUI>(mode)
Public Function WzOldTitleUI_new(ByVal mode As tMode) As WzTitleUI
    Dim ui As WzOldTitleUI
    Set ui = New WzOldTitleUI
    ui.init mode
    Set WzOldTitleUI_new = ui
End Function

' WzMsgBoxTitleUI::WzMsgBoxTitleUI(WzString title, WzString text, std::shared_ptr<WzTitleUI> next)
Public Function WzMsgBoxTitleUI_new(ByVal title_ As String, ByVal text As String, ByVal nxt As WzTitleUI) As WzTitleUI
    Dim ui As WzMsgBoxTitleUI
    Set ui = New WzMsgBoxTitleUI
    ui.init title_, text, nxt
    Set WzMsgBoxTitleUI_new = ui
End Function

' target Nothing: nullptr
Public Sub changeTitleUI(ByVal target As WzTitleUI)
    Set wzTitleUICurrent = target
    If Not psWScreen Is Nothing Then
        ' Deletes the backdrop, which in turn deletes the rest of the UI from whatever was on screen before this UI
        widgDelete psWScreen, FRONTEND_BACKDROP
    End If
    If Not target Is Nothing Then
        target.start
    End If
End Sub
