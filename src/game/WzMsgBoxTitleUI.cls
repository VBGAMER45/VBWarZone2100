VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMsgBoxTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/msgbox.cpp: class WzMsgBoxTitleUI, a message box that goes on to
' another title screen when acknowledged. Made by WzMsgBoxTitleUI_new
' (titleui_c).
'
' This prevents various problems that came up with using the lobby-based-console,
'  including (for example) the big tangling-up of everything...
' -- 20kdc
Option Explicit

Implements WzTitleUI

' titleui.h
Private Const WZ_MSGBOX_TUI_LEAVE As Long = 4597000

Private title_ As String
Private text As String
' Where to go after the user has acknowledged.
Private nxt As WzTitleUI

Friend Sub init(ByVal newTitle As String, ByVal newText As String, ByVal newNext As WzTitleUI)
    title_ = newTitle
    text = newText
    Set nxt = newNext
End Sub

Private Sub WzTitleUI_start()
    Dim psBotForm As WIDGET, msgboxTitle As WIDGET, paragraph As WIDGET
    Dim msgboxTitleContentsPadding As Long, msgboxTitleBottomY As Long, msgboxContentsY As Long
    addBackdrop
    addTopForm False
    addBottomForm
    addMultiBut_screen psWScreen, FRONTEND_BOTFORM, WZ_MSGBOX_TUI_LEAVE, FRONTEND_BOTFORMW - (MULTIOP_OKW * 2), FRONTEND_BOTFORMH - (MULTIOP_OKH * 2), MULTIOP_OKW, MULTIOP_OKH, "Continue", IMAGE_OK, IMAGE_OK, 1 ' (VB) hi: true (1)
    Set psBotForm = widgGetFromID(psWScreen, FRONTEND_BOTFORM)
    If Not psBotForm Is Nothing Then
        Set msgboxTitle = W_LABEL_new()
        If Not msgboxTitle Is Nothing Then ' silence a GCC warning ...
            W_LABEL_setFont msgboxTitle, font_medium_bold, psPalette(WZCOL_TEXT_BRIGHT)
            msgboxTitle.setString title_
            msgboxTitle.setGeometry MULTIOP_OKW, MULTIOP_OKH, W_LABEL_getMaxLineWidth(msgboxTitle), iV_GetTextLineSize(font_regular_bold)
            psBotForm.attach msgboxTitle
        End If
        Set paragraph = Paragraph_new()
        msgboxTitleContentsPadding = 20
        msgboxTitleBottomY = msgboxTitle.y() + msgboxTitle.height()
        msgboxContentsY = msgboxTitleBottomY + msgboxTitleContentsPadding
        paragraph.setGeometry MULTIOP_OKW, msgboxContentsY, FRONTEND_BOTFORMW - (MULTIOP_OKW * 2), FRONTEND_BOTFORMH - msgboxContentsY - (MULTIOP_OKH * 2)
        Paragraph_setFontColour paragraph, psPalette(WZCOL_TEXT_BRIGHT)
        Paragraph_setLineSpacing paragraph, 5
        Paragraph_setFont paragraph, font_regular
        Paragraph_addText paragraph, text
        psBotForm.attach paragraph
    Else
        debug_ 82, LOG_FATAL, "start", "No bottom form?"
    End If
End Sub

Private Function WzTitleUI_run() As Long
    Dim triggers() As WidgetTrigger, n As Long, id As Long
    screen_disableMapPreview
    n = widgRunScreen(psWScreen, triggers)
    If n = 0 Then id = 0 Else id = triggers(0).widget.id
    If id = WZ_MSGBOX_TUI_LEAVE Then
        changeTitleUI nxt
    End If
    widgDisplayScreen psWScreen
    WzTitleUI_run = TITLECODE_CONTINUE
End Function

Private Sub WzTitleUI_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ' WzTitleUI::screenSizeDidChange: no-op
End Sub
