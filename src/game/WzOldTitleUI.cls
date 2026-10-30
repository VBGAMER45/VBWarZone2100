VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzOldTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/old.cpp: class WzOldTitleUI, the catch-all title screen for
' the menus of frontend.cpp (tMode), whole. Made by WzOldTitleUI_new
' (titleui_c).
Option Explicit

Implements WzTitleUI

Private mode As tMode

Friend Sub init(ByVal mode_ As tMode)
    mode = mode_
End Sub

Private Sub WzTitleUI_start()
    Select Case mode
    Case SINGLE_
        ActivityManager_navigateToMenu "Single Player"
        startSinglePlayerMenu
    Case TUTORIAL
        ActivityManager_navigateToMenu "Tutorial"
        startTutorialMenu
    Case TITLE
        ActivityManager_navigateToMenu "Main"
        startTitleMenu
    Case MULTI
        ActivityManager_navigateToMenu "Multiplayer"
        startMultiPlayerMenu       ' goto multiplayer menu
    Case STARTGAME, QUIT, LOADSAVEGAME
        bLimiterLoaded = False
    Case SHOWINTRO
        ActivityManager_navigateToMenu "Show Intro"
    Case Else
        debug_ 78, LOG_FATAL, "start", "Unknown title mode requested"
        exit_ 3             ' abort()
    End Select

    Exit Sub
End Sub

Private Function WzTitleUI_run() As Long
    If mode <> STARTGAME Then
        screen_disableMapPreview
    End If

    Select Case mode ' run relevant title screen code.
    ' MULTIPLAYER screens
    Case MULTI
        runMultiPlayerMenu

    Case TITLE
        runTitleMenu

    Case SINGLE_
        runSinglePlayerMenu

    Case TUTORIAL
        runTutorialMenu

    Case QUIT
        WzTitleUI_run = TITLECODE_QUITGAME
        Exit Function

    ' The "don't flip" behavior is preserved in wrappers.cpp by checking for these title codes (which are clearly special)
    Case STARTGAME
        WzTitleUI_run = TITLECODE_STARTGAME
        Exit Function
    Case LOADSAVEGAME
        WzTitleUI_run = TITLECODE_SAVEGAMELOAD
        Exit Function

    Case SHOWINTRO
        pie_SetFogStatus False
        pie_ScreenFrameRenderEnd
        pie_ScreenFrameRenderBegin
        changeTitleMode TITLE
        WzTitleUI_run = TITLECODE_SHOWINTRO
        Exit Function

    Case Else
        debug_ 129, LOG_FATAL, "run", "unknown title screen mode"
        exit_ 3             ' abort()
    End Select
    WzTitleUI_run = TITLECODE_CONTINUE
End Function

Private Sub WzTitleUI_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ' no-op
End Sub
