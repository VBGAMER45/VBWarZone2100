VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMultiLimitTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/multilimit.cpp, titleui.h: class WzMultiLimitTitleUI, the structure limits screen of the game's options.
' Made by WzMultiLimitTitleUI_new (multilimit_c, which has its ids, displayStructureBar and its lambdas).
Option Explicit

Implements WzTitleUI

' The parent WzMultiplayerOptionsTitleUI to return to.
Private parent As WzMultiplayerOptionsTitleUI

' WzMultiLimitTitleUI(std::shared_ptr<WzMultiplayerOptionsTitleUI> parent) : parent(parent)
Friend Sub init(ByVal newParent As WzMultiplayerOptionsTitleUI)
    Set parent = newParent
End Sub

Private Sub WzTitleUI_start()
    Dim limitsForm As WIDGET, limitsList As WIDGET, button As WIDGET, slider As WIDGET, limitsButtonId As Long, i As Long
    addBackdrop 'background

    ' load stats...
    If Not bLimiterLoaded Then
        initLoadingScreen True

        If Not resLoad("wrf/limiter_data.wrf", 503) Then
            debug_ 116, LOG_INFO, "start", "Unable to load limiter_data during WzMultiLimitTitleUI start; returning..."
            changeTitleUI parent
            closeLoadingScreen
            Exit Sub
        End If

        bLimiterLoaded = True

        closeLoadingScreen
    End If

    If challengeActive Then
        resetLimits
    End If

    ' TRANSLATORS: Sidetext of structure limits screen
    addSideText FRONTEND_SIDETEXT1, LIMITSX - 2, LIMITSY, "LIMITS"

    Set limitsForm = IntFormAnimated_new(False)
    widgGetFromID(psWScreen, FRONTEND_BACKDROP).attach limitsForm
    limitsForm.id = IDLIMITS
    limitsForm.setCalcLayout WzFunc_new(FN_limitsForm_calcLayout)

    ' force limits off button (hi: C++ passes true, 1)
    addMultiBut_screen psWScreen, IDLIMITS, IDLIMITS_FORCEOFF, LIMITS_FORCEOFFX, LIMITS_OKY, _
        iV_GetImageWidth(FrontImages, IMAGE_DARK_UNLOCKED), iV_GetImageHeight(FrontImages, IMAGE_DARK_UNLOCKED), _
        "Map Can Exceed Limits", IMAGE_DARK_UNLOCKED, IMAGE_DARK_UNLOCKED, 1
    ' force limits button
    addMultiBut_screen psWScreen, IDLIMITS, IDLIMITS_FORCE, LIMITS_FORCEX, LIMITS_OKY, _
        iV_GetImageWidth(FrontImages, IMAGE_DARK_LOCKED), iV_GetImageHeight(FrontImages, IMAGE_DARK_LOCKED), _
        "Force Limits at Start", IMAGE_DARK_LOCKED, IMAGE_DARK_LOCKED, 1
    If challengeActive Then
        widgSetButtonState psWScreen, IDLIMITS_FORCE, WBUT_DISABLE
        widgSetButtonState psWScreen, IDLIMITS_FORCEOFF, WBUT_DISABLE
    ElseIf (ingame.flags And MPFLAGS_FORCELIMITS) <> 0 Then
        widgSetButtonState psWScreen, IDLIMITS_FORCE, WBUT_DISABLE
        widgSetButtonState psWScreen, IDLIMITS_FORCEOFF, 0
    Else
        widgSetButtonState psWScreen, IDLIMITS_FORCE, 0
        widgSetButtonState psWScreen, IDLIMITS_FORCEOFF, WBUT_DISABLE
    End If

    ' return button.
    addMultiBut_screen psWScreen, IDLIMITS, IDLIMITS_RETURN, LIMITS_OKX - 40, LIMITS_OKY, _
        iV_GetImageWidth(FrontImages, IMAGE_NO), iV_GetImageHeight(FrontImages, IMAGE_NO), _
        "Apply Defaults and Return To Previous Screen", IMAGE_NO, IMAGE_NO, 1
    If challengeActive Then
        widgSetButtonState psWScreen, IDLIMITS_RETURN, WBUT_DISABLE
    End If

    ' ok button
    addMultiBut_screen psWScreen, IDLIMITS, IDLIMITS_OK, LIMITS_OKX, LIMITS_OKY, _
        iV_GetImageWidth(FrontImages, IMAGE_OK), iV_GetImageHeight(FrontImages, IMAGE_OK), _
        "Accept Settings", IMAGE_OK, IMAGE_OK, 1

    ' add tab form..
    Set limitsList = IntListTabWidget_make()
    limitsForm.attach limitsList
    limitsList.setCalcLayout WzFunc_new(FN_limitsList_calcLayout)

    'Put the buttons on it
    limitsButtonId = IDLIMITS_ENTRIES_START

    For i = 0 To numStructureStats - 1
        If asStructureStats(i).base.limit <> LOTS_OF Then
            Set button = W_FORM_new()
            limitsList.attach button
            button.id = limitsButtonId
            Set button.displayFunction = WzFunc_new(FN_displayStructureBar)
            button.UserData = i
            button.pUserData = Array(WzText_new("", font_regular), WzText_new("", font_regular))   ' new DisplayStructureBarCache()
            button.setOnDelete WzFunc_new(FN_limitsButton_onDelete)
            ListTabWidget_addWidgetToLayout limitsList, button
            limitsButtonId = limitsButtonId + 1

            Set slider = addFESlider(limitsButtonId, limitsButtonId - 1, 290, 11, asStructureStats(i).maxLimit, asStructureStats(i).upgrade(0).limit)
            If challengeActive Then
                W_SLIDER_disable slider
            End If
            limitsButtonId = limitsButtonId + 1
        End If
    Next
End Sub

Private Function WzTitleUI_run() As Long
    Dim triggers() As WidgetTrigger, n As Long, id As Long, statid As Long
    parent.frontendMultiMessages False                  ' network stuff.

    n = widgRunScreen(psWScreen, triggers)
    If n = 0 Then id = 0 Else id = triggers(0).widget.id ' Just use first click here, since the next click could be on another menu.

    ' sliders
    If (id > IDLIMITS_ENTRIES_START) And (id < IDLIMITS_ENTRIES_END) Then
        If Not challengeActive Then
            statid = widgGetFromID(psWScreen, id - 1).UserData
            If statid <> 0 Then
                asStructureStats(statid).upgrade(0).limit = W_SLIDER_(widgGetFromID(psWScreen, id)).pos And &HFF&     ' (UBYTE)
            End If
        End If
    Else
        ' icons that are always about.
        Select Case id
        Case IDLIMITS_FORCE
            widgSetButtonState psWScreen, IDLIMITS_FORCE, WBUT_DISABLE
            widgSetButtonState psWScreen, IDLIMITS_FORCEOFF, 0
        Case IDLIMITS_FORCEOFF
            widgSetButtonState psWScreen, IDLIMITS_FORCE, 0
            widgSetButtonState psWScreen, IDLIMITS_FORCEOFF, WBUT_DISABLE
        Case IDLIMITS_RETURN
            If challengeActive Then
                changeTitleUI parent
            Else
                ' reset the sliders..
                resetLimits
                ' free limiter structure
                freeLimitSet
                If widgGetButtonState(psWScreen, IDLIMITS_FORCE) <> 0 Then
                    ingame.flags = ingame.flags And Not MPFLAGS_FORCELIMITS
                End If
                'inform others
                If NetPlay_isHost Then
                    sendOptions
                End If

                changeTitleUI parent

                ' make some noize.
                If Not ingame.localOptionsReceived Then
                    displayRoomSystemMessage "Limits reset to default values"
                Else
                    sendRoomSystemMessage "Limits Reset To Default Values"
                End If

                resetReadyStatus False
            End If
        Case IDLIMITS_OK
            If challengeActive Then
                changeTitleUI parent
            Else
                If Not challengeActive And widgGetButtonState(psWScreen, IDLIMITS_FORCE) <> 0 Then
                    ingame.flags = ingame.flags Or MPFLAGS_FORCELIMITS
                Else
                    ingame.flags = ingame.flags And Not MPFLAGS_FORCELIMITS
                End If
                resetReadyStatus False
                createLimitSet
                changeTitleUI parent
            End If
        Case Else
        End Select
    End If

    widgDisplayScreen psWScreen                         ' show the widgets currently running
    WzTitleUI_run = TITLECODE_CONTINUE
End Function

Private Sub WzTitleUI_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ' WzTitleUI::screenSizeDidChange: no-op
End Sub
