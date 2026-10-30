VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCampaignSelectorTitleUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/campaign.cpp, campaign.h: class WzCampaignSelectorTitleUI, the campaign selector title screen.
' Made by campaign_title_c's WzCampaignSelectorTitleUI_new, which has its widgets and lambdas. Its std::weak_ptr
' is its address, registered while it lives (campaign_title_c's campaignSelectorTitleUI_lock).
Option Explicit

Implements WzTitleUI

Private parent As WzTitleUI             ' std::shared_ptr<WzTitleUI>
Private campaignSelector As WIDGET      ' std::shared_ptr<WzCampaignSelectorForm>
Private displayedPanel As WIDGET        ' std::shared_ptr<WIDGET>
Private registeredPtr As Long

' WzCampaignSelectorTitleUI::WzCampaignSelectorTitleUI(std::shared_ptr<WzTitleUI> parent)
Friend Sub construct(ByVal parent_ As WzTitleUI)
    Set parent = parent_
    registeredPtr = ObjPtr(Me)
    campaignSelectorTitleUI_register registeredPtr
End Sub

' WzCampaignSelectorTitleUI::~WzCampaignSelectorTitleUI()
Private Sub Class_Terminate()
    campaignSelectorTitleUI_unregister registeredPtr
    campaign_title_c_clearCachedAvailableCampaignMods
End Sub

Private Sub WzTitleUI_screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
End Sub

Private Sub WzTitleUI_start()
    Dim psBotForm As WIDGET, weakSelf As Long, psQuitButton As WIDGET, campaignSelectorY0 As Long
    addBackdrop
    addCampaignBaseForm

    Set psBotForm = widgGetFromID(psWScreen, FRONTEND_BOTFORM)
    If Not (Not psBotForm Is Nothing) Then ASSERT_FAILURE "psBotForm != nullptr", "src/titleui/campaign.cpp:2761", "start", "Unable to find FRONTEND_BOTFORM?"

    addSideText FRONTEND_SIDETEXT, FRONTEND_SIDEX, FRONTEND_TOPFORM_WIDEY, "CAMPAIGNS"

    weakSelf = ObjPtr(Me)

    Set psQuitButton = addMultiBut_screen(psWScreen, FRONTEND_BOTFORM, FRONTEND_QUIT, 10, 10, 30, 29, "Return", IMAGE_RETURN, IMAGE_RETURN_HI, IMAGE_RETURN_HI) ' P_("menu", "Return")
    W_BUTTON_addOnClickHandler psQuitButton, WzFunc_new(FN_campaignSelector_quitClick, weakSelf)
    campaignSelectorY0 = psQuitButton.y() + psQuitButton.height() + 10

    Set campaignSelector = WzCampaignSelectorForm_make(weakSelf)
    psBotForm.attach campaignSelector
    campaignSelector.setCalcLayout WzFunc_new(FN_campaignSelector_calcLayout, campaignSelectorY0)

    displayStartOptionsForm Nothing

    ' Call loadCampaignModsAsync
    loadCampaignModsAsync WzFunc_new(FN_loadCampaignMods_completion, weakSelf)

    ' TODO: Display loading spinner?

    ' show this only when the video sequences are not installed
    If Not seq_hasVideos() Then
        notifyAboutMissingVideos
    End If
End Sub

' (modInfo: Nothing = nullopt)
Friend Sub displayStartOptionsForm(ByVal modInfo As WzCampaignModInfo)
    Dim psBotForm As WIDGET, weakSelf As Long
    Set psBotForm = widgGetFromID(psWScreen, FRONTEND_BOTFORM)
    If Not (Not psBotForm Is Nothing) Then ASSERT_FAILURE "psBotForm != nullptr", "src/titleui/campaign.cpp:2827", "displayStartOptionsForm", "Unable to find FRONTEND_BOTFORM?"

    If Not displayedPanel Is Nothing Then
        psBotForm.detach displayedPanel
    End If

    weakSelf = ObjPtr(Me)

    Set displayedPanel = CampaignStartOptionsForm_make(weakSelf, modInfo)
    psBotForm.attach displayedPanel
    displayedPanel.setCalcLayout WzFunc_new(FN_displayedPanel_calcLayout, weakSelf)
End Sub

Public Function getParentTitleUI() As WzTitleUI
    Set getParentTitleUI = parent
End Function

' (VB) the lambdas' access to the private fields (friend class WzCampaignSelectorForm, and start's lambdas)
Friend Function getCampaignSelector() As WIDGET
    Set getCampaignSelector = campaignSelector
End Function

Friend Function getDisplayedPanel() As WIDGET
    Set getDisplayedPanel = displayedPanel
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
