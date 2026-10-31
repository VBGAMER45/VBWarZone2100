Attribute VB_Name = "campaign_title_c"
' src/titleui/campaign.cpp, campaign.h, whole: the campaign selector title UI (WzCampaignSelectorTitleUI.cls, made by
' WzCampaignSelectorTitleUI_new) and its widget classes (their fields: the data classes named as the classes;
' WzCampaignTweakOptionsSummaryWidget's is WzCampaignTweakOptSummaryData, VB's 31 characters). X::make is X_make.
' (VB) The std::weak_ptr captures are the widget's address (ObjPtr, WIDGET_fromPtr) as the other title screens
' have them; the title UI's is checked against the live ones (campaignSelectorTitleUI_lock). optional<T> of a
' class is Nothing for nullopt. A std::vector of objects is a Collection; the shared
' std::vector<WzCampaignTweakOptionSetting> is one Collection both forms hold (assignCollection replaces its
' contents, as *tweakOptionSettings = ... does).
' modinfo.h's enums are here (its classes are WzCampaignModInfo.cls and WzCampaignTweakOption.cls; modinfo.cpp is
' modinfo_c: loadCampaignModInfoFromFile, getModCompatibility). The mods' .wz files mount through WzFramework's
' PhysFS (physfs_win, physfs_zip_c). Threads are not ported: loadCampaignModsAsync runs loadCampaignModsAsyncImpl
' at once (its completion handler queues its work with wzAsyncExecOnMainThread, as from the thread). Translations
' (_()) are not ported: the English text is used.
'
'  Campaign Selector / Configuration / Startup Title UI.
' (VB) widget class WzCampaignModButton : W_BUTTON
' (VB) widget class WzCampaignListSection : WIDGET
' (VB) widget class WzCampaignSelectorForm : WIDGET
' (VB) widget class CampaignStartButton : W_BUTTON
' (VB) widget class WzCampaignOptionButton : W_BUTTON
' (VB) widget class WzCampaignTweakOptionsEditForm : WIDGET
' (VB) widget class WidgetInfoButtonWrapper : WIDGET
' (VB) widget class WzCampaignTweakOptionsSummaryWidget : WIDGET
' (VB) widget class FixedIdealWidthLabel : W_LABEL
' (VB) widget class WzCampaignTitleWidget : WIDGET
' (VB) widget class WzCampaignSubInfoBox : WIDGET
' (VB) widget class CampaignStartOptionsForm : WIDGET
' (VB) widget class WzCampaignHelpPanel : WIDGET
' (VB) widget class WZOptionTitleLabel : WIDGET
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' typedef WzAdvCheckbox WzCampaignTweakOptionToggle;

Private Const WZ2100_CORE_CLASSIC_BALANCE_MOD_FILENAME As String = "wz2100_camclassic.wz"
Private Const WZ2100_CORE_PUMPKIN_BALANCE_MOD_FILENAME As String = "wz2100_campumpkin.wz"

Private Const FRONTEND_CAMPAIGNUI_BUTHEIGHT As Long = 28

' modinfo.h
Public Enum WzModType
    WzModType_CampaignBalanceMod
    WzModType_AlternateCampaign
End Enum

Public Enum WzModCompatibilityResult
    WzModCompatibilityResult_NOT_COMPATIBLE
    WzModCompatibilityResult_OK
    WzModCompatibilityResult_MAYBE_COMPATIBLE
End Enum

Public Enum WzCampaignUniverse
    WzCampaignUniverse_WZ2100
    WzCampaignUniverse_WZ2100Extended
    WzCampaignUniverse_Unique
End Enum

' struct WzCampaignTweakOption's enum class Type
Public Enum WzCampaignTweakOption_Type
    WzCampaignTweakOption_Type_Bool
End Enum

' the lambdas (WzFunc: campaign_title_c_fn)
Public Const FN_campaignModButton_onClick As Long = 14400               ' makeCampaignModButton, [weakParent, modInfo](W_BUTTON& but)
Public Const FN_campaignListSection_titleCalcLayout As Long = 14401     ' WzCampaignListSection::initialize, sectionTitle
Public Const FN_campaignListSection_listCalcLayout As Long = 14402      ' WzCampaignListSection::initialize, sectionList
Public Const FN_selectorForm_originalCalcLayout As Long = 14403         ' WzCampaignSelectorForm::initialize, originalCampaignButton
Public Const FN_selectorForm_alternateCalcLayout As Long = 14404        ' WzCampaignSelectorForm::initialize, alternateCampaigns
Public Const FN_campaign_closeButtonClick As Long = 14405               ' the edit form's and help panel's [closeHandlerFunc](W_BUTTON&)
Public Const FN_campaign_closeButtonTask As Long = 14406                ' widgScheduleTask([closeHandlerFunc]() { closeHandlerFunc(); })
Public Const FN_tweakEditForm_resetClick As Long = 14407                ' [weakEditForm](W_BUTTON&)
Public Const FN_tweakEditForm_resetTask As Long = 14408                 ' widgScheduleTask([weakEditForm]() {...})
Public Const FN_tweakEditForm_toggleClick As Long = 14409               ' [weakEditForm, i](W_BUTTON& but)
Public Const FN_infoWrapper_infoClick As Long = 14410                   ' WidgetInfoButtonWrapper::initialize, [weakSelf](W_BUTTON&)
Public Const FN_tweakSummary_infoCalcLayout As Long = 14411             ' WzCampaignTweakOptionsSummaryWidget::initialize, infoButton
Public Const FN_tweakSummary_infoClick As Long = 14412                  ' [weakSelf](W_BUTTON&)
Public Const FN_tweakSummary_listCalcLayout As Long = 14413             ' tweakOptionsList
Public Const FN_tweakSummary_listClick As Long = 14414                  ' [weakSelf](ClickableScrollableList&)
Public Const FN_tweakSummary_listHighlight As Long = 14415              ' [weakSelf](ClickableScrollableList&, bool isHighlighted)
Public Const FN_titleWidget_labelCalcLayout As Long = 14416             ' WzCampaignTitleWidget::initialize, nameLabel and authorLabel
Public Const FN_titleWidget_descriptionCalcLayout As Long = 14417       ' [descriptionWidget](WIDGET *psWidget)
Public Const FN_subInfoBox_universeCalcLayout As Long = 14418           ' WzCampaignSubInfoBox::initialize, universeLabel
Public Const FN_subInfoBox_compatibilityCalcLayout As Long = 14419      ' compatibilityLabel
Public Const FN_difficultyDropdown_canChange As Long = 14420            ' [dropDownChoices](DropdownWidget &, size_t newIndex, newSelectedWidget) -> bool
Public Const FN_balanceDropdown_canChange As Long = 14421               ' [](DropdownWidget &, size_t newIndex, newSelectedWidget) -> bool
Public Const FN_balanceDropdown_onChange As Long = 14422                ' [weakSelf](DropdownWidget &widget)
Public Const FN_balanceDropdown_onChangeTask As Long = 14423            ' widgScheduleTask([weakSelf]() {...})
Public Const FN_balanceWrapper_infoClick As Long = 14424                ' infoWrappedCampaignBalanceDropdown->setInfoClickHandler([weakSelf]() {...})
Public Const FN_balanceInfo_closeOverlay As Long = 14425                ' WzCampaignHelpPanel::make(..., [weakSelf]() {...})
Public Const FN_balanceInfo_onClose As Long = 14426                     ' openOverlay(..., [weakSelf]() { /* no-op */ })
Public Const FN_overlay_clicked As Long = 14427                         ' newRootFrm->onClickedFunc = [psWeakOverlayScreen, psWeakSelf]() {...}
Public Const FN_overlay_popoverCalcLayout As Long = 14428               ' popoverForm->setCalcLayout([psWeakSelf](WIDGET *psWidget) {...})
Public Const FN_tweakEditFormOverlay_close As Long = 14429              ' WzCampaignTweakOptionsEditForm::make(..., [psWeakSelf](){...})
Public Const FN_tweakEditFormOverlay_onClose As Long = 14430            ' openOverlay(..., [weakSelf]() {...})
Public Const FN_startForm_titleAreaCalcLayout As Long = 14431           ' CampaignStartOptionsForm::initialize, titleArea
Public Const FN_startForm_subInfoBoxCalcLayout As Long = 14432          ' subInfoBox
Public Const FN_startForm_tweakSummaryInfoClick As Long = 14433         ' tweakOptionsSummary->setInfoClickHandler([weakSelf]() {...})
Public Const FN_startForm_scrollableListCalcLayout As Long = 14434      ' scrollableList
Public Const FN_startForm_startButtonCalcLayout As Long = 14435         ' psStartNewGameButton
Public Const FN_startForm_startButtonClick As Long = 14436              ' [weakSelf](W_BUTTON&)
Public Const FN_startForm_startButtonTask As Long = 14437               ' widgScheduleTask([weakSelf]() {...})
Public Const FN_campaignBaseForm_calcLayout As Long = 14438             ' addCampaignBaseForm, botForm
Public Const FN_campaignSelector_quitClick As Long = 14439              ' WzCampaignSelectorTitleUI::start, [weakSelf](W_BUTTON& button)
Public Const FN_campaignSelector_quitTask As Long = 14440               ' widgScheduleTask([weakSelf]() {...})
Public Const FN_campaignSelector_calcLayout As Long = 14441             ' [campaignSelectorY0](WIDGET *psWidget)
Public Const FN_loadCampaignMods_completion As Long = 14442             ' [weakSelf](std::vector<WzCampaignModInfo> campaignMods)
Public Const FN_loadCampaignMods_mainThread As Long = 14443             ' wzAsyncExecOnMainThread([campaignMods, weakSelf]() {...})
Public Const FN_displayedPanel_calcLayout As Long = 14444               ' displayStartOptionsForm, [weakSelf](WIDGET *psWidget)

Private cached_availableCampaignMods As Collection  ' static std::vector<WzCampaignModInfo> (Nothing: empty)

' (VB) the live WzCampaignSelectorTitleUIs ("p<ObjPtr>"): their std::weak_ptr lock
Private liveTitleUIs As Collection

'--------------------------------------------------------------------------
' (VB) helpers

' a PIELIGHT's alpha (its byte.a: the top byte)
Private Function pielightAlpha(ByVal c As Long) As Long
    pielightAlpha = SHRU(c, 24)
End Function

Private Function pielightWithAlpha(ByVal c As Long, ByVal a As Long) As Long
    pielightWithAlpha = (c And &HFFFFFF) Or SHL(a And &HFF&, 24)
End Function

' std::dynamic_pointer_cast<X>(w): w when it is an X, else Nothing
Private Function dynCast(ByVal w As WIDGET, ByVal kind As WIDGET_KIND) As WIDGET
    If WIDGET_isKindOf(w, kind) Then Set dynCast = w
End Function

Private Function minL(ByVal a As Long, ByVal b As Long) As Long
    If b < a Then minL = b Else minL = a
End Function

Private Function maxL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then maxL = b Else maxL = a
End Function

' *dst = src (a std::vector's contents, the Collection object kept)
Private Sub assignCollection(ByVal dst As Collection, ByVal src As Collection)
    Dim item As Variant
    Do While dst.Count > 0
        dst.Remove 1
    Loop
    For Each item In src
        dst.Add item
    Next
End Sub

Private Function collectionHasKey(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Boolean
    On Error GoTo missing
    v = IsObject(c.Item(key))   ' (VB) not v = c.Item(key): an object item would be asked for a default member
    collectionHasKey = True
    Exit Function
missing:
    collectionHasKey = False
End Function

Private Function availableCampaignModsCache() As Collection
    If cached_availableCampaignMods Is Nothing Then Set cached_availableCampaignMods = New Collection
    Set availableCampaignModsCache = cached_availableCampaignMods
End Function

' ~WzCampaignSelectorTitleUI's cached_availableCampaignMods.clear();
Public Sub campaign_title_c_clearCachedAvailableCampaignMods()
    Set cached_availableCampaignMods = New Collection
End Sub

Public Sub campaignSelectorTitleUI_register(ByVal p As Long)
    If liveTitleUIs Is Nothing Then Set liveTitleUIs = New Collection
    liveTitleUIs.Add p, "p" & p
End Sub

Public Sub campaignSelectorTitleUI_unregister(ByVal p As Long)
    If liveTitleUIs Is Nothing Then Exit Sub
    If collectionHasKey(liveTitleUIs, "p" & p) Then liveTitleUIs.Remove "p" & p
End Sub

' std::weak_ptr<WzCampaignSelectorTitleUI>::lock(): the title UI, or Nothing once it is gone
Private Function campaignSelectorTitleUI_lock(ByVal p As Long) As WzCampaignSelectorTitleUI
    Dim o As WzCampaignSelectorTitleUI
    If p = 0 Or liveTitleUIs Is Nothing Then Exit Function
    If Not collectionHasKey(liveTitleUIs, "p" & p) Then Exit Function
    CopyMemory o, p, 4                  ' a reference without AddRef...
    Set campaignSelectorTitleUI_lock = o ' ...this one counts
    CopyMemory o, 0&, 4                 ' and the first goes without Release
End Function

' std::make_shared<WzCampaignSelectorTitleUI>(parent)
Public Function WzCampaignSelectorTitleUI_new(ByVal parent As WzTitleUI) As WzTitleUI
    Dim ui As WzCampaignSelectorTitleUI
    Set ui = New WzCampaignSelectorTitleUI
    ui.construct parent
    Set WzCampaignSelectorTitleUI_new = ui
End Function

'--------------------------------------------------------------------------

Private Function validateCampaignMod(ByVal modInfo As WzCampaignModInfo, ByVal baseDir As String, ByVal realModFilePathAndName As String) As Boolean
    Dim RecognizedSPDXLicenseExpressions As Variant, recognized As Boolean, k As Long
    Dim invalidCampaignBalanceModContents As Boolean, files() As String, nFiles As Long, file As String
    Dim realFileName_platformIndependent As String, isDirectory As Boolean

    ' Verify the name is not empty
    If Len(modInfo.name_) = 0 Then
        debug_ 75, LOG_ERROR, "validateCampaignMod", "Campaign mod missing valid name field: %s", realModFilePathAndName
        validateCampaignMod = False
        Exit Function
    End If
    ' Verify the author is not empty
    If Len(modInfo.author) = 0 Then
        debug_ 81, LOG_ERROR, "validateCampaignMod", "Campaign mod missing valid author field: %s", realModFilePathAndName
        validateCampaignMod = False
        Exit Function
    End If
    ' Verify the license is not empty
    If Len(modInfo.license) = 0 Then
        debug_ 87, LOG_ERROR, "validateCampaignMod", "Campaign mod missing valid license field: %s", realModFilePathAndName
        validateCampaignMod = False
        Exit Function
    End If
    ' Warn if the license isn't recognized (from the list of suggested examples)
    RecognizedSPDXLicenseExpressions = Array( _
        "CC-BY-4.0 OR GPL-2.0-or-later", _
        "CC-BY-3.0 OR GPL-2.0-or-later", _
        "GPL-2.0-or-later", _
        "CC0-1.0")
    For k = 0 To UBound(RecognizedSPDXLicenseExpressions)
        If StrComp(RecognizedSPDXLicenseExpressions(k), modInfo.license, vbBinaryCompare) = 0 Then recognized = True
    Next
    If Not recognized Then
        debug_ 99, LOG_INFO, "validateCampaignMod", "Campaign mod has license ""%s"", consider one of the recommended options (ex. ""GPL-2.0-or-later""): %s", modInfo.license, realModFilePathAndName
        ' not fatal validation failure - continue
    End If

    ' Verify original campaign balance mods
    If modInfo.type_ = WzModType_CampaignBalanceMod Then
        ' It is expected that original campaign balance mods only modify the following directories:
        ' - stats/
        ' - script/
        ' Anything else should be packaged as an AlternateCampaign

        invalidCampaignBalanceModContents = False
        If WZ_PHYSFS_enumerateFiles(baseDir, files, nFiles) Then
            For k = 0 To nFiles - 1
                file = files(k)
                ' (if (file == nullptr) / if (*file == '\0'): return true;)
                If Len(file) = 0 Then GoTo continue_
                realFileName_platformIndependent = baseDir & "/" & file
                isDirectory = WZ_PHYSFS_isDirectory(realFileName_platformIndependent) <> 0
                If isDirectory Then
                    If file = "stats" Then
                        GoTo continue_          ' allowed - continue enumeration
                    ElseIf file = "script" Then
                        GoTo continue_          ' allowed - continue enumeration
                    Else
                        ' overriding other directories is not allowed
                        invalidCampaignBalanceModContents = True
                        Exit For                ' return false;
                    End If
                End If
continue_:
            Next                                ' return true; // continue enumeration
        End If

        If invalidCampaignBalanceModContents Then
            debug_ 139, LOG_INFO, "validateCampaignMod", "Campaign balance mod contains unsupported contents - treating as alternate campaign: %s", realModFilePathAndName
            modInfo.type_ = WzModType_AlternateCampaign
        End If
    End If

    ' Verify alternate campaign has at least one campaign file listed
    If modInfo.type_ = WzModType_AlternateCampaign Then
        If modInfo.campaignFiles_size = 0 Then
            debug_ 149, LOG_ERROR, "validateCampaignMod", "Alternate campaign mod must list one or more valid campaign JSON files in ""campaigns"": %s", realModFilePathAndName
            validateCampaignMod = False
            Exit Function
        End If
    End If

    validateCampaignMod = True
End Function

' static int loadCampaignModsAsyncImpl(void* data) (data: the LoadCampaingModsRequest, its fields the parameters)
Private Function loadCampaignModsAsyncImpl(ByVal versionedCampaignModsPath As String, ByVal versionedCampaignModsPath_platformDependent As String, ByVal completionHandler As WzFunc) As Long
    Dim availableMods As Collection, basePathFull As String, basePathFull_PlatformDependent As String, enumSuccess As Boolean
    Dim files() As String, nFiles As Long, k As Long, file As String, realFileName_platformIndependent As String
    Dim realFileName_platformDependent As String, isDirectory As Boolean, pRealDirStr As String, realFilePathAndName As String
    Dim modInfoPath As String, modInfo As WzCampaignModInfo, modBannerPath As String, bannerData() As Byte, bannerSize As Long

    ' (if (!pRequestInfo) return 1;: the request is always given)

    ' 1. Enumerate mods/campaign
    Set availableMods = New Collection
    basePathFull = versionedCampaignModsPath
    basePathFull_PlatformDependent = versionedCampaignModsPath_platformDependent
    enumSuccess = WZ_PHYSFS_enumerateFiles(basePathFull, files, nFiles)
    If enumSuccess Then
        For k = 0 To nFiles - 1
            file = files(k)
            ' (if (file == nullptr) / if (*file == '\0'): return true;)
            If Len(file) = 0 Then GoTo continue_
            realFileName_platformIndependent = basePathFull & "/" & file
            realFileName_platformDependent = basePathFull_PlatformDependent & PHYSFS_getDirSeparator() & file
            isDirectory = WZ_PHYSFS_isDirectory(realFileName_platformIndependent) <> 0
            If isDirectory Then
                GoTo continue_
            End If

            If Left$(file, 1) = "." Or Mid$(realFileName_platformIndependent, InStrRev(realFileName_platformIndependent, ".") + 1) <> "wz" Then
                GoTo continue_                  ' skip and continue
            End If

            ' if it's a .wz file, mount it in a temporary path
            pRealDirStr = PHYSFS_getRealDir(realFileName_platformIndependent)
            If StrPtr(pRealDirStr) = 0 Then
                debug_ 189, LOG_ERROR, "loadCampaignModsAsyncImpl", "Failed to find realdir for: %s", realFileName_platformIndependent
                GoTo continue_                  ' skip and continue
            End If
            realFilePathAndName = pRealDirStr & realFileName_platformDependent

            If PHYSFS_mount(realFilePathAndName, "WZTempProcessCamMod", PHYSFS_APPEND) = 0 Then
                debug_ 196, LOG_ERROR, "loadCampaignModsAsyncImpl", "Could not mount %s, because: %s." & vbLf & "Please delete or move the file specified.", realFilePathAndName, WZ_PHYSFS_getLastError()
                GoTo continue_                  ' skip and continue
            End If

            ' Process the mod .wz file (using the mounted path in WZTempProcessCamMod)
            modInfoPath = "WZTempProcessCamMod/mod-info.json"
            Set modInfo = loadCampaignModInfoFromFile(modInfoPath, "WZTempProcessCamMod")
            If Not modInfo Is Nothing Then
                ' Add the modFilename and modPath
                modInfo.modFilename = file
                modInfo.modPath = pRealDirStr & basePathFull_PlatformDependent

                ' attempt to load the mod_banner.png data, if it exists
                modBannerPath = "WZTempProcessCamMod/mod-banner.png"
                If PHYSFS_exists(modBannerPath) <> 0 Then
                    ' (loadFileToBufferVectorT<unsigned char>(modBannerPath, modBannerPNGData, false, false))
                    If loadFile(modBannerPath, bannerData, bannerSize, False) Then
                        modInfo.setModBannerPNGData bannerData, bannerSize
                    End If
                End If
            End If

            ' Validate the mod info
            If Not modInfo Is Nothing Then
                If Not validateCampaignMod(modInfo, "WZTempProcessCamMod", realFilePathAndName) Then
                    Set modInfo = Nothing       ' not loadable
                End If
            End If

            ' (WZ_PHYSFS_unmount: PHYSFS_unmount)
            If PHYSFS_unmount(realFilePathAndName) = 0 Then
                debug_ 228, LOG_ERROR, "loadCampaignModsAsyncImpl", "Could not unmount %s, %s", realFilePathAndName, WZ_PHYSFS_getLastError()
            End If

            If Not modInfo Is Nothing Then
                availableMods.Add modInfo
            End If
continue_:
        Next                                    ' return true; // continue enumeration
    End If
    If Not enumSuccess Then
        debug_ 240, LOG_INFO, "loadCampaignModsAsyncImpl", "Failed to enumerate available campaign mods?"
    End If
    WzFunc_call completionHandler, availableMods
    ' (delete pRequestInfo;)
    loadCampaignModsAsyncImpl = 0
End Function

' (VB) Public: WzCampaignSelectorTitleUI.cls's start calls it
Public Sub loadCampaignModsAsync(ByVal completionHandler As WzFunc)
    Dim versionedCampaignModsPath As String, versionedCampaignModsPath_platformDependent As String
    ' Build the configuration data
    ' (auto pRequest = new LoadCampaingModsRequest(): its fields)
    versionedCampaignModsPath = versionedModsPath(MODS_CAMPAIGN)
    versionedCampaignModsPath_platformDependent = convertToPlatformDependentPath(versionedCampaignModsPath)

    ' (WZ_THREAD * pLoadModsThread = wzThreadCreate(loadCampaignModsAsyncImpl, pRequest); wzThreadDetach: the thread
    ' is not ported, its function runs here; "Failed to create thread" cannot happen)
    loadCampaignModsAsyncImpl versionedCampaignModsPath, versionedCampaignModsPath_platformDependent, completionHandler
End Sub

' MARK: -

' WzCampaignModButton

Public Sub WzCampaignModButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_WzCampaignModButton ' (VB) virtual calls from here on reach WzCampaignModButton's overrides
    WzCampaignModButton_(self).isSelected = False
End Sub

Private Sub WzCampaignModButton_initialize(ByVal self As WIDGET, ByVal text As String)
    Dim minButtonWidthForText As Long
    W_BUTTON_(self).FontID = font_regular_bold
    self.setString text
    minButtonWidthForText = iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID)
    self.setGeometry 0, 0, minButtonWidthForText + 20, iV_GetTextLineSize(W_BUTTON_(self).FontID) + 16
End Sub

Public Function WzCampaignModButton_make(ByVal text As String) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignModButton)
    WzCampaignModButton_construct widget
    WzCampaignModButton_initialize widget, text
    Set WzCampaignModButton_make = widget
End Function

Public Sub WzCampaignModButton_setIsSelected(ByVal self As WIDGET, ByVal val_ As Boolean)
    WzCampaignModButton_(self).isSelected = val_
End Sub

Public Function WzCampaignModButton_getIsSelected(ByVal self As WIDGET) As Boolean
    WzCampaignModButton_getIsSelected = WzCampaignModButton_(self).isSelected
End Function

Public Sub WzCampaignModButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Const leftPadding As Long = 25
    Const rightPadding As Long = 10
    Dim d As WzCampaignModButton, x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, has_backColor As Boolean, backColor As Long
    Dim textColor As Long, textX0 As Long, textY0 As Long, maxTextDisplayableWidth As Long, isTruncated As Boolean
    Set d = WzCampaignModButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(W_BUTTON_(self).pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    has_backColor = False                       ' optional<PIELIGHT> backColor = nullopt;
    If d.isSelected Or isDown Then
        has_backColor = True
        backColor = psPalette(WZCOL_MENU_SCORE_BUILT)
    End If
    If has_backColor Then
        pie_UniTransBoxFill x0, y0, x1, y1, backColor
    End If
    If isHighlight Then
        ' do something?
    End If

    If haveText Then
        d.wzText.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID

        If isDown Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_FORM_TEXT)
        If isDisabled Then
            textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        End If

        textX0 = x0 + leftPadding
        textY0 = y0 + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()

        maxTextDisplayableWidth = self.width() - (leftPadding + rightPadding)
        isTruncated = maxTextDisplayableWidth < d.wzText.width()
        If isTruncated Then
            maxTextDisplayableWidth = maxTextDisplayableWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
        End If
        d.wzText.render textX0, textY0, textColor, 0!, maxTextDisplayableWidth
        If isTruncated Then
            ' Render ellipsis
            iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxTextDisplayableWidth + 2, textY0, textColor
        End If
    End If
End Sub

' WzCampaignListSection

Public Sub WzCampaignListSection_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzCampaignListSection ' (VB) virtual calls from here on reach WzCampaignListSection's overrides
End Sub

Public Function WzCampaignListSection_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignListSection)
    WzCampaignListSection_construct widget
    WzCampaignListSection_initialize widget
    Set WzCampaignListSection_make = widget
End Function

' (modButtons: a Collection of WzCampaignModButtons)
Public Sub WzCampaignListSection_updateList(ByVal self As WIDGET, ByVal modButtons As Collection)
    Dim d As WzCampaignListSection, k As Long
    Set d = WzCampaignListSection_(self)
    Set d.modButtons = modButtons
    ScrollableListWidget_clear d.sectionList

    For k = 1 To d.modButtons.Count
        ScrollableListWidget_addItem d.sectionList, d.modButtons(k)
    Next

    self.geometryChanged
End Sub

Public Sub WzCampaignListSection_geometryChanged(ByVal self As WIDGET)
    Dim d As WzCampaignListSection
    Set d = WzCampaignListSection_(self)
    d.sectionTitle.callCalcLayout
    d.sectionList.callCalcLayout
End Sub

Private Sub WzCampaignListSection_initialize(ByVal self As WIDGET)
    Dim d As WzCampaignListSection
    Set d = WzCampaignListSection_(self)
    Set d.sectionTitle = W_LABEL_new()
    W_LABEL_setFont d.sectionTitle, font_regular_bold, psPalette(WZCOL_TEXT_DARK)
    d.sectionTitle.setString "Additional Campaigns"
    W_LABEL_setCanTruncate d.sectionTitle, True
    self.attach d.sectionTitle
    d.sectionTitle.setCalcLayout WzFunc_new(FN_campaignListSection_titleCalcLayout)

    Set d.sectionList = ScrollableListWidget_make()
    self.attach d.sectionList
    d.sectionList.setCalcLayout WzFunc_new(FN_campaignListSection_listCalcLayout)
End Sub

Public Function WzCampaignListSection_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzCampaignListSection
    Set d = WzCampaignListSection_(self)
    WzCampaignListSection_idealHeight = d.sectionTitle.idealHeight() + 10 + d.sectionList.idealHeight()
End Function

Public Sub WzCampaignListSection_clearAllIsSelected(ByVal self As WIDGET)
    Dim d As WzCampaignListSection, k As Long
    Set d = WzCampaignListSection_(self)
    For k = 1 To d.modButtons.Count
        WzCampaignModButton_setIsSelected d.modButtons(k), False
    Next
End Sub

' WzCampaignSelectorForm

Public Sub WzCampaignSelectorForm_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_WzCampaignSelectorForm ' (VB) virtual calls from here on reach WzCampaignSelectorForm's overrides
End Sub

' (parent: the WzCampaignSelectorTitleUI's address, its std::weak_ptr)
Public Function WzCampaignSelectorForm_make(ByVal parent As Long) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignSelectorForm)
    WzCampaignSelectorForm_construct widget
    WzCampaignSelectorForm_initialize widget, parent
    Set WzCampaignSelectorForm_make = widget
End Function

Private Function WzCampaignSelectorForm_makeCampaignModButton(ByVal self As WIDGET, ByVal title_ As String, ByVal modInfo As WzCampaignModInfo) As WIDGET
    Dim button As WIDGET, weakParent As Long
    Set button = WzCampaignModButton_make(title_)
    weakParent = ObjPtr(self)               ' std::weak_ptr<WzCampaignSelectorForm>
    W_BUTTON_addOnClickHandler button, WzFunc_new(FN_campaignModButton_onClick, weakParent, modInfo)
    Set WzCampaignSelectorForm_makeCampaignModButton = button
End Function

Public Sub WzCampaignSelectorForm_geometryChanged(ByVal self As WIDGET)
    Dim d As WzCampaignSelectorForm
    Set d = WzCampaignSelectorForm_(self)
    d.originalCampaignButton.callCalcLayout
    d.alternateCampaigns.callCalcLayout
End Sub

Public Sub WzCampaignSelectorForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
End Sub

Private Sub WzCampaignSelectorForm_initialize(ByVal self As WIDGET, ByVal parent As Long)
    Dim d As WzCampaignSelectorForm
    Set d = WzCampaignSelectorForm_(self)
    d.parentCampaignSelectorUI = parent

    Set d.originalCampaignButton = WzCampaignSelectorForm_makeCampaignModButton(self, "Original Campaign", Nothing)
    self.attach d.originalCampaignButton
    d.originalCampaignButton.setCalcLayout WzFunc_new(FN_selectorForm_originalCalcLayout)
    WzCampaignModButton_setIsSelected d.originalCampaignButton, True

    Set d.alternateCampaigns = WzCampaignListSection_make()
    self.attach d.alternateCampaigns
    d.alternateCampaigns.setCalcLayout WzFunc_new(FN_selectorForm_alternateCalcLayout)
End Sub

' (availableMods: a Collection of WzCampaignModInfo)
Public Sub WzCampaignSelectorForm_updateList(ByVal self As WIDGET, ByVal availableMods As Collection)
    Dim sortedAlternateCampaigns As Collection, modItem As WzCampaignModInfo, k As Long, inserted As Boolean
    Dim modButtons As Collection
    Set sortedAlternateCampaigns = New Collection
    For Each modItem In availableMods
        If modItem.type_ <> WzModType_AlternateCampaign Then
            GoTo continue_
        End If
        ' (push_back, then std::sort by [](a, b) { return a.name < b.name; }: an insertion in order)
        inserted = False
        For k = 1 To sortedAlternateCampaigns.Count
            If StrComp(modItem.name_, sortedAlternateCampaigns(k).name_, vbBinaryCompare) < 0 Then
                sortedAlternateCampaigns.Add modItem, Before:=k
                inserted = True
                Exit For
            End If
        Next
        If Not inserted Then sortedAlternateCampaigns.Add modItem
continue_:
    Next

    Set modButtons = New Collection
    For Each modItem In sortedAlternateCampaigns
        modButtons.Add WzCampaignSelectorForm_makeCampaignModButton(self, modItem.name_, modItem)
    Next

    WzCampaignListSection_updateList WzCampaignSelectorForm_(self).alternateCampaigns, modButtons
End Sub

Public Sub WzCampaignSelectorForm_changeSelectedItem(ByVal self As WIDGET, ByVal modInfo As WzCampaignModInfo)
    Dim d As WzCampaignSelectorForm, strongParent As WzCampaignSelectorTitleUI
    Set d = WzCampaignSelectorForm_(self)
    Set strongParent = campaignSelectorTitleUI_lock(d.parentCampaignSelectorUI)
    If modInfo Is Nothing Then
        ' original campaign
        WzCampaignModButton_setIsSelected d.originalCampaignButton, True
        WzCampaignListSection_clearAllIsSelected d.alternateCampaigns
    Else
        WzCampaignModButton_setIsSelected d.originalCampaignButton, False
        WzCampaignListSection_clearAllIsSelected d.alternateCampaigns ' specific item is selected later by caller
    End If

    strongParent.displayStartOptionsForm modInfo
End Sub

' MARK: -

' CampaignStartButton

Public Sub CampaignStartButton_construct(ByVal self As WIDGET)
    W_BUTTON_construct self
    self.kind = K_CampaignStartButton ' (VB) virtual calls from here on reach CampaignStartButton's overrides
End Sub

Public Function CampaignStartButton_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_CampaignStartButton)
    CampaignStartButton_construct widget
    CampaignStartButton_initialize widget
    Set CampaignStartButton_make = widget
End Function

Private Sub CampaignStartButton_initialize(ByVal self As WIDGET)
    Dim minButtonWidthForText As Long
    self.setString "Start Game"
    W_BUTTON_(self).FontID = font_large
    minButtonWidthForText = iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID)
    self.setGeometry 0, 0, minButtonWidthForText + 10, iV_GetTextLineSize(W_BUTTON_(self).FontID) + 10
End Sub

Public Sub CampaignStartButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As CampaignStartButton, x0 As Long, y0 As Long, x1 As Long, y1 As Long, haveText As Boolean
    Dim isDown As Boolean, isDisabled As Boolean, isHighlight As Boolean, border_color As Long, background_color As Long
    Dim fw As Long, fx As Long, fy As Long, textColor As Long
    Set d = CampaignStartButton_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    haveText = Len(W_BUTTON_(self).pText) <> 0

    isDown = (self.getState() And (WBUT_DOWN Or WBUT_LOCK Or WBUT_CLICKLOCK)) <> 0
    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And ((self.getState() And WBUT_HIGHLIGHT) <> 0)

    ' Display the button.
    If isHighlight Then border_color = psPalette(WZCOL_TEXT_BRIGHT) Else border_color = psPalette(WZCOL_TEXT_MEDIUM)
    background_color = psPalette(WZCOL_TRANSPARENT_BOX)
    If isDisabled Then
        border_color = psPalette(WZCOL_TEXT_DARK)
        background_color = pielightWithAlpha(background_color, pielightAlpha(background_color) \ 2)
    ElseIf isDown Then
        background_color = pal_RGBA(0, 0, 0, 150)
    End If
    pie_UniTransBoxFill x0, y0, x1, y1, background_color
    iV_Box2 x0, y0, x1, y1, border_color, border_color

    If haveText Then
        d.wzText.setText W_BUTTON_(self).pText, W_BUTTON_(self).FontID
        fw = d.wzText.width()
        fx = x0 + (self.width() - fw) \ 2
        fy = y0 + (self.height() - d.wzText.lineSize()) \ 2 - d.wzText.aboveBase()
        textColor = psPalette(WZCOL_TEXT_BRIGHT)
        If isDisabled Then
            textColor = psPalette(WZCOL_TEXT_DARK)
            textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
        End If
        d.wzText.render fx, fy, textColor
    End If
End Sub

' WzCampaignOptionButton

Public Sub WzCampaignOptionButton_construct(ByVal self As WIDGET)
    Dim d As WzCampaignOptionButton
    W_BUTTON_construct self
    self.kind = K_WzCampaignOptionButton ' (VB) virtual calls from here on reach WzCampaignOptionButton's overrides
    Set d = WzCampaignOptionButton_(self)
    d.descriptionFontID = font_small
    d.cachedIdealWidth = 0
    d.lastWidgetWidth = 0
    d.horizontalPadding = 10
    d.verticalPadding = 5
    d.internalPadding = 3
End Sub

Private Sub WzCampaignOptionButton_initialize(ByVal self As WIDGET)
    W_BUTTON_(self).FontID = font_regular_bold
    ' Enable right clicks
    self.style = self.style Or WBUT_SECONDARY
End Sub

Public Function WzCampaignOptionButton_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignOptionButton)
    WzCampaignOptionButton_construct widget
    WzCampaignOptionButton_initialize widget
    Set WzCampaignOptionButton_make = widget
End Function

Public Sub WzCampaignOptionButton_setDescriptionText(ByVal self As WIDGET, ByVal description As String)
    Dim d As WzCampaignOptionButton
    Set d = WzCampaignOptionButton_(self)
    d.wzDescriptionText.setText description, d.descriptionFontID
    WzCampaignOptionButton_recalcIdealWidth self
End Sub

Private Sub WzCampaignOptionButton_recalcIdealWidth(ByVal self As WIDGET)
    Dim d As WzCampaignOptionButton
    Set d = WzCampaignOptionButton_(self)
    d.cachedIdealWidth = (d.horizontalPadding * 2) + maxL(iV_GetTextWidth(W_BUTTON_(self).pText, W_BUTTON_(self).FontID), d.wzDescriptionText.width())
End Sub

Public Function WzCampaignOptionButton_idealWidth(ByVal self As WIDGET) As Long
    WzCampaignOptionButton_idealWidth = WzCampaignOptionButton_(self).cachedIdealWidth
End Function

Public Function WzCampaignOptionButton_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzCampaignOptionButton
    Set d = WzCampaignOptionButton_(self)
    WzCampaignOptionButton_idealHeight = (d.verticalPadding * 2) + iV_GetTextLineSize(W_BUTTON_(self).FontID) + d.internalPadding + iV_GetTextLineSize(d.descriptionFontID)
End Function

Public Sub WzCampaignOptionButton_setString(ByVal self As WIDGET, ByVal string_ As String)
    W_BUTTON_setString self, string_
    WzCampaignOptionButton_recalcIdealWidth self
End Sub

Public Sub WzCampaignOptionButton_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzCampaignOptionButton, x0 As Long, y0 As Long, w As Long, isDisabled As Boolean, isHighlight As Boolean
    Dim fontID As Long, result As iV_fonts, availableButtonTextWidth As Long, textColor As Long, textX0 As Long, textY0 As Long
    Dim maxTextDisplayableWidth As Long, maxDisplayableMainTextWidth As Long, isTruncated As Boolean, pText As String
    Dim descriptionColor As Long, descriptionX0 As Long, descriptionY0 As Long, maxDisplayableDescriptionTextWidth As Long
    Set d = WzCampaignOptionButton_(self)
    pText = W_BUTTON_(self).pText
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    w = self.width()

    isDisabled = (self.getState() And WBUT_DISABLE) <> 0
    isHighlight = Not isDisabled And self.isMouseOverWidget()

    fontID = d.wzText.getFontID()
    If fontID = font_count Or d.lastWidgetWidth <> self.width() Or d.wzText.getText() <> pText Then
        fontID = W_BUTTON_(self).FontID
    End If
    d.wzText.setText pText, fontID

    availableButtonTextWidth = w - (d.horizontalPadding * 2)
    If d.wzText.width() > availableButtonTextWidth Then
        ' text would exceed the bounds of the button area
        ' try to shrink font so it fits
        Do
            If fontID = font_small Or fontID = font_bar Then
                Exit Do
            End If
            If Not iV_ShrinkFont(fontID, result) Then
                Exit Do
            End If
            fontID = result
            d.wzText.setText pText, fontID
        Loop While d.wzText.width() > availableButtonTextWidth
    End If
    d.lastWidgetWidth = self.width()

    ' Draw the main text
    If isHighlight Then textColor = psPalette(WZCOL_TEXT_BRIGHT) Else textColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        textColor = pielightWithAlpha(textColor, pielightAlpha(textColor) \ 2)
    End If

    textX0 = x0 + d.horizontalPadding
    textY0 = y0 + d.verticalPadding - d.wzText.aboveBase()

    maxTextDisplayableWidth = w - (d.horizontalPadding * 2)
    maxDisplayableMainTextWidth = maxTextDisplayableWidth
    isTruncated = maxDisplayableMainTextWidth < d.wzText.width()
    If isTruncated Then
        maxDisplayableMainTextWidth = maxDisplayableMainTextWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If
    d.wzText.render textX0, textY0, textColor, 0!, maxDisplayableMainTextWidth
    If isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxDisplayableMainTextWidth + 2, textY0, textColor
    End If

    ' Draw the description (truncate if needed)
    If isHighlight Then descriptionColor = psPalette(WZCOL_TEXT_BRIGHT) Else descriptionColor = psPalette(WZCOL_TEXT_MEDIUM)
    If isDisabled Then
        descriptionColor = pielightWithAlpha(descriptionColor, pielightAlpha(descriptionColor) \ 2)
    End If
    descriptionX0 = textX0
    descriptionY0 = y0 + d.verticalPadding + d.wzText.lineSize() + d.internalPadding - d.wzDescriptionText.aboveBase()
    maxDisplayableDescriptionTextWidth = maxTextDisplayableWidth
    isTruncated = maxDisplayableDescriptionTextWidth < d.wzDescriptionText.width()
    If isTruncated Then
        maxDisplayableDescriptionTextWidth = maxDisplayableDescriptionTextWidth - (iV_GetEllipsisWidth(d.wzDescriptionText.getFontID()) + 2)
    End If
    d.wzDescriptionText.render descriptionX0, descriptionY0, descriptionColor, 0!, maxDisplayableDescriptionTextWidth
    If isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzDescriptionText.getFontID(), descriptionX0 + maxDisplayableDescriptionTextWidth + 2, descriptionY0, descriptionColor
    End If
End Sub

' class WzCampaignTweakOptionSetting's constructor (WzCampaignTweakOptionSetting.cls)
Private Function WzCampaignTweakOptionSetting_new(ByVal uniqueIdentifier As String, ByVal displayName As String, ByVal description As String, ByVal defaultValue As Boolean, ByVal editable As Boolean) As WzCampaignTweakOptionSetting
    Dim o As WzCampaignTweakOptionSetting
    Set o = New WzCampaignTweakOptionSetting
    o.uniqueIdentifier = uniqueIdentifier
    o.displayName = displayName
    o.description = description
    o.defaultValue = defaultValue
    o.currentValue = defaultValue
    o.editable = editable
    Set WzCampaignTweakOptionSetting_new = o
End Function

' WzCampaignTweakOptionsEditForm

Public Sub WzCampaignTweakOptionsEditForm_construct(ByVal self As WIDGET)
    Dim d As WzCampaignTweakOptionsEditForm
    WIDGET_construct self
    self.kind = K_WzCampaignTweakOptionsEditForm ' (VB) virtual calls from here on reach WzCampaignTweakOptionsEditForm's overrides
    Set d = WzCampaignTweakOptionsEditForm_(self)
    d.innerPadding = 20
    d.betweenButtonPadding = 15
End Sub

' (tweakOptions: the shared Collection of WzCampaignTweakOptionSetting; closeHandlerFunc: Nothing = empty)
Public Function WzCampaignTweakOptionsEditForm_make(ByVal tweakOptions As Collection, ByVal closeHandlerFunc As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignTweakOptionsEditForm)
    WzCampaignTweakOptionsEditForm_construct widget
    WzCampaignTweakOptionsEditForm_initialize widget, tweakOptions, closeHandlerFunc
    Set WzCampaignTweakOptionsEditForm_make = widget
End Function

Private Sub WzCampaignTweakOptionsEditForm_initialize(ByVal self As WIDGET, ByVal tweakOptions As Collection, ByVal closeHandlerFunc As WzFunc)
    Dim d As WzCampaignTweakOptionsEditForm, weakEditForm As Long, i As Long, o As WzCampaignTweakOptionSetting
    Dim toggleWidget As WIDGET, listY0 As Long
    Set d = WzCampaignTweakOptionsEditForm_(self)
    Set d.tweakOptions = tweakOptions

    Set d.formTitle = W_LABEL_new()
    W_LABEL_setFont d.formTitle, font_regular_bold, psPalette(WZCOL_TEXT_MEDIUM)
    d.formTitle.setString "Tweak Options:"
    W_LABEL_setCanTruncate d.formTitle, True
    d.formTitle.setTransparentToMouse True
    d.formTitle.setGeometry d.innerPadding, d.innerPadding, W_LABEL_getMaxLineWidth(d.formTitle), iV_GetTextLineSize(font_regular_bold)
    self.attach d.formTitle

    weakEditForm = ObjPtr(self)

    If Not closeHandlerFunc Is Nothing Then
        ' Add "Close" button
        Set d.closeButton = WzFrontendImageButton_make(False, 0)
        d.closeButton.setString "Close"
        W_BUTTON_addOnClickHandler d.closeButton, WzFunc_new(FN_campaign_closeButtonClick, closeHandlerFunc)
        self.attach d.closeButton
    End If

    ' Add "Reset to Defaults" button
    Set d.resetButton = WzFrontendImageButton_make(True, IMAGE_ARROW_UNDO)
    WzFrontendImageButton_setImageDimensions d.resetButton, 12
    WzFrontendImageButton_setPadding d.resetButton, 6, 4
    d.resetButton.setTip "Reset to Defaults"
    W_BUTTON_addOnClickHandler d.resetButton, WzFunc_new(FN_tweakEditForm_resetClick, weakEditForm)
    self.attach d.resetButton
    WzCampaignTweakOptionsEditForm_updateResetButtonStatus self

    Set d.tweaksDisplayList = ScrollableListWidget_make()
    ScrollableListWidget_setItemSpacing d.tweaksDisplayList, 10
    ScrollableListWidget_setExpandWhenScrollbarInvisible d.tweaksDisplayList, False

    For i = 0 To d.tweakOptions.Count - 1
        Set o = d.tweakOptions(i + 1)

        If Not o.editable Then
            GoTo continue_                      ' skip non-editable settings
        End If

        Set toggleWidget = WzAdvCheckbox_make(o.displayName, o.description)  ' WzCampaignTweakOptionToggle::make
        WzAdvCheckbox_setIsChecked toggleWidget, o.currentValue
        W_BUTTON_addOnClickHandler toggleWidget, WzFunc_new(FN_tweakEditForm_toggleClick, weakEditForm, i)

        d.toggleWidgets.Add toggleWidget
        ScrollableListWidget_addItem d.tweaksDisplayList, toggleWidget
        d.tweakOptionIdxToToggleWidgetIdxMap.Add d.toggleWidgets.Count - 1, "i" & i
continue_:
    Next

    listY0 = d.formTitle.y() + d.formTitle.height() + d.innerPadding
    d.tweaksDisplayList.setGeometry d.innerPadding, listY0, d.tweaksDisplayList.idealWidth(), 100
    self.attach d.tweaksDisplayList
End Sub

Public Sub WzCampaignTweakOptionsEditForm_geometryChanged(ByVal self As WIDGET)
    Dim d As WzCampaignTweakOptionsEditForm, childWidth As Long, k As Long, widg As WIDGET, listY0 As Long, listH As Long
    Dim listW As Long, listIdealHeight As Long, lastButtonX0 As Long, buttonY0 As Long, topButtonHeight As Long, buttonX0 As Long
    Set d = WzCampaignTweakOptionsEditForm_(self)
    If self.width() <= 1 Or self.height() <= 1 Then
        Exit Sub
    End If

    ' before sending the event to all children, resize all of the toggleWidgets to the new width (which causes them to recalculate their new needed height)
    childWidth = self.width() - (d.innerPadding * 2) - ScrollableListWidget_getScrollbarWidth(d.tweaksDisplayList)
    For k = 1 To d.toggleWidgets.Count
        Set widg = d.toggleWidgets(k)
        ' first call sets the width, and causes the internal paragraph to recalculate itself and its needed height
        widg.setGeometry widg.x(), widg.y(), childWidth, widg.height()
        ' second call sets the height to the needed height (including the recalculated paragraph)
        widg.setGeometry widg.x(), widg.y(), childWidth, widg.idealHeight()
    Next

    ' recalculate position of tweaks list
    listY0 = d.tweaksDisplayList.y()
    listH = self.height() - listY0 - d.innerPadding
    listW = self.width() - (d.innerPadding * 2)
    If listW > 1 And listH > 1 Then
        d.tweaksDisplayList.setGeometry d.innerPadding, listY0, listW, listH

        ' truncate list height at needed height (if scrollable height is less)
        listIdealHeight = d.tweaksDisplayList.idealHeight()
        If listIdealHeight < listH Then
            d.tweaksDisplayList.setGeometry d.innerPadding, listY0, listW, listIdealHeight
        End If
    End If

    ' position top buttons
    lastButtonX0 = self.width() - d.innerPadding
    buttonY0 = d.innerPadding
    topButtonHeight = maxL(d.closeButton.idealHeight(), d.resetButton.idealHeight())
    If Not d.closeButton Is Nothing Then
        buttonX0 = lastButtonX0 - d.closeButton.idealWidth()
        d.closeButton.setGeometry buttonX0, buttonY0, d.closeButton.idealWidth(), topButtonHeight
        lastButtonX0 = buttonX0 - d.betweenButtonPadding
    End If
    If Not d.resetButton Is Nothing Then
        buttonX0 = lastButtonX0 - d.resetButton.idealWidth()
        d.resetButton.setGeometry buttonX0, buttonY0, d.resetButton.idealWidth(), topButtonHeight
        lastButtonX0 = buttonX0
    End If
End Sub

' ideal height *for the current width*
Public Function WzCampaignTweakOptionsEditForm_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzCampaignTweakOptionsEditForm
    Set d = WzCampaignTweakOptionsEditForm_(self)
    WzCampaignTweakOptionsEditForm_idealHeight = d.tweaksDisplayList.y() + d.tweaksDisplayList.height() + d.innerPadding
End Function

Public Sub WzCampaignTweakOptionsEditForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, backColor As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    backColor = psPalette(WZCOL_TRANSPARENT_BOX)
    backColor = pielightWithAlpha(backColor, maxL(220, pielightAlpha(backColor)))
    pie_UniTransBoxFill x0, y0, x1, y1, backColor

    iV_Box x0, y0, x1, y1, pal_RGBA(255, 255, 255, 175)
End Sub

Private Sub WzCampaignTweakOptionsEditForm_resetToDefaults(ByVal self As WIDGET)
    Dim d As WzCampaignTweakOptionsEditForm, i As Long, o As WzCampaignTweakOptionSetting, idx As Long
    Set d = WzCampaignTweakOptionsEditForm_(self)
    For i = 0 To d.tweakOptions.Count - 1
        Set o = d.tweakOptions(i + 1)
        o.currentValue = o.defaultValue
        If collectionHasKey(d.tweakOptionIdxToToggleWidgetIdxMap, "i" & i) Then
            idx = d.tweakOptionIdxToToggleWidgetIdxMap("i" & i)
            If Not (idx < d.toggleWidgets.Count) Then ASSERT_FAILURE "it->second < toggleWidgets.size()", "src/titleui/campaign.cpp:989", "resetToDefaults", "Invalid index: %zu", idx
            WzAdvCheckbox_setIsChecked d.toggleWidgets(idx + 1), o.defaultValue
        End If
    Next
    WzCampaignTweakOptionsEditForm_updateResetButtonStatus self
End Sub

Private Sub WzCampaignTweakOptionsEditForm_updateResetButtonStatus(ByVal self As WIDGET)
    Dim d As WzCampaignTweakOptionsEditForm, hasSettingThatDiffersFromDefaults As Boolean, k As Long, o As WzCampaignTweakOptionSetting
    Set d = WzCampaignTweakOptionsEditForm_(self)
    For k = 1 To d.tweakOptions.Count
        Set o = d.tweakOptions(k)
        If o.currentValue <> o.defaultValue Then
            hasSettingThatDiffersFromDefaults = True
            Exit For
        End If
    Next
    If hasSettingThatDiffersFromDefaults Then d.resetButton.setState 0 Else d.resetButton.setState WBUT_DISABLE
End Sub

' WidgetInfoButtonWrapper

Public Sub WidgetInfoButtonWrapper_construct(ByVal self As WIDGET)
    Dim d As WidgetInfoButtonWrapper
    WIDGET_construct self
    self.kind = K_WidgetInfoButtonWrapper ' (VB) virtual calls from here on reach WidgetInfoButtonWrapper's overrides
    Set d = WidgetInfoButtonWrapper_(self)
    d.topPadding = 0
    d.bottomPadding = 0
    d.internalHorizontalPadding = 10
    d.imageDimensions = 16
End Sub

Public Function WidgetInfoButtonWrapper_make(ByVal widget As WIDGET) As WIDGET
    Dim result As WIDGET
    Set result = WIDGET_alloc(K_WidgetInfoButtonWrapper)
    WidgetInfoButtonWrapper_construct result
    WidgetInfoButtonWrapper_initialize result, widget
    Set WidgetInfoButtonWrapper_make = result
End Function

Public Function WidgetInfoButtonWrapper_idealWidth(ByVal self As WIDGET) As Long
    Dim d As WidgetInfoButtonWrapper
    Set d = WidgetInfoButtonWrapper_(self)
    WidgetInfoButtonWrapper_idealWidth = d.widget.idealWidth() + d.internalHorizontalPadding + d.infoButton.idealWidth()
End Function

Public Function WidgetInfoButtonWrapper_idealHeight(ByVal self As WIDGET) As Long
    WidgetInfoButtonWrapper_idealHeight = WidgetInfoButtonWrapper_(self).widget.idealHeight()
End Function

Private Sub WidgetInfoButtonWrapper_initialize(ByVal self As WIDGET, ByVal widget As WIDGET)
    Dim d As WidgetInfoButtonWrapper, weakSelf As Long
    Set d = WidgetInfoButtonWrapper_(self)
    Set d.widget = widget

    Set d.infoButton = WzInfoButton_make()
    WzFrontendImageButton_setImageDimensions d.infoButton, d.imageDimensions
    self.attach d.infoButton
    d.infoButton.hide
    weakSelf = ObjPtr(self)
    W_BUTTON_addOnClickHandler d.infoButton, WzFunc_new(FN_infoWrapper_infoClick, weakSelf)

    self.attach widget
End Sub

Public Sub WidgetInfoButtonWrapper_geometryChanged(ByVal self As WIDGET)
    Dim d As WidgetInfoButtonWrapper, w As Long, h As Long, infoButtonWidth As Long, infoButtonX0 As Long
    Dim infoButtonHeight As Long, widgetWidth As Long, widgetHeight As Long
    Set d = WidgetInfoButtonWrapper_(self)
    w = self.width()
    h = self.height()

    infoButtonWidth = d.imageDimensions + d.internalHorizontalPadding
    infoButtonX0 = self.width() - infoButtonWidth
    infoButtonHeight = self.height() - (d.topPadding + d.bottomPadding)
    d.infoButton.setGeometry infoButtonX0, d.topPadding, infoButtonWidth, infoButtonHeight

    widgetWidth = w - d.infoButton.width() - d.internalHorizontalPadding
    widgetHeight = h - (d.topPadding + d.bottomPadding)
    d.widget.setGeometry 0, d.topPadding, widgetWidth, widgetHeight
End Sub

' (func: Nothing = an empty std::function)
Public Sub WidgetInfoButtonWrapper_setInfoClickHandler(ByVal self As WIDGET, ByVal func As WzFunc)
    Dim d As WidgetInfoButtonWrapper
    Set d = WidgetInfoButtonWrapper_(self)
    Set d.infoClickHandler = func
    If Not d.infoClickHandler Is Nothing Then
        d.infoButton.show
    Else
        d.infoButton.hide
    End If
End Sub

' WzCampaignTweakOptionsSummaryWidget

Public Sub WzCampaignTweakOptionsSummaryWidget_construct(ByVal self As WIDGET)
    Dim d As WzCampaignTweakOptSummaryData
    WIDGET_construct self
    self.kind = K_WzCampaignTweakOptionsSummaryWidget ' (VB) virtual calls from here on reach its overrides
    Set d = WzCampaignTweakOptionsSummaryWidget_(self)
    d.betweenRowPadding = 0
    d.topPadding = 8
    d.bottomPadding = 5
    d.internalHorizontalPadding = 10
End Sub

Public Function WzCampaignTweakOptionsSummaryWidget_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignTweakOptionsSummaryWidget)
    WzCampaignTweakOptionsSummaryWidget_construct widget
    WzCampaignTweakOptionsSummaryWidget_initialize widget
    Set WzCampaignTweakOptionsSummaryWidget_make = widget
End Function

Public Function WzCampaignTweakOptionsSummaryWidget_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzCampaignTweakOptSummaryData, rowHeight As Long
    Set d = WzCampaignTweakOptionsSummaryWidget_(self)
    ' the height for two rows
    rowHeight = iV_GetTextLineSize(font_small)
    WzCampaignTweakOptionsSummaryWidget_idealHeight = (d.topPadding + d.bottomPadding) + (rowHeight * 2) + d.betweenRowPadding
End Function

Public Sub WzCampaignTweakOptionsSummaryWidget_geometryChanged(ByVal self As WIDGET)
    Dim i As Long
    For i = 0 To self.childCount() - 1
        self.childAt(i).callCalcLayout
    Next
End Sub

Public Sub WzCampaignTweakOptionsSummaryWidget_setInfoClickHandler(ByVal self As WIDGET, ByVal func As WzFunc)
    Dim d As WzCampaignTweakOptSummaryData
    Set d = WzCampaignTweakOptionsSummaryWidget_(self)
    Set d.infoClickHandler = func
    If Not d.infoClickHandler Is Nothing Then
        d.infoButton.show
    Else
        d.infoButton.hide
    End If
End Sub

' FixedIdealWidthLabel

Public Sub FixedIdealWidthLabel_construct(ByVal self As WIDGET, Optional ByVal fixedIdealWidth As Long = -1)
    W_LABEL_construct self
    self.kind = K_FixedIdealWidthLabel ' (VB) virtual calls from here on reach FixedIdealWidthLabel's overrides
    FixedIdealWidthLabel_(self).m_fixedIdealWidth = fixedIdealWidth
End Sub

Public Function FixedIdealWidthLabel_new(Optional ByVal fixedIdealWidth As Long = -1) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_FixedIdealWidthLabel)
    FixedIdealWidthLabel_construct widget, fixedIdealWidth
    Set FixedIdealWidthLabel_new = widget
End Function

Public Function FixedIdealWidthLabel_idealWidth(ByVal self As WIDGET) As Long
    Dim d As FixedIdealWidthLabel
    Set d = FixedIdealWidthLabel_(self)
    If d.m_fixedIdealWidth >= 0 Then
        FixedIdealWidthLabel_idealWidth = d.m_fixedIdealWidth
    Else
        FixedIdealWidthLabel_idealWidth = W_LABEL_idealWidth(self)
    End If
End Function

' WzCampaignTweakOptionsSummaryWidget::update's lambda makeTweakOptionSmallLabel
Private Function makeTweakOptionSmallLabel(ByVal str_ As String, Optional ByVal fixedIdealWidth As Long = -1) As WIDGET
    Dim label As WIDGET
    Set label = FixedIdealWidthLabel_new(fixedIdealWidth)
    W_LABEL_setFont label, font_small, psPalette(WZCOL_TEXT_MEDIUM)
    label.setString str_
    W_LABEL_setCanTruncate label, True
    label.setTransparentToMouse True
    label.setGeometry 0, 0, W_LABEL_getMaxLineWidth(label), iV_GetTextLineSize(font_small)
    Set makeTweakOptionSmallLabel = label
End Function

' (tweakOptionSettings: a Collection of WzCampaignTweakOptionSetting)
Public Sub WzCampaignTweakOptionsSummaryWidget_update(ByVal self As WIDGET, ByVal tweakOptionSettings As Collection)
    Dim d As WzCampaignTweakOptSummaryData, grid As WIDGET, row As grid_allocation_slot, col As Long
    Dim statusLabelWidth As Long, k As Long, setting As WzCampaignTweakOptionSetting, s As String
    Dim newStatusLabel As WIDGET, newTweakLabel As WIDGET
    Set d = WzCampaignTweakOptionsSummaryWidget_(self)

    Set d.tweakOptionsLabels = New Collection   ' tweakOptionsLabels.clear();

    Set grid = GridLayout_new()
    row = grid_allocation_slot_(0)
    col = 0

    ' ("U+2713": the check mark)
    statusLabelWidth = maxL(iV_GetTextWidth(ChrW$(&H2713), font_small), iV_GetTextWidth("x", font_small))

    ' Add list of tweak options (2 columns)
    For k = 1 To tweakOptionSettings.Count
        Set setting = tweakOptionSettings(k)
        If Not setting.editable Then
            GoTo continue_                      ' skip non-editable settings
        End If
        If setting.currentValue Then s = ChrW$(&H2713) Else s = "x"
        Set newStatusLabel = makeTweakOptionSmallLabel(s, statusLabelWidth)
        Set newTweakLabel = makeTweakOptionSmallLabel(setting.displayName)
        GridLayout_place grid, grid_allocation_slot_(col, 1, False), row, Margin_wrap(0, 3, 0, 0, newStatusLabel)
        col = col + 1
        GridLayout_place grid, grid_allocation_slot_(col, 1, True), row, newTweakLabel
        col = col + 1
        If col >= 3 Then
            row.start = row.start + 1
            col = 0
        End If
        d.tweakOptionsLabels.Add newStatusLabel
        d.tweakOptionsLabels.Add newTweakLabel
continue_:
    Next

    grid.setGeometry 0, 0, FRONTEND_BUTWIDTH, grid.idealHeight()
    grid.setTransparentToMouse True

    ScrollableListWidget_clear d.tweakOptionsList
    ScrollableListWidget_addItem d.tweakOptionsList, grid
End Sub

Private Sub WzCampaignTweakOptionsSummaryWidget_initialize(ByVal self As WIDGET)
    Dim d As WzCampaignTweakOptSummaryData, weakSelf As Long
    Set d = WzCampaignTweakOptionsSummaryWidget_(self)
    Set d.infoButton = WzInfoButton_make()
    WzFrontendImageButton_setImageDimensions d.infoButton, 16
    self.attach d.infoButton
    d.infoButton.setCalcLayout WzFunc_new(FN_tweakSummary_infoCalcLayout)
    d.infoButton.hide
    weakSelf = ObjPtr(self)
    W_BUTTON_addOnClickHandler d.infoButton, WzFunc_new(FN_tweakSummary_infoClick, weakSelf)

    Set d.tweakOptionsList = ClickableScrollableList_make()
    self.attach d.tweakOptionsList
    d.tweakOptionsList.setCalcLayout WzFunc_new(FN_tweakSummary_listCalcLayout)
    ClickableScrollableList_setOnClickHandler d.tweakOptionsList, WzFunc_new(FN_tweakSummary_listClick, weakSelf)
    ClickableScrollableList_setOnHighlightHandler d.tweakOptionsList, WzFunc_new(FN_tweakSummary_listHighlight, weakSelf)
End Sub

' MARK: - WzCampaignTitleWidget

Public Sub WzCampaignTitleWidget_construct(ByVal self As WIDGET)
    Dim d As WzCampaignTitleWidget
    WIDGET_construct self
    self.kind = K_WzCampaignTitleWidget ' (VB) virtual calls from here on reach WzCampaignTitleWidget's overrides
    Set d = WzCampaignTitleWidget_(self)
    d.leftPadding = 40
    d.topPadding = 20
    d.titleAuthorPadding = 3
End Sub

Private Sub WzCampaignTitleWidget_initialize(ByVal self As WIDGET, ByVal modInfo As WzCampaignModInfo)
    Dim d As WzCampaignTitleWidget, displayName As String, author As String, description As String, version As String
    Dim bannerData() As Byte, n As Long, ivImage As iV_Image, resultText As String, nextLineY0 As Long
    Dim authorLabel As WIDGET, descriptionWidget As WIDGET, scrollableDescriptionWidget As WIDGET
    Set d = WzCampaignTitleWidget_(self)
    If Not modInfo Is Nothing Then
        ' load details from mod
        displayName = modInfo.name_
        author = modInfo.author
        description = modInfo.description         ' description.getLocalizedString()
        n = modInfo.getModBannerPNGData(bannerData)
        If n > 0 Then
            Set ivImage = New iV_Image
            resultText = iV_loadImage_PNG_mem(bannerData, ivImage)
            If Len(resultText) = 0 Then         ' result.noError()
                Set d.pBackgroundTexture = gfx_api_context_loadTextureFromUncompressedImage(ivImage, texture_type_user_interface, "mem::mod_banner_png")
            Else
                debug_ 1303, LOG_ERROR, "initialize", "Failed to load image from memory buffer: %s", resultText
            End If
        End If
        If modInfo.hasVersion Then version = modInfo.version Else version = ""  ' version.value_or("")
    Else
        ' load built-in details for original campaign
        displayName = "Original Campaign"
        author = "Warzone 2100"
        description = "Command the forces of The Project in a battle to rebuild the world"
        version = version_getVersionString()
        Set d.pBackgroundTexture = gfx_api_context_loadTextureFromFile("images/wz_original_campaign_banner.png", texture_type_user_interface, -1, -1, True)
    End If

    ' Create all the widgets and position them

    Set d.nameLabel = W_LABEL_new()
    W_LABEL_setFont d.nameLabel, font_large, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setCanTruncate d.nameLabel, True
    d.nameLabel.setString displayName
    d.nameLabel.setGeometry d.leftPadding, d.topPadding, W_LABEL_getMaxLineWidth(d.nameLabel), iV_GetTextLineSize(font_large)
    self.attach d.nameLabel
    d.nameLabel.setCalcLayout WzFunc_new(FN_titleWidget_labelCalcLayout)

    nextLineY0 = d.nameLabel.y() + d.nameLabel.height() + d.titleAuthorPadding
    Set authorLabel = W_LABEL_new()
    W_LABEL_setFont authorLabel, font_regular_bold, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setCanTruncate authorLabel, True
    authorLabel.setString author
    authorLabel.setGeometry d.leftPadding, nextLineY0, W_LABEL_getMaxLineWidth(authorLabel), iV_GetTextLineSize(font_regular_bold)
    self.attach authorLabel
    authorLabel.setCalcLayout WzFunc_new(FN_titleWidget_labelCalcLayout)

    nextLineY0 = authorLabel.y() + authorLabel.height() + 20
    Set descriptionWidget = Paragraph_new()
    Paragraph_setFont descriptionWidget, font_regular
    Paragraph_setFontColour descriptionWidget, psPalette(WZCOL_TEXT_BRIGHT)
    If Not d.pBackgroundTexture Is Nothing Then
        Paragraph_setShadeColour descriptionWidget, pal_RGBA(0, 0, 0, 80)
    End If
    descriptionWidget.setGeometry d.leftPadding, nextLineY0, 400, 40
    Paragraph_addText descriptionWidget, description

    Set scrollableDescriptionWidget = ScrollableListWidget_make()
    ScrollableListWidget_addItem scrollableDescriptionWidget, descriptionWidget
    scrollableDescriptionWidget.setGeometry d.leftPadding, nextLineY0, 400, 45
    self.attach scrollableDescriptionWidget

    scrollableDescriptionWidget.setCalcLayout WzFunc_new(FN_titleWidget_descriptionCalcLayout, descriptionWidget)
End Sub

' ~WzCampaignTitleWidget()
Public Sub WzCampaignTitleWidget_destruct(ByVal self As WIDGET)
    Dim d As WzCampaignTitleWidget
    Set d = WzCampaignTitleWidget_(self)
    If Not d.pBackgroundTexture Is Nothing Then Set d.pBackgroundTexture = Nothing
End Sub

' (modInfo: Nothing = nullopt)
Public Function WzCampaignTitleWidget_make(ByVal modInfo As WzCampaignModInfo) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignTitleWidget)
    WzCampaignTitleWidget_construct widget
    WzCampaignTitleWidget_initialize widget, modInfo
    Set WzCampaignTitleWidget_make = widget
End Function

Public Sub WzCampaignTitleWidget_geometryChanged(ByVal self As WIDGET)
    Dim i As Long
    For i = 0 To self.childCount() - 1
        self.childAt(i).callCalcLayout
    Next
End Sub

Public Sub WzCampaignTitleWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzCampaignTitleWidget, x0 As Long, y0 As Long
    Set d = WzCampaignTitleWidget_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    ' Draw the background
    If Not d.pBackgroundTexture Is Nothing Then
        iV_DrawImageAnisotropic d.pBackgroundTexture, x0, y0, 0!, 0!, self.width(), self.height(), 0!, pal_RGBA(255, 255, 255, 255)
    Else
        ' Draw a static background
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), psPalette(WZCOL_TRANSPARENT_BOX)
        pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(0, 0, 0, 80)
    End If
End Sub

' MARK: - WzCampaignSubInfoBox

Public Sub WzCampaignSubInfoBox_construct(ByVal self As WIDGET)
    Dim d As WzCampaignSubInfoBox
    WIDGET_construct self
    self.kind = K_WzCampaignSubInfoBox ' (VB) virtual calls from here on reach WzCampaignSubInfoBox's overrides
    Set d = WzCampaignSubInfoBox_(self)
    d.leftPadding = 40
    d.topPadding = 20
    d.imgDimensions = 16
    d.imgTextPadding = 5
    d.universeTextureColor = pal_RGBA(255, 255, 255, 255)
    d.compatibilityResult = WzModCompatibilityResult_OK
End Sub

Private Sub WzCampaignSubInfoBox_initialize(ByVal self As WIDGET, ByVal modInfo As WzCampaignModInfo)
    Dim d As WzCampaignSubInfoBox, universe As WzCampaignUniverse, universeDisplayStr As String, compatibilityStr As String
    Dim minVerString As String, verCompatRangeStr As String
    Set d = WzCampaignSubInfoBox_(self)
    If Not modInfo Is Nothing Then
        ' load details from mod
        universe = modInfo.universe
    Else
        ' load built-in details for original campaign
        universe = WzCampaignUniverse_WZ2100
    End If

    ' Create the "universe" label
    Select Case universe
    Case WzCampaignUniverse_WZ2100
        universeDisplayStr = "Warzone 2100 Universe"
        Set d.pUniverseTexture = gfx_api_context_loadTextureFromFile("images/warzone2100.png", texture_type_game_texture, -1, -1, True)
        d.universeTextureColor = pal_RGBA(255, 255, 255, 255)
    Case WzCampaignUniverse_WZ2100Extended
        universeDisplayStr = "WZ2100 Extended Universe"
        Set d.pUniverseTexture = gfx_api_context_loadTextureFromFile("images/warzone2100.png", texture_type_game_texture, -1, -1, True)
        d.universeTextureColor = pal_RGBA(255, 99, 0, 255)
    Case WzCampaignUniverse_Unique
        universeDisplayStr = "Unique Universe"
    End Select

    Set d.universeLabel = W_LABEL_new()
    W_LABEL_setFont d.universeLabel, font_small, psPalette(WZCOL_TEXT_BRIGHT)
    W_LABEL_setCanTruncate d.universeLabel, True
    d.universeLabel.setString universeDisplayStr
    d.universeLabel.setGeometry d.leftPadding, d.topPadding, W_LABEL_getMaxLineWidth(d.universeLabel), iV_GetTextLineSize(font_small)
    self.attach d.universeLabel
    d.universeLabel.setCalcLayout WzFunc_new(FN_subInfoBox_universeCalcLayout)

    If Not modInfo Is Nothing Then
        ' Also display compatibility info
        d.compatibilityResult = modInfo.getModCompatibility()
        Set d.compatibilityLabel = W_LABEL_new()
        W_LABEL_setFont d.compatibilityLabel, font_small, psPalette(WZCOL_TEXT_BRIGHT)
        W_LABEL_setCanTruncate d.compatibilityLabel, True
        Select Case d.compatibilityResult
        Case WzModCompatibilityResult_NOT_COMPATIBLE
            minVerString = ChrW$(&H2265) & " "  ' "U+2265 "
            minVerString = minVerString & modInfo.minVersionSupported
            compatibilityStr = "Requires " & minVerString
        Case WzModCompatibilityResult_OK
            compatibilityStr = "Compatible"
        Case WzModCompatibilityResult_MAYBE_COMPATIBLE
            If Len(modInfo.maxVersionTested) <> 0 Then
                verCompatRangeStr = modInfo.minVersionSupported & " - " & modInfo.maxVersionTested
            Else
                verCompatRangeStr = ChrW$(&H2265) & " " & modInfo.minVersionSupported
            End If
            compatibilityStr = "Compatible " & verCompatRangeStr
        End Select
        d.compatibilityLabel.setString compatibilityStr
        d.compatibilityLabel.setGeometry d.leftPadding, d.topPadding, W_LABEL_getMaxLineWidth(d.compatibilityLabel), iV_GetTextLineSize(font_small)
        self.attach d.compatibilityLabel
        d.compatibilityLabel.setCalcLayout WzFunc_new(FN_subInfoBox_compatibilityCalcLayout)
    End If
End Sub

Public Function WzCampaignSubInfoBox_make(ByVal modInfo As WzCampaignModInfo) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignSubInfoBox)
    WzCampaignSubInfoBox_construct widget
    WzCampaignSubInfoBox_initialize widget, modInfo
    Set WzCampaignSubInfoBox_make = widget
End Function

' ~WzCampaignSubInfoBox()
Public Sub WzCampaignSubInfoBox_destruct(ByVal self As WIDGET)
    Dim d As WzCampaignSubInfoBox
    Set d = WzCampaignSubInfoBox_(self)
    If Not d.pUniverseTexture Is Nothing Then Set d.pUniverseTexture = Nothing
End Sub

Public Sub WzCampaignSubInfoBox_geometryChanged(ByVal self As WIDGET)
    Dim d As WzCampaignSubInfoBox, i As Long
    Set d = WzCampaignSubInfoBox_(self)
    d.imgDimensions = maxL(minL(minL(self.width(), self.height() - 2), 14), 0)
    For i = 0 To self.childCount() - 1
        self.childAt(i).callCalcLayout
    Next
End Sub

Public Sub WzCampaignSubInfoBox_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WzCampaignSubInfoBox, x0 As Long, y0 As Long, imgPosY0 As Long, imgPosX0 As Long, compatImgPosX0 As Long
    Dim img As Long, imgColor As Long
    Set d = WzCampaignSubInfoBox_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), pal_RGBA(0, 0, 0, 128)

    imgPosY0 = y0 + (self.height() - d.imgDimensions) \ 2

    ' draw universe image
    imgPosX0 = x0 + (d.universeLabel.x() - (d.imgDimensions + d.imgTextPadding))
    If Not d.pUniverseTexture Is Nothing Then
        iV_DrawImageAnisotropic d.pUniverseTexture, imgPosX0, imgPosY0, 0!, 0!, d.imgDimensions + 1, d.imgDimensions + 1, 0!, d.universeTextureColor
    Else
        ' draw "other" universe image
        iV_DrawImageFileAnisotropicTint FrontImages, IMAGE_CAM_UNIVERSE_STARS, imgPosX0, imgPosY0, d.imgDimensions, d.imgDimensions, pal_RGBA(255, 255, 255, 200)
    End If

    If Not d.compatibilityLabel Is Nothing Then
        ' draw compatibility image
        compatImgPosX0 = x0 + (d.compatibilityLabel.x() - (d.imgDimensions + d.imgTextPadding))
        img = 0
        Select Case d.compatibilityResult
        Case WzModCompatibilityResult_NOT_COMPATIBLE
            img = IMAGE_BAN_CIRCLE
            imgColor = pal_RGBA(191, 2, 2, 255)
        Case WzModCompatibilityResult_OK
            img = IMAGE_CHECK_CIRCLE
            imgColor = pal_RGBA(0, 168, 10, 255)
        Case WzModCompatibilityResult_MAYBE_COMPATIBLE
            img = IMAGE_QUESTION_CIRCLE
            imgColor = pal_RGBA(150, 150, 150, 255)
        End Select
        iV_DrawImageFileAnisotropicTint FrontImages, img, compatImgPosX0, imgPosY0, d.imgDimensions, d.imgDimensions, imgColor
    End If
End Sub

' MARK: - CampaignStartOptionsForm

' (modInfo: Nothing = nullopt; the result: a Collection of WzCampaignTweakOptionSetting)
Private Function buildTweakOptionSettings(ByVal modInfo As WzCampaignModInfo) As Collection
    Dim results As Collection, k As Long, it As WzCampaignTweakOptionSetting, uniqueIdentifier As String
    Dim modCamIt As WzCampaignTweakOption, opt As WzCampaignTweakOption, customOpt As WzCampaignTweakOption
    Dim localizedDisplayName As String, localizedDescription As String
    Set results = New Collection

    ' built-in, camTweakOptions

    results.Add WzCampaignTweakOptionSetting_new( _
        "timerPowerBonus", _
        "Timer Power Bonus", _
        "Grant power based on the remaining level timer, so there's no need to wait out the clock.", _
        True, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "classicTimers", _
        "Classic Timers", _
        "Alter / disable timers on specific levels, like the original release. (Allows power-cheating.)", _
        False, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "playerUnitCap40", _
        "40 Unit Limit", _
        "Lower the player's unit limit to 40, matching the original PS1 release.", _
        False, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "autosavesOnly", _
        "Autosaves-Only", _
        "Disable the ability to manually save / quick-save, limit autosaves to the beginning of each level.", _
        False, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "ps1Modifiers", _
        "PS1 Modifiers", _
        "Reduces the damage the enemy deals to a third of the current difficulty modifier.", _
        False, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "infiniteTime", _
        "Infinite Time", _
        "Missions have no timer where possible (Timer Power Bonus replaced with a small power reward if enabled).", _
        False, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "victoryHints", _
        "Victory Hints", _
        "Displays a console message every few minutes showing victory related information.", _
        True, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "gammaEndBonus", _
        "Finale Fun", _
        "Activate Final Gamma mission bonus content.", _
        True, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "gammaBonusLevel", _
        "Gamma Bonus", _
        "Extra Gamma mission making use of a map created by the original developement team (map updated by: DARwins).", _
        False, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "insanePlus", _
        "Additional Spawns (Insane difficulty)", _
        "Enables additional enemy spawns and behavior for Insane difficulty (or higher).", _
        True, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "insanePlusLowDiff", _
        "Additional Spawns (Lower difficulties than Insane)", _
        "Enables additional enemy spawns and behavior for Hard difficulty or lower.", _
        False, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "fastExp", _
        "Fast EXP gain", _
        "Increases unit experience point gain by 2x.", _
        False, True _
    )

    results.Add WzCampaignTweakOptionSetting_new( _
        "towerWars", _
        "Tower Wars", _
        "Player gets significantly stronger structures.", _
        False, True _
    )

    If Not modInfo Is Nothing Then
        k = 1
        Do While k <= results.Count
            Set it = results(k)
            uniqueIdentifier = it.uniqueIdentifier
            ' (std::find_if(modCamTweakOptions..., [uniqueIdentifier](const WzCampaignTweakOption& opt) {...}))
            Set modCamIt = Nothing
            For Each opt In modInfo.camTweakOptions
                If opt.uniqueIdentifier = uniqueIdentifier Then
                    Set modCamIt = opt
                    Exit For
                End If
            Next
            If modCamIt Is Nothing Then
                ' builtin option not found in mod's camTweakOptions - remove it from list
                results.Remove k
            Else
                it.defaultValue = modCamIt.enabled
                it.currentValue = modCamIt.enabled
                it.editable = modCamIt.userEditable
                k = k + 1
            End If
        Loop

        ' append any customTweakOptions
        For Each customOpt In modInfo.customTweakOptions
            localizedDisplayName = customOpt.displayName    ' displayName.getLocalizedString()
            localizedDescription = customOpt.description    ' description.getLocalizedString()
            If Len(localizedDisplayName) = 0 Then
                GoTo continue_
            End If
            If customOpt.type_ <> WzCampaignTweakOption_Type_Bool Then
                ' this doesn't currently support anything else
                GoTo continue_
            End If
            results.Add WzCampaignTweakOptionSetting_new( _
                customOpt.uniqueIdentifier, _
                localizedDisplayName, _
                localizedDescription, _
                customOpt.enabled, customOpt.userEditable _
            )
continue_:
        Next
    End If

    Set buildTweakOptionSettings = results
End Function

' struct OriginalCampaignBalanceChoice (OriginalCampaignBalanceChoice.cls) {displayName, description, modInfo}
Private Function OriginalCampaignBalanceChoice_new(ByVal displayName As String, ByVal description As String, ByVal modInfo As WzCampaignModInfo) As OriginalCampaignBalanceChoice
    Dim c As OriginalCampaignBalanceChoice
    Set c = New OriginalCampaignBalanceChoice
    c.displayName = displayName
    c.description = description
    Set c.modInfo = modInfo
    Set OriginalCampaignBalanceChoice_new = c
End Function

Public Sub CampaignStartOptionsForm_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_CampaignStartOptionsForm ' (VB) virtual calls from here on reach CampaignStartOptionsForm's overrides
End Sub

' (parent: the WzCampaignSelectorTitleUI's address; modInfo: Nothing = nullopt)
Public Function CampaignStartOptionsForm_make(ByVal parent As Long, ByVal modInfo As WzCampaignModInfo) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_CampaignStartOptionsForm)
    CampaignStartOptionsForm_construct widget
    Set CampaignStartOptionsForm_(widget).modInfo = modInfo
    CampaignStartOptionsForm_initialize widget, parent
    Set CampaignStartOptionsForm_make = widget
End Function

Private Function CampaignStartOptionsForm_getCurrentBalanceModInfo(ByVal self As WIDGET) As WzCampaignModInfo
    Dim d As CampaignStartOptionsForm, idx As Long, balanceChoice As OriginalCampaignBalanceChoice
    Set d = CampaignStartOptionsForm_(self)
    If Not d.campaignBalanceDropdown Is Nothing And d.campaignBalanceChoices.Count > 0 Then
        ' If an alternative balance is selected, get the balance mod
        If Not DropdownWidget_getSelectedIndex(d.campaignBalanceDropdown, idx) Then idx = 0   ' value_or(0)
        Set balanceChoice = d.campaignBalanceChoices(idx + 1)
        Set CampaignStartOptionsForm_getCurrentBalanceModInfo = balanceChoice.modInfo
        Exit Function
    End If
    Set CampaignStartOptionsForm_getCurrentBalanceModInfo = Nothing
End Function

Private Sub CampaignStartOptionsForm_handleStartGame(ByVal self As WIDGET)
    Dim d As CampaignStartOptionsForm, idx As Long, balanceChoice As OriginalCampaignBalanceChoice, campaignFile As CAMPAIGN_FILE
    Dim tweakOptionsJSONMap As nlohmann_json, k As Long, o As WzCampaignTweakOptionSetting
    Set d = CampaignStartOptionsForm_(self)
    If Not (Not d.campaignSelectorDropdown Is Nothing) Then ASSERT_FAILURE "campaignSelectorDropdown != nullptr", "src/titleui/campaign.cpp:1830", "handleStartGame", "No campaign selector dropdown?": Exit Sub

    StringVector_clear campaign_mods
    If Not d.modInfo Is Nothing Then
        ' Set campaign_mods
        StringVector_push_back campaign_mods, d.modInfo.modFilename
        ' Clear existing level list
        levShutDown
        levInitialise
        rebuildSearchPath mod_campaign, True
        pal_Init ' Update palettes.
        ' must reload campaign levels, as the campaign mod may provide its own levels
        If Not buildMapList(True) Then
            debug_ 1845, LOG_FATAL, "handleStartGame", "Failed to build map list"
            Exit Sub
        End If
    End If

    If Not d.campaignBalanceDropdown Is Nothing And d.campaignBalanceChoices.Count > 0 Then
        ' If an alternative balance is selected, load the balance mod
        If Not DropdownWidget_getSelectedIndex(d.campaignBalanceDropdown, idx) Then idx = 0   ' value_or(0)
        Set balanceChoice = d.campaignBalanceChoices(idx + 1)
        If Not balanceChoice.modInfo Is Nothing Then
            ' Load campaign balance mod
            StringVector_push_back campaign_mods, balanceChoice.modInfo.modFilename
            ' NOTE: Campaign balance mods shouldn't provide alternate levels, so there shouldn't be a need to reload levels
        End If
    End If

    SPinit LEVEL_TYPE_CAMPAIGN

    If Not DropdownWidget_getSelectedIndex(d.campaignSelectorDropdown, idx) Then idx = 0   ' value_or(0)
    campaignFile = d.campaignList(idx)

    aLevelName = Left$(campaignFile.level, MAX_LEVEL_NAME_SIZE)    ' sstrcpy(aLevelName, ...)
    setCampaignName campaignFile.name_

    ' show video
    If Len(campaignFile.video) <> 0 Then
        seq_ClearSeqList
        seq_AddSeqToList campaignFile.video, vbNullString, campaignFile.captions, False
        seq_StartNextFullScreenVideo
    End If

    ' set tweak options
    ' (std::unordered_map<std::string, nlohmann::json> tweakOptionsJSONMap: a json object)
    Set tweakOptionsJSONMap = json_object()
    For k = 1 To d.tweakOptionSettings.Count
        Set o = d.tweakOptionSettings(k)
        tweakOptionsJSONMap.set_item o.uniqueIdentifier, json_value(o.currentValue)
    Next
    setCamTweakOptions tweakOptionsJSONMap

    debug_ 1885, LOG_WZ, "handleStartGame", "Loading campaign level -- %s", aLevelName
    changeTitleMode STARTGAME
End Sub

Public Sub CampaignStartOptionsForm_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, leftLineX0 As Long, y1 As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    pie_UniTransBoxFill x0, y0, x0 + self.width(), y0 + self.height(), psPalette(WZCOL_TRANSPARENT_BOX)

    leftLineX0 = x0
    y1 = y0 + self.height()
    iV_Line leftLineX0, y0, leftLineX0, y1, pal_RGBA(0, 0, 0, 255)
End Sub

' const char* difficultyLevelToString(DIFFICULTY_LEVEL difficulty)
Public Function difficultyLevelToString(ByVal difficulty As Long) As String
    Select Case difficulty
    Case DL_SUPER_EASY: difficultyLevelToString = "Super Easy": Exit Function
    Case DL_EASY: difficultyLevelToString = "Easy": Exit Function
    Case DL_NORMAL: difficultyLevelToString = "Normal": Exit Function
    Case DL_HARD: difficultyLevelToString = "Hard": Exit Function
    Case DL_INSANE: difficultyLevelToString = "Insane": Exit Function
    End Select
    difficultyLevelToString = "Unsupported"
End Function

' const char* getCampaignDifficultyDescriptionString(DIFFICULTY_LEVEL difficulty)
Public Function getCampaignDifficultyDescriptionString(ByVal difficulty As Long) As String
    Select Case difficulty
    Case DL_SUPER_EASY: getCampaignDifficultyDescriptionString = "A more relaxed playthrough - easier than easy": Exit Function
    Case DL_EASY: getCampaignDifficultyDescriptionString = "A slightly easier challenge than Normal": Exit Function
    Case DL_NORMAL: getCampaignDifficultyDescriptionString = "A fun & challenging experience": Exit Function
    Case DL_HARD: getCampaignDifficultyDescriptionString = "A challenge for players who have beaten the game": Exit Function
    Case DL_INSANE: getCampaignDifficultyDescriptionString = "An unforgiving challenge for experienced players only": Exit Function
    End Select
    getCampaignDifficultyDescriptionString = "Unsupported"
End Function

' std::shared_ptr<W_BUTTON> makeCampaignUITextButton(UDWORD id, const std::string &txt, unsigned int style, optional<unsigned int> minimumWidth = nullopt)
Public Function makeCampaignUITextButton(ByVal id As Long, ByVal txt As String, ByVal style As Long, Optional ByVal has_minimumWidth As Boolean = False, Optional ByVal minimumWidth As Long = 0) As WIDGET
    Dim sButInit As W_BUTINIT, textWidth As Long, minW As Long, pUserData As DisplayTextOptionCache, button As WIDGET
    W_BUTINIT_construct sButInit

    ' Align
    sButInit.FontID = font_regular_bold
    If (style And WBUT_TXTCENTRE) = 0 Then
        textWidth = iV_GetTextWidth(txt, sButInit.FontID)
        If has_minimumWidth Then minW = minimumWidth Else minW = 0
        If ULT(minW, textWidth) Then sButInit.width = ToInt16(textWidth) Else sButInit.width = ToInt16(minW)
    Else
        sButInit.style = sButInit.style Or WBUT_TXTCENTRE
    End If

    ' Enable right clicks
    If (style And WBUT_SECONDARY) <> 0 Then
        sButInit.style = sButInit.style Or WBUT_SECONDARY
    End If

    sButInit.UserData = (style And WBUT_DISABLE) ' store disable state
    Set pUserData = New DisplayTextOptionCache
    pUserData.overflowBehavior_ = OverflowBehavior_ShrinkFont
    Set sButInit.pUserData = pUserData
    Set sButInit.onDelete = WzFunc_new(FN_deleteDisplayTextOptionCache)

    sButInit.height = FRONTEND_CAMPAIGNUI_BUTHEIGHT
    Set sButInit.pDisplay = WzFunc_new(FN_displayTextOption)
    sButInit.pText = txt

    Set button = W_BUTTON_new_init(sButInit)
    ' Disable button
    If (style And WBUT_DISABLE) <> 0 Then
        button.setState WBUT_DISABLE
    End If

    Set makeCampaignUITextButton = button
End Function

' (availableDifficultyLevels(0 .. nLevels - 1))
Private Function CampaignStartOptionsForm_makeDifficultySelectorDropdown(ByVal self As WIDGET, availableDifficultyLevels() As Long, ByVal nLevels As Long) As WIDGET
    Dim choiceLevels As Variant, i As Long, currentSettingIdx As Long, currValue As Long, found_ As Long
    Dim dropdown As WIDGET, maxItemHeight As Long, item As WIDGET, lvl As Long, levels() As Long
    ' (std::vector<std::tuple<WzString, DIFFICULTY_LEVEL, WzString>> dropDownChoices: the levels, in choiceLevels;
    ' the strings are difficultyLevelToString and getCampaignDifficultyDescriptionString of them)
    If nLevels > 0 Then
        ReDim levels(0 To nLevels - 1)
        For i = 0 To nLevels - 1
            levels(i) = availableDifficultyLevels(i)
        Next
        choiceLevels = levels
    Else
        choiceLevels = Empty
    End If

    ' If current value (from config) is not one of the presets in dropDownChoices, add a "Custom" entry
    currentSettingIdx = 0
    currValue = getDifficultyLevel()
    found_ = -1
    For i = 0 To nLevels - 1
        If choiceLevels(i) = currValue Then
            found_ = i
            Exit For
        End If
    Next
    If found_ >= 0 Then
        currentSettingIdx = found_
    Else
        ' Future TODO: If the current difficulty setting doesn't exist in availableDifficultyLevels, find the nearest available option
        currentSettingIdx = 0
    End If

    Set dropdown = DropdownWidget_new()

    maxItemHeight = 0
    For i = 0 To nLevels - 1
        lvl = choiceLevels(i)
        Set item = WzCampaignOptionButton_make()
        item.setString difficultyLevelToString(lvl)
        WzCampaignOptionButton_setDescriptionText item, getCampaignDifficultyDescriptionString(lvl)
        item.setGeometry 0, 0, item.idealWidth(), item.idealHeight()
        maxItemHeight = maxL(maxItemHeight, item.idealHeight())
        DropdownWidget_addItem dropdown, item
    Next

    DropdownWidget_setListHeight dropdown, maxItemHeight * minL(4, nLevels)

    DropdownWidget_setSelectedIndex dropdown, currentSettingIdx

    DropdownWidget_setCanChange dropdown, WzFunc_new(FN_difficultyDropdown_canChange, choiceLevels)

    Set CampaignStartOptionsForm_makeDifficultySelectorDropdown = Margin_wrap(0, 10, 0, 10, dropdown)   ' Margin(0, 10).wrap(dropdown)
End Function

' (availableCampaignMods: a Collection of WzCampaignModInfo)
Public Function CampaignStartOptionsForm_updateCampaignBalanceDropdown(ByVal self As WIDGET, ByVal availableCampaignMods As Collection) As Boolean
    Const classicModIdx As Long = 1
    Const pumpkinModIdx As Long = 2
    Dim d As CampaignStartOptionsForm, modItem As WzCampaignModInfo, maxItemHeight As Long, i As Long
    Dim opt As OriginalCampaignBalanceChoice, item As WIDGET
    Set d = CampaignStartOptionsForm_(self)
    If Not (Not d.campaignBalanceDropdown Is Nothing) Then ASSERT_FAILURE "campaignBalanceDropdown != nullptr", "src/titleui/campaign.cpp:2026", "updateCampaignBalanceDropdown", "No campaign balance dropdown?": CampaignStartOptionsForm_updateCampaignBalanceDropdown = False: Exit Function

    If d.campaignBalanceChoices.Count > 1 Then
        ' erase all but first element, and repopulate
        Do While d.campaignBalanceChoices.Count > 1
            d.campaignBalanceChoices.Remove 2
        Loop
    End If

    d.campaignBalanceChoices.Add OriginalCampaignBalanceChoice_new("Classic", """Classic"" campaign balance", Nothing)
    d.campaignBalanceChoices.Add OriginalCampaignBalanceChoice_new("Pumpkin", """Pumpkin"" campaign balance", Nothing)

    For Each modItem In availableCampaignMods
        If modItem.type_ <> WzModType_CampaignBalanceMod Then
            GoTo continue_
        End If
        If modItem.modFilename = WZ2100_CORE_CLASSIC_BALANCE_MOD_FILENAME Then
            ' already added above in all cases - just update the modinfo
            Set d.campaignBalanceChoices(classicModIdx + 1).modInfo = modItem
            GoTo continue_
        End If
        If modItem.modFilename = WZ2100_CORE_PUMPKIN_BALANCE_MOD_FILENAME Then
            ' already added above in all cases - just update the modinfo
            Set d.campaignBalanceChoices(pumpkinModIdx + 1).modInfo = modItem
            GoTo continue_
        End If
        d.campaignBalanceChoices.Add OriginalCampaignBalanceChoice_new(modItem.name_, modItem.description, modItem)
continue_:
    Next

    DropdownWidget_clear d.campaignBalanceDropdown

    maxItemHeight = 0
    For i = 0 To d.campaignBalanceChoices.Count - 1
        Set opt = d.campaignBalanceChoices(i + 1)
        Set item = WzCampaignOptionButton_make()
        item.setString opt.displayName
        WzCampaignOptionButton_setDescriptionText item, opt.description
        item.setGeometry 0, 0, item.idealWidth(), item.idealHeight()
        maxItemHeight = maxL(maxItemHeight, item.idealHeight())

        If ((i = classicModIdx) Or (i = pumpkinModIdx)) And (opt.modInfo Is Nothing) Then
            ' disable classic balance mod, which is always displayed, if unavailable
            item.setState WBUT_DISABLE
        End If

        If Not opt.modInfo Is Nothing Then
            If opt.modInfo.getModCompatibility() = WzModCompatibilityResult_NOT_COMPATIBLE Then
                ' disable incompatible mods
                item.setState WBUT_DISABLE
            End If
        End If

        DropdownWidget_addItem d.campaignBalanceDropdown, item
    Next

    DropdownWidget_setListHeight d.campaignBalanceDropdown, maxItemHeight * minL(4, d.campaignBalanceChoices.Count)

    DropdownWidget_setSelectedIndex d.campaignBalanceDropdown, 0

    CampaignStartOptionsForm_updateCampaignBalanceDropdown = True
End Function

' WzCampaignHelpPanel

Public Sub WzCampaignHelpPanel_construct(ByVal self As WIDGET)
    Dim d As WzCampaignHelpPanel
    WIDGET_construct self
    self.kind = K_WzCampaignHelpPanel ' (VB) virtual calls from here on reach WzCampaignHelpPanel's overrides
    Set d = WzCampaignHelpPanel_(self)
    d.baseFontId = font_regular
    d.innerPadding = 20
    d.betweenButtonPadding = 15
End Sub

' (choiceHelpDescriptions(0 .. nChoices - 1); closeHandlerFunc: Nothing = empty)
Public Function WzCampaignHelpPanel_make(ByVal optionInfo As OptionInfo, choiceHelpDescriptions() As OptionChoiceHelpDescription, ByVal nChoices As Long, ByVal idealWidth As Long, ByVal closeHandlerFunc As WzFunc) As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzCampaignHelpPanel)
    WzCampaignHelpPanel_construct widget
    WzCampaignHelpPanel_initialize widget, optionInfo, choiceHelpDescriptions, nChoices, idealWidth, closeHandlerFunc
    Set WzCampaignHelpPanel_make = widget
End Function

Private Function WzCampaignHelpPanel_buildParagraph(ByVal self As WIDGET, ByVal optionInfo As OptionInfo, choiceHelpDescriptions() As OptionChoiceHelpDescription, ByVal nChoices As Long, ByVal idealWidth As Long) As WIDGET
    Dim d As WzCampaignHelpPanel, baseFont As iV_fonts, baseFontBold As iV_fonts, paragraph As WIDGET, wroteALine As Boolean
    Dim optionHelpDescription As String, k As Long
    Set d = WzCampaignHelpPanel_(self)
    If Not iV_FontModifyBold(d.baseFontId, False, baseFont) Then baseFont = d.baseFontId       ' value_or(baseFontId)
    If Not iV_FontModifyBold(d.baseFontId, True, baseFontBold) Then baseFontBold = d.baseFontId

    ' Build description paragraph
    Set paragraph = Paragraph_new()
    wroteALine = False
    Paragraph_setLineSpacing paragraph, 1
    paragraph.setGeometry 0, 0, idealWidth, 20

    ' Add general option info text description
    Paragraph_setFont paragraph, baseFont
    Paragraph_setFontColour paragraph, psPalette(WZCOL_TEXT_BRIGHT)
    optionHelpDescription = optionInfo.getTranslatedHelpDescription()
    If Len(optionHelpDescription) <> 0 Then
        Paragraph_addText paragraph, optionHelpDescription
        wroteALine = True
    End If

    For k = 0 To nChoices - 1
        If Len(choiceHelpDescriptions(k).helpDescription) = 0 Then
            GoTo continue_
        End If

        If wroteALine Then
            Paragraph_addText paragraph, vbLf & " " & vbLf
        End If

        ' Add a bold string with the displayName and then append the choice help description
        If Len(choiceHelpDescriptions(k).displayName) <> 0 Then
            Paragraph_setFont paragraph, baseFontBold
            Paragraph_addText paragraph, choiceHelpDescriptions(k).displayName & ": "
        End If
        Paragraph_setFont paragraph, baseFont
        Paragraph_addText paragraph, choiceHelpDescriptions(k).helpDescription

        wroteALine = True
continue_:
    Next

    If Not wroteALine Then
        Set WzCampaignHelpPanel_buildParagraph = Nothing
        Exit Function
    End If

    Set WzCampaignHelpPanel_buildParagraph = paragraph
End Function

Private Sub WzCampaignHelpPanel_initialize(ByVal self As WIDGET, ByVal optionInfo As OptionInfo, choiceHelpDescriptions() As OptionChoiceHelpDescription, ByVal nChoices As Long, ByVal idealWidth As Long, ByVal closeHandlerFunc As WzFunc)
    Dim d As WzCampaignHelpPanel, listY0 As Long, paragraph As WIDGET
    Set d = WzCampaignHelpPanel_(self)
    Set d.formTitle = W_LABEL_new()
    W_LABEL_setFont d.formTitle, font_regular_bold, psPalette(WZCOL_TEXT_MEDIUM)
    d.formTitle.setString "About Balance Options"
    W_LABEL_setCanTruncate d.formTitle, True
    d.formTitle.setTransparentToMouse True
    d.formTitle.setGeometry d.innerPadding, d.innerPadding, W_LABEL_getMaxLineWidth(d.formTitle), iV_GetTextLineSize(font_regular_bold)
    self.attach d.formTitle

    ' (auto weakEditForm = std::weak_ptr<WzCampaignTweakOptionsEditForm>(std::dynamic_pointer_cast<...>(shared_from_this())):
    ' unused, and empty - this is no WzCampaignTweakOptionsEditForm)

    If Not closeHandlerFunc Is Nothing Then
        ' Add "Close" button
        Set d.closeButton = WzFrontendImageButton_make(False, 0)
        d.closeButton.setString "Close"
        W_BUTTON_addOnClickHandler d.closeButton, WzFunc_new(FN_campaign_closeButtonClick, closeHandlerFunc)
        self.attach d.closeButton
    End If

    Set d.paragraphContainer = ScrollableListWidget_make()
    ScrollableListWidget_setItemSpacing d.paragraphContainer, 10
    ScrollableListWidget_setExpandWhenScrollbarInvisible d.paragraphContainer, False
    listY0 = d.formTitle.y() + d.formTitle.height() + d.innerPadding
    d.paragraphContainer.setGeometry d.innerPadding, listY0, idealWidth - (d.innerPadding * 2), 100
    self.attach d.paragraphContainer

    Set paragraph = WzCampaignHelpPanel_buildParagraph(self, optionInfo, choiceHelpDescriptions, nChoices, idealWidth - (d.innerPadding * 2))
    If Not paragraph Is Nothing Then
        paragraph.setGeometry 0, 0, d.paragraphContainer.width(), 20
        ScrollableListWidget_addItem d.paragraphContainer, paragraph
    End If
End Sub

Public Sub WzCampaignHelpPanel_geometryChanged(ByVal self As WIDGET)
    Dim d As WzCampaignHelpPanel, listY0 As Long, listH As Long, listW As Long, listIdealHeight As Long
    Dim lastButtonX0 As Long, buttonY0 As Long, topButtonHeight As Long, buttonX0 As Long
    Set d = WzCampaignHelpPanel_(self)
    If self.width() <= 1 Or self.height() <= 1 Then
        Exit Sub
    End If

    ' recalculate position of paragraph container
    listY0 = d.paragraphContainer.y()
    listH = self.height() - listY0 - d.innerPadding
    listW = self.width() - (d.innerPadding * 2)
    If listW > 1 And listH > 1 Then
        d.paragraphContainer.setGeometry d.innerPadding, listY0, listW, listH

        ' truncate list height at needed height (if scrollable height is less)
        listIdealHeight = d.paragraphContainer.idealHeight()
        If listIdealHeight < listH Then
            d.paragraphContainer.setGeometry d.innerPadding, listY0, listW, listIdealHeight
        End If
    End If

    ' position top button(s)
    lastButtonX0 = self.width() - d.innerPadding
    buttonY0 = d.innerPadding
    topButtonHeight = d.closeButton.idealHeight()
    If Not d.closeButton Is Nothing Then
        buttonX0 = lastButtonX0 - d.closeButton.idealWidth()
        d.closeButton.setGeometry buttonX0, buttonY0, d.closeButton.idealWidth(), topButtonHeight
        lastButtonX0 = buttonX0 - d.betweenButtonPadding
    End If
End Sub

' ideal height *for the current width*
Public Function WzCampaignHelpPanel_idealHeight(ByVal self As WIDGET) As Long
    Dim d As WzCampaignHelpPanel
    Set d = WzCampaignHelpPanel_(self)
    WzCampaignHelpPanel_idealHeight = d.paragraphContainer.y() + d.paragraphContainer.height() + d.innerPadding
End Function

Public Sub WzCampaignHelpPanel_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long, backColor As Long
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    backColor = psPalette(WZCOL_TRANSPARENT_BOX)
    backColor = pielightWithAlpha(backColor, maxL(220, pielightAlpha(backColor)))
    pie_UniTransBoxFill x0, y0, x1, y1, backColor

    iV_Box x0, y0, x1, y1, pal_RGBA(255, 255, 255, 175)
End Sub

Private Sub CampaignStartOptionsForm_openCampaignBalanceInfoOverlay(ByVal self As WIDGET)
    Const remasteredBalanceIdx As Long = 0
    Const classicModIdx As Long = 1
    Const pumpkinModIdx As Long = 2
    Dim d As CampaignStartOptionsForm, optionInfo As OptionInfo, choiceHelpDescriptions() As OptionChoiceHelpDescription
    Dim nChoices As Long, i As Long, opt As OriginalCampaignBalanceChoice, choiceHelpDescription As OptionChoiceHelpDescription
    Dim blankChoiceHelpDescription As OptionChoiceHelpDescription, weakSelf As Long, widgetWidth As Long, helpWidget As WIDGET
    Set d = CampaignStartOptionsForm_(self)

    Set optionInfo = OptionInfo_new("campaign.balancemod", "Balance", "Warzone 2100 offers multiple balance options for the Original Campaign, reflecting different experiences over its long history.")
    nChoices = 0
    For i = 0 To d.campaignBalanceChoices.Count - 1
        Set opt = d.campaignBalanceChoices(i + 1)

        choiceHelpDescription = blankChoiceHelpDescription  ' OptionChoiceHelpDescription choiceHelpDescription;
        choiceHelpDescription.displayName = opt.displayName

        If i = remasteredBalanceIdx Then
            choiceHelpDescription.helpDescription = "The remastered campaign experience"
            choiceHelpDescription.helpDescription = choiceHelpDescription.helpDescription & "." & vbLf
            choiceHelpDescription.helpDescription = choiceHelpDescription.helpDescription & "Increases the viability of weaponry and designs on a per mission basis."
        ElseIf i = classicModIdx Then
            choiceHelpDescription.helpDescription = "Stats of game versions ~2.x - 3.x."
            choiceHelpDescription.helpDescription = choiceHelpDescription.helpDescription & vbLf
            choiceHelpDescription.helpDescription = choiceHelpDescription.helpDescription & "Unintentionally affected by game patches over the years."
        ElseIf i = pumpkinModIdx Then
            choiceHelpDescription.helpDescription = "Matching game version 1.10, this is the closest representation of stats that the original developers (Pumpkin Studios) intended."
        ElseIf Not opt.modInfo Is Nothing Then
            choiceHelpDescription.helpDescription = opt.modInfo.description   ' description.getLocalizedString()
        End If

        ReDim Preserve choiceHelpDescriptions(0 To nChoices)
        choiceHelpDescriptions(nChoices) = choiceHelpDescription
        nChoices = nChoices + 1
    Next

    weakSelf = ObjPtr(self)

    widgetWidth = minL(self.width() - 20, 450)
    Set helpWidget = WzCampaignHelpPanel_make(optionInfo, choiceHelpDescriptions, nChoices, widgetWidth, WzFunc_new(FN_balanceInfo_closeOverlay, weakSelf))
    If helpWidget Is Nothing Then
        ' nothing to display
        Exit Sub
    End If
    helpWidget.setGeometry 0, 0, widgetWidth, 370 ' set the width, and a fixed maximum height
    helpWidget.setGeometry 0, 0, widgetWidth, minL(helpWidget.idealHeight(), 370) ' shrink to idealHeight if needed for current width

    CampaignStartOptionsForm_openOverlay self, helpWidget, d.campaignBalanceDropdown, WzFunc_new(FN_balanceInfo_onClose, weakSelf)
End Sub

Private Function CampaignStartOptionsForm_makeCampaignBalanceDropdown(ByVal self As WIDGET) As WIDGET
    Const classicModIdx As Long = 1
    Const pumpkinModIdx As Long = 2
    Dim d As CampaignStartOptionsForm, maxItemHeight As Long, i As Long, opt As OriginalCampaignBalanceChoice, item As WIDGET
    Dim weakSelf As Long, infoWrappedCampaignBalanceDropdown As WIDGET
    Set d = CampaignStartOptionsForm_(self)
    Set d.campaignBalanceChoices = New Collection
    d.campaignBalanceChoices.Add OriginalCampaignBalanceChoice_new("Remastered", "The remastered campaign experience", Nothing)
    d.campaignBalanceChoices.Add OriginalCampaignBalanceChoice_new("Classic", """Classic"" campaign balance", Nothing)
    d.campaignBalanceChoices.Add OriginalCampaignBalanceChoice_new("Pumpkin", """Pumpkin"" campaign balance", Nothing)

    Set d.campaignBalanceDropdown = DropdownWidget_new()

    maxItemHeight = 0
    For i = 0 To d.campaignBalanceChoices.Count - 1
        Set opt = d.campaignBalanceChoices(i + 1)
        Set item = WzCampaignOptionButton_make()
        item.setString opt.displayName
        WzCampaignOptionButton_setDescriptionText item, opt.description
        item.setGeometry 0, 0, item.idealWidth(), item.idealHeight()
        maxItemHeight = maxL(maxItemHeight, item.idealHeight())

        If ((i = classicModIdx) Or (i = pumpkinModIdx)) And (opt.modInfo Is Nothing) Then
            ' disable classic balance mod (which is always displayed) by default - updateCampaignBalanceDropdown handles enabling if available
            item.setState WBUT_DISABLE
        End If

        DropdownWidget_addItem d.campaignBalanceDropdown, item
    Next

    DropdownWidget_setListHeight d.campaignBalanceDropdown, maxItemHeight * minL(4, d.campaignBalanceChoices.Count)

    DropdownWidget_setSelectedIndex d.campaignBalanceDropdown, 0

    DropdownWidget_setCanChange d.campaignBalanceDropdown, WzFunc_new(FN_balanceDropdown_canChange)

    weakSelf = ObjPtr(self)
    DropdownWidget_setOnChange d.campaignBalanceDropdown, WzFunc_new(FN_balanceDropdown_onChange, weakSelf)

    ' wrap with info button
    Set infoWrappedCampaignBalanceDropdown = WidgetInfoButtonWrapper_make(d.campaignBalanceDropdown)
    WidgetInfoButtonWrapper_setInfoClickHandler infoWrappedCampaignBalanceDropdown, WzFunc_new(FN_balanceWrapper_infoClick, weakSelf)

    Set CampaignStartOptionsForm_makeCampaignBalanceDropdown = Margin_wrap(0, 10, 0, 10, infoWrappedCampaignBalanceDropdown)
End Function

Private Function CampaignStartOptionsForm_makeCampaignSelectorDropdown(ByVal self As WIDGET) As WIDGET
    Const paddingSize As Long = 10
    Dim d As CampaignStartOptionsForm, i As Long, opt As CAMPAIGN_FILE, item As WIDGET
    Set d = CampaignStartOptionsForm_(self)
    Set d.campaignSelectorDropdown = DropdownWidget_new()
    DropdownWidget_setListHeight d.campaignSelectorDropdown, FRONTEND_CAMPAIGNUI_BUTHEIGHT * minL(5, d.campaignList_size)

    For i = 0 To d.campaignList_size - 1
        opt = d.campaignList(i)
        Set item = makeCampaignUITextButton(0, opt.name_, 0)
        DropdownWidget_addItem d.campaignSelectorDropdown, Margin_wrap(0, paddingSize, 0, paddingSize, item)
    Next

    If d.campaignList_size <> 0 Then
        DropdownWidget_setSelectedIndex d.campaignSelectorDropdown, 0
    End If

    Set CampaignStartOptionsForm_makeCampaignSelectorDropdown = Margin_wrap(0, 10, 0, 10, d.campaignSelectorDropdown)
End Function

Private Sub CampaignStartOptionsForm_updateTweakOptions(ByVal self As WIDGET)
    Dim d As CampaignStartOptionsForm
    Set d = CampaignStartOptionsForm_(self)
    If Not d.tweakOptionsSummary Is Nothing Then
        If Not d.modInfo Is Nothing Then
            assignCollection d.tweakOptionSettings, buildTweakOptionSettings(d.modInfo)
        Else
            assignCollection d.tweakOptionSettings, buildTweakOptionSettings(CampaignStartOptionsForm_getCurrentBalanceModInfo(self))
        End If
        WzCampaignTweakOptionsSummaryWidget_update d.tweakOptionsSummary, d.tweakOptionSettings
    End If
End Sub

Private Sub CampaignStartOptionsForm_closeOverlay(ByVal self As WIDGET)
    Dim d As CampaignStartOptionsForm
    Set d = CampaignStartOptionsForm_(self)
    If d.overlayScreen Is Nothing Then
        Exit Sub
    End If
    widgRemoveOverlayScreen d.overlayScreen
    Set d.overlayScreen = Nothing
    If Not d.overlayScreenOnCloseFunc Is Nothing Then
        WzFunc_call d.overlayScreenOnCloseFunc
    End If
End Sub

Private Sub CampaignStartOptionsForm_openOverlay(ByVal self As WIDGET, ByVal popoverForm As WIDGET, ByVal psTarget As WIDGET, ByVal onCloseFunc As WzFunc)
    Dim d As CampaignStartOptionsForm, lockedScreen As W_SCREEN, newRootFrm As WIDGET, psWeakOverlayScreen As Long, psWeakSelf As Long
    Set d = CampaignStartOptionsForm_(self)
    If Not (d.overlayScreen Is Nothing) Then ASSERT_FAILURE "overlayScreen == nullptr", "src/titleui/campaign.cpp:2445", "openOverlay", "Overlay already open?": Exit Sub

    Set lockedScreen = self.screenPointer
    If Not (Not lockedScreen Is Nothing) Then ASSERT_FAILURE "lockedScreen != nullptr", "src/titleui/campaign.cpp:2448", "openOverlay", "The CampaignStartOptionsForm does not have an associated screen pointer?": Exit Sub

    ' Initialize the options overlay screen
    Set d.overlayScreen = W_SCREEN_make()
    Set d.overlayScreenOnCloseFunc = onCloseFunc
    Set newRootFrm = W_FULLSCREENOVERLAY_CLICKFORM_make()
    psWeakOverlayScreen = ObjPtr(d.overlayScreen)
    psWeakSelf = ObjPtr(self)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc = WzFunc_new(FN_overlay_clicked, psWeakOverlayScreen, psWeakSelf)
    Set W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onCancelPressed = W_FULLSCREENOVERLAY_CLICKFORM_(newRootFrm).onClickedFunc
    d.overlayScreen.psForm.attach newRootFrm

    ' Attach the pop-over form
    newRootFrm.attach popoverForm

    ' Position the pop-over form above the target widget
    popoverForm.setCalcLayout WzFunc_new(FN_overlay_popoverCalcLayout, psWeakSelf)

    widgRegisterOverlayScreenOnTopOfScreen d.overlayScreen, lockedScreen
End Sub

Private Sub CampaignStartOptionsForm_openTweakOptionsEditFormOverlay(ByVal self As WIDGET)
    Dim d As CampaignStartOptionsForm, psWeakSelf As Long, optionsPopOver As WIDGET, maxPopoverHeight As Long
    Dim psTarget As WIDGET, weakSelf As Long
    Set d = CampaignStartOptionsForm_(self)
    If Not (d.overlayScreen Is Nothing) Then ASSERT_FAILURE "overlayScreen == nullptr", "src/titleui/campaign.cpp:2501", "openTweakOptionsEditFormOverlay", "Overlay already open?": Exit Sub

    psWeakSelf = ObjPtr(self)
    Set optionsPopOver = WzCampaignTweakOptionsEditForm_make(d.tweakOptionSettings, WzFunc_new(FN_tweakEditFormOverlay_close, psWeakSelf))
    maxPopoverHeight = self.height() - 50
    optionsPopOver.setGeometry 0, 0, self.width() - 20, maxPopoverHeight ' set once to set width
    optionsPopOver.setGeometry 0, 0, optionsPopOver.width(), minL(maxPopoverHeight, optionsPopOver.idealHeight()) ' set again to set height based on idealHeight for that width (if smaller)

    Set psTarget = d.tweakOptionsSummary
    weakSelf = ObjPtr(self)
    CampaignStartOptionsForm_openOverlay self, optionsPopOver, psTarget, WzFunc_new(FN_tweakEditFormOverlay_onClose, weakSelf)
End Sub

Private Function addCampaignUIMargin(ByVal widget As WIDGET) As WIDGET
    Set addCampaignUIMargin = Margin_wrap(0, 10, 0, 10, widget)   ' Margin(0, 10).wrap(widget)
End Function

' WZOptionTitleLabel

Public Sub WZOptionTitleLabel_construct(ByVal self As WIDGET)
    Dim d As WZOptionTitleLabel
    WIDGET_construct self
    self.kind = K_WZOptionTitleLabel ' (VB) virtual calls from here on reach WZOptionTitleLabel's overrides
    Set d = WZOptionTitleLabel_(self)
    d.FontID = font_regular
    d.verticalPadding = 5
    d.isTruncated = False
End Sub

' std::make_shared<WZOptionTitleLabel>()
Public Function WZOptionTitleLabel_new() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WZOptionTitleLabel)
    WZOptionTitleLabel_construct widget
    Set WZOptionTitleLabel_new = widget
End Function

Public Sub WZOptionTitleLabel_setFont(ByVal self As WIDGET, ByVal newFont As iV_fonts)
    WZOptionTitleLabel_(self).FontID = newFont
End Sub

Public Function WZOptionTitleLabel_idealWidth(ByVal self As WIDGET) As Long
    WZOptionTitleLabel_idealWidth = WZOptionTitleLabel_(self).wzText.width()
End Function

Public Function WZOptionTitleLabel_getString(ByVal self As WIDGET) As String
    WZOptionTitleLabel_getString = WZOptionTitleLabel_(self).wzText.getText()
End Function

Public Sub WZOptionTitleLabel_setString(ByVal self As WIDGET, ByVal string_ As String)
    WZOptionTitleLabel_(self).wzText.setText string_, font_regular
End Sub

Public Function WZOptionTitleLabel_getTip(ByVal self As WIDGET) As String
    Dim d As WZOptionTitleLabel
    Set d = WZOptionTitleLabel_(self)
    If Not d.isTruncated Then
        WZOptionTitleLabel_getTip = ""
        Exit Function
    End If
    WZOptionTitleLabel_getTip = d.wzText.getText()
End Function

Public Sub WZOptionTitleLabel_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As WZOptionTitleLabel, x0 As Long, y0 As Long, textColor As Long, textX0 As Long, textY0 As Long
    Dim maxTextDisplayableWidth As Long
    Set d = WZOptionTitleLabel_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset

    textColor = psPalette(WZCOL_TEXT_MEDIUM)

    textX0 = x0
    textY0 = y0 + d.verticalPadding - d.wzText.aboveBase()

    maxTextDisplayableWidth = self.width()
    d.isTruncated = maxTextDisplayableWidth < d.wzText.width()
    If d.isTruncated Then
        maxTextDisplayableWidth = maxTextDisplayableWidth - (iV_GetEllipsisWidth(d.wzText.getFontID()) + 2)
    End If
    d.wzText.render textX0, textY0, textColor, 0!, maxTextDisplayableWidth
    If d.isTruncated Then
        ' Render ellipsis
        iV_DrawEllipsis d.wzText.getFontID(), textX0 + maxTextDisplayableWidth + 2, textY0, textColor
    End If
End Sub

Private Function makeCampaignUIOptionTitleLabel(ByVal text As String, ByVal font As iV_fonts) As WIDGET
    Dim psLabel As WIDGET
    Set psLabel = WZOptionTitleLabel_new()
    WZOptionTitleLabel_setFont psLabel, font
    psLabel.setString text
    psLabel.setGeometry 0, 0, psLabel.idealWidth(), iV_GetTextLineSize(font)
    Set makeCampaignUIOptionTitleLabel = psLabel
End Function

' (parent: the WzCampaignSelectorTitleUI's address)
Private Sub CampaignStartOptionsForm_initialize(ByVal self As WIDGET, ByVal parent As Long)
    Dim d As CampaignStartOptionsForm, weakParent As Long, weakSelf As Long, grid As WIDGET, row As grid_allocation_slot
    Dim availableDifficultyLevels() As Long, nLevels As Long, filtered() As Long, nFiltered As Long, i As Long
    Dim list() As CAMPAIGN_FILE, n As Long, anyEditable As Boolean, k As Long, o As WzCampaignTweakOptionSetting
    Dim scrollableList As WIDGET, pad As Padding, psStartNewGameButton As WIDGET
    Set d = CampaignStartOptionsForm_(self)
    weakParent = parent
    weakSelf = ObjPtr(self)

    Set d.titleArea = WzCampaignTitleWidget_make(d.modInfo)
    self.attach d.titleArea
    d.titleArea.setCalcLayout WzFunc_new(FN_startForm_titleAreaCalcLayout)

    Set d.subInfoBox = WzCampaignSubInfoBox_make(d.modInfo)
    self.attach d.subInfoBox
    d.subInfoBox.setCalcLayout WzFunc_new(FN_startForm_subInfoBoxCalcLayout)

    Set grid = GridLayout_new()
    row = grid_allocation_slot_(0)

    ' Add Difficulty selector widget
    ReDim availableDifficultyLevels(0 To 4)
    availableDifficultyLevels(0) = DL_SUPER_EASY
    availableDifficultyLevels(1) = DL_EASY
    availableDifficultyLevels(2) = DL_NORMAL
    availableDifficultyLevels(3) = DL_HARD
    availableDifficultyLevels(4) = DL_INSANE
    nLevels = 5
    If Not d.modInfo Is Nothing Then
        If d.modInfo.difficultyLevels.Count > 0 Then
            ' (std::copy_if(..., [modInfoValueRef](DIFFICULTY_LEVEL lvl){ return modInfoValueRef.difficultyLevels.count(lvl) > 0; }))
            nFiltered = 0
            For i = 0 To nLevels - 1
                If d.modInfo.hasDifficultyLevel(availableDifficultyLevels(i)) Then
                    ReDim Preserve filtered(0 To nFiltered)
                    filtered(nFiltered) = availableDifficultyLevels(i)
                    nFiltered = nFiltered + 1
                End If
            Next
            availableDifficultyLevels = filtered
            nLevels = nFiltered
        End If
    End If
    GridLayout_place grid, grid_allocation_slot_(0, 1, False), row, addCampaignUIMargin(makeCampaignUIOptionTitleLabel("Difficulty:", font_regular_bold))
    GridLayout_place grid, grid_allocation_slot_(1, 1, True), row, CampaignStartOptionsForm_makeDifficultySelectorDropdown(self, availableDifficultyLevels, nLevels)
    row.start = row.start + 1

    ' Only display Balance Choice: Remastered / Classic / Pumpkin / etc if !modInfo.has_value() (i.e. if original campaign)
    If d.modInfo Is Nothing Then
        ' Original Campaign - add balance dropdown
        GridLayout_place grid, grid_allocation_slot_(0, 1, False), row, addCampaignUIMargin(makeCampaignUIOptionTitleLabel("Balance:", font_regular_bold))
        GridLayout_place grid, grid_allocation_slot_(1, 1, True), row, CampaignStartOptionsForm_makeCampaignBalanceDropdown(self)
        row.start = row.start + 1
        CampaignStartOptionsForm_updateCampaignBalanceDropdown self, availableCampaignModsCache()
    End If

    ' Add Campaign selector widget
    If Not d.modInfo Is Nothing Then
        n = d.modInfo.getCampaignFiles(list)
    Else
        ' read the base campaign list
        n = readCampaignFiles(list)
    End If
    d.setCampaignList list, n
    GridLayout_place grid, grid_allocation_slot_(0, 1, False), row, addCampaignUIMargin(makeCampaignUIOptionTitleLabel("Start at:", font_regular_bold))
    GridLayout_place grid, grid_allocation_slot_(1, 1, True), row, CampaignStartOptionsForm_makeCampaignSelectorDropdown(self)
    row.start = row.start + 1

    ' Add Tweak Options widget (optional)
    assignCollection d.tweakOptionSettings, buildTweakOptionSettings(d.modInfo)
    anyEditable = False
    For k = 1 To d.tweakOptionSettings.Count
        Set o = d.tweakOptionSettings(k)
        If o.editable Then
            anyEditable = True
            Exit For
        End If
    Next
    If anyEditable Then
        Set d.tweakOptionsSummary = WzCampaignTweakOptionsSummaryWidget_make()
        WzCampaignTweakOptionsSummaryWidget_update d.tweakOptionsSummary, d.tweakOptionSettings
        WzCampaignTweakOptionsSummaryWidget_setInfoClickHandler d.tweakOptionsSummary, WzFunc_new(FN_startForm_tweakSummaryInfoClick, weakSelf)
        GridLayout_place grid, grid_allocation_slot_(0, 1, False), row, addCampaignUIMargin(makeCampaignUIOptionTitleLabel("Tweaks:", font_regular_bold))
        GridLayout_place grid, grid_allocation_slot_(1, 1, True), row, Margin_wrap(0, 10, 0, 10, d.tweakOptionsSummary)
        row.start = row.start + 1
    End If

    grid.setGeometry 0, 0, FRONTEND_BUTWIDTH, grid.idealHeight()

    Set scrollableList = ScrollableListWidget_make()
    ScrollableListWidget_addItem scrollableList, grid
    pad.top = 4
    pad.right_ = 4
    pad.bottom = 4
    pad.left_ = 4
    ScrollableListWidget_setPadding scrollableList, pad
    self.attach scrollableList
    scrollableList.setCalcLayout WzFunc_new(FN_startForm_scrollableListCalcLayout)

    Set psStartNewGameButton = CampaignStartButton_make()
    self.attach psStartNewGameButton
    psStartNewGameButton.setCalcLayout WzFunc_new(FN_startForm_startButtonCalcLayout)
    W_BUTTON_addOnClickHandler psStartNewGameButton, WzFunc_new(FN_startForm_startButtonClick, weakSelf)
    If Not d.modInfo Is Nothing Then
        If d.modInfo.getModCompatibility() = WzModCompatibilityResult_NOT_COMPATIBLE Then
            psStartNewGameButton.setState WBUT_DISABLE
        End If
    End If
End Sub

Public Sub CampaignStartOptionsForm_geometryChanged(ByVal self As WIDGET)
    Dim i As Long
    For i = 0 To self.childCount() - 1
        self.childAt(i).callCalcLayout
    Next
End Sub

' MARK: - WzCampaignSelectorTitleUI (WzCampaignSelectorTitleUI.cls)

Public Sub addCampaignBaseForm()
    Dim parent As WIDGET, botForm As WIDGET
    Set parent = widgGetFromID(psWScreen, FRONTEND_BACKDROP)

    Set botForm = IntFormAnimated_new()
    parent.attach botForm
    botForm.id = FRONTEND_BOTFORM

    botForm.setCalcLayout WzFunc_new(FN_campaignBaseForm_calcLayout)
End Sub

' the lambdas
Public Function campaign_title_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Dim psWidget As WIDGET, parent As WIDGET, psParent As WIDGET, strongParent As WIDGET, psSelf As WIDGET, strongSelf As WIDGET
    Dim strongEditForm As WIDGET, pTweakOptionsToggle As WIDGET, x0 As Long, y0 As Long, w As Long, h As Long, x As Long
    Dim v As Variant, newIndex As Long, newSelectedWidget As WIDGET, isDisabled As Boolean, labelFontColor As Long, k As Long
    Dim psOverlayScreen As W_SCREEN, popOverX0 As Long, popOverY0 As Long, strongTitleUI As WzCampaignSelectorTitleUI
    Dim campaignMods As Collection, pLeftPanel As WIDGET, pDisplayCampaignStartOptions As WIDGET, o As WzCampaignTweakOptionSetting
    Select Case f.id
    Case FN_campaignModButton_onClick           ' [weakParent, modInfo](W_BUTTON& but)
        Set psSelf = dynCast(args(0), K_WzCampaignModButton)
        Set strongParent = WIDGET_fromPtr(f.cap(0))
        If Not WzCampaignModButton_getIsSelected(psSelf) Then
            WzCampaignSelectorForm_changeSelectedItem strongParent, f.cap(1)
            WzCampaignModButton_setIsSelected psSelf, True
        End If

    Case FN_campaignListSection_titleCalcLayout ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set parent = dynCast(psWidget.parent(), K_WzCampaignListSection)
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/titleui/campaign.cpp:416", "initialize", "No parent?": Exit Function
        x0 = 15
        w = minL(W_LABEL_getMaxLineWidth(psWidget), parent.width() - x0)
        psWidget.setGeometry x0, 0, w, iV_GetTextLineSize(font_regular_bold)

    Case FN_campaignListSection_listCalcLayout  ' [](WIDGET* psWidget)
        Set psWidget = args(0)
        Set parent = dynCast(psWidget.parent(), K_WzCampaignListSection)
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/titleui/campaign.cpp:426", "initialize", "No parent?": Exit Function
        y0 = WzCampaignListSection_(parent).sectionTitle.y() + WzCampaignListSection_(parent).sectionTitle.height() + 10
        h = parent.height() - y0
        psWidget.setGeometry 0, y0, parent.width(), h

    Case FN_selectorForm_originalCalcLayout     ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = dynCast(psWidget.parent(), K_WzCampaignSelectorForm)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:505", "initialize", "No parent?": Exit Function
        psWidget.setGeometry 0, 0, psParent.width(), psWidget.height()

    Case FN_selectorForm_alternateCalcLayout    ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = dynCast(psWidget.parent(), K_WzCampaignSelectorForm)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:514", "initialize", "No parent?": Exit Function
        y0 = WzCampaignSelectorForm_(psParent).originalCampaignButton.y() + WzCampaignSelectorForm_(psParent).originalCampaignButton.height() + 20
        psWidget.setGeometry 0, y0, psParent.width(), psParent.height() - y0

    Case FN_campaign_closeButtonClick           ' [closeHandlerFunc](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_campaign_closeButtonTask, f.cap(0))

    Case FN_campaign_closeButtonTask            ' [closeHandlerFunc]()
        WzFunc_call f.cap(0)

    Case FN_tweakEditForm_resetClick            ' [weakEditForm](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_tweakEditForm_resetTask, f.cap(0))

    Case FN_tweakEditForm_resetTask             ' [weakEditForm]()
        Set strongEditForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongEditForm Is Nothing) Then ASSERT_FAILURE "strongEditForm != nullptr", "src/titleui/campaign.cpp:868", "initialize", "No parent?": Exit Function
        WzCampaignTweakOptionsEditForm_resetToDefaults strongEditForm

    Case FN_tweakEditForm_toggleClick           ' [weakEditForm, i](W_BUTTON& but)
        Set strongEditForm = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongEditForm Is Nothing) Then ASSERT_FAILURE "strongEditForm != nullptr", "src/titleui/campaign.cpp:892", "initialize", "No parent?": Exit Function
        Set pTweakOptionsToggle = dynCast(args(0), K_WzAdvCheckbox)
        If Not (Not pTweakOptionsToggle Is Nothing) Then ASSERT_FAILURE "pTweakOptionsToggle != nullptr", "src/titleui/campaign.cpp:894", "initialize", "Invalid button type": Exit Function
        Set o = WzCampaignTweakOptionsEditForm_(strongEditForm).tweakOptions(f.cap(1) + 1)   ' tweakOptions->at(i)
        o.currentValue = WzAdvCheckbox_isChecked(pTweakOptionsToggle)
        WzCampaignTweakOptionsEditForm_updateResetButtonStatus strongEditForm

    Case FN_infoWrapper_infoClick               ' [weakSelf](W_BUTTON&)
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:1060", "initialize", "No parent?": Exit Function
        If Not WidgetInfoButtonWrapper_(strongSelf).infoClickHandler Is Nothing Then
            WzFunc_call WidgetInfoButtonWrapper_(strongSelf).infoClickHandler
        End If

    Case FN_tweakSummary_infoCalcLayout         ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = dynCast(psWidget.parent(), K_WzCampaignTweakOptionsSummaryWidget)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:1230", "initialize", "No parent?": Exit Function
        w = 16 + WzCampaignTweakOptionsSummaryWidget_(psParent).internalHorizontalPadding
        x0 = psParent.width() - w
        h = psParent.height() - (WzCampaignTweakOptionsSummaryWidget_(psParent).topPadding + WzCampaignTweakOptionsSummaryWidget_(psParent).bottomPadding)
        psWidget.setGeometry x0, WzCampaignTweakOptionsSummaryWidget_(psParent).topPadding, w, h

    Case FN_tweakSummary_infoClick, FN_tweakSummary_listClick   ' [weakSelf](W_BUTTON&), [weakSelf](ClickableScrollableList&)
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If f.id = FN_tweakSummary_infoClick Then
            If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:1240", "initialize", "No parent?": Exit Function
        Else
            If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:1258", "initialize", "No parent?": Exit Function
        End If
        If Not WzCampaignTweakOptionsSummaryWidget_(strongSelf).infoClickHandler Is Nothing Then
            WzFunc_call WzCampaignTweakOptionsSummaryWidget_(strongSelf).infoClickHandler
        End If

    Case FN_tweakSummary_listCalcLayout         ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = dynCast(psWidget.parent(), K_WzCampaignTweakOptionsSummaryWidget)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:1251", "initialize", "No parent?": Exit Function
        w = psParent.width() - WzCampaignTweakOptionsSummaryWidget_(psParent).infoButton.width() - (WzCampaignTweakOptionsSummaryWidget_(psParent).internalHorizontalPadding * 2)
        h = psParent.height() - (WzCampaignTweakOptionsSummaryWidget_(psParent).topPadding + WzCampaignTweakOptionsSummaryWidget_(psParent).bottomPadding)
        psWidget.setGeometry WzCampaignTweakOptionsSummaryWidget_(psParent).internalHorizontalPadding, WzCampaignTweakOptionsSummaryWidget_(psParent).topPadding, w, h

    Case FN_tweakSummary_listHighlight          ' [weakSelf](ClickableScrollableList&, bool isHighlighted)
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:1266", "initialize", "No parent?": Exit Function
        If CBool(args(1)) Then labelFontColor = psPalette(WZCOL_TEXT_BRIGHT) Else labelFontColor = psPalette(WZCOL_TEXT_MEDIUM)
        For k = 1 To WzCampaignTweakOptionsSummaryWidget_(strongSelf).tweakOptionsLabels.Count
            W_LABEL_setFontColour WzCampaignTweakOptionsSummaryWidget_(strongSelf).tweakOptionsLabels(k), labelFontColor
        Next

    Case FN_titleWidget_labelCalcLayout         ' LAMBDA_CALCLAYOUT_SIMPLE (nameLabel, authorLabel)
        Set psWidget = args(0)
        Set parent = psWidget.parent()
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/titleui/campaign.cpp:1328", "initialize", "No parent?": Exit Function
        w = maxL(parent.width() - (psWidget.x() * 2), 0)
        psWidget.setGeometry psWidget.x(), psWidget.y(), w, psWidget.height()

    Case FN_titleWidget_descriptionCalcLayout   ' [descriptionWidget](WIDGET *psWidget)
        Set psWidget = args(0)
        Set parent = psWidget.parent()
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/titleui/campaign.cpp:1365", "initialize", "No parent?": Exit Function
        If parent.width() <= 1 Or parent.height() <= 1 Then
            Exit Function
        End If
        w = maxL(parent.width() - (psWidget.x() * 2), 0)
        If w = 0 Then
            Exit Function
        End If
        psWidget.setGeometry psWidget.x(), psWidget.y(), w, 45

    Case FN_subInfoBox_universeCalcLayout       ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set parent = psWidget.parent()
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/titleui/campaign.cpp:1475", "initialize", "No parent?": Exit Function
        w = psWidget.width()
        x = parent.width() - 40 - w
        psWidget.setGeometry x, 0, w, parent.height()

    Case FN_subInfoBox_compatibilityCalcLayout  ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set parent = dynCast(psWidget.parent(), K_WzCampaignSubInfoBox)
        If Not (Not parent Is Nothing) Then ASSERT_FAILURE "parent != nullptr", "src/titleui/campaign.cpp:1521", "initialize", "No parent?": Exit Function
        w = psWidget.width()
        x = 40 + WzCampaignSubInfoBox_(parent).imgDimensions + WzCampaignSubInfoBox_(parent).imgTextPadding
        psWidget.setGeometry x, 0, w, parent.height()

    Case FN_difficultyDropdown_canChange        ' [dropDownChoices](DropdownWidget &widget, size_t newIndex, newSelectedWidget) -> bool
        v = f.cap(0)
        newIndex = args(1)
        If IsArray(v) Then k = UBound(v) + 1 Else k = 0     ' dropDownChoices.size()
        If Not (ULT(newIndex, k)) Then ASSERT_FAILURE "newIndex < dropDownChoices.size()", "src/titleui/campaign.cpp:2015", "makeDifficultySelectorDropdown", "Invalid index": campaign_title_c_fn = False: Exit Function
        setDifficultyLevel v(newIndex)
        campaign_title_c_fn = True

    Case FN_balanceDropdown_canChange           ' [](DropdownWidget &widget, size_t newIndex, newSelectedWidget) -> bool
        Set newSelectedWidget = args(2)
        If Not (Not newSelectedWidget Is Nothing) Then ASSERT_FAILURE "newSelectedWidget != nullptr", "src/titleui/campaign.cpp:2367", "makeCampaignBalanceDropdown", "Nothing selected?": campaign_title_c_fn = False: Exit Function
        isDisabled = (newSelectedWidget.getState() And WBUT_DISABLE) <> 0
        campaign_title_c_fn = Not isDisabled

    Case FN_balanceDropdown_onChange            ' [weakSelf](DropdownWidget &widget)
        widgScheduleTask WzFunc_new(FN_balanceDropdown_onChangeTask, f.cap(0))

    Case FN_balanceDropdown_onChangeTask        ' [weakSelf]()
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:2376", "makeCampaignBalanceDropdown", "No parent?": Exit Function
        CampaignStartOptionsForm_updateTweakOptions strongSelf

    Case FN_balanceWrapper_infoClick            ' [weakSelf]()
        debug_ 2384, LOG_INFO, "makeCampaignBalanceDropdown", "Info button clicked"
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:2386", "makeCampaignBalanceDropdown", "No parent?": Exit Function
        CampaignStartOptionsForm_openCampaignBalanceInfoOverlay strongSelf

    Case FN_balanceInfo_closeOverlay, FN_tweakEditFormOverlay_close     ' [weakSelf]() / [psWeakSelf]()
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not strongSelf Is Nothing Then
            CampaignStartOptionsForm_closeOverlay strongSelf
        End If

    Case FN_balanceInfo_onClose                 ' [weakSelf]()
        ' no-op onCloseFunc

    Case FN_overlay_clicked                     ' [psWeakOverlayScreen, psWeakSelf]()
        ' Destroy overlay screen
        Set strongSelf = WIDGET_fromPtr(f.cap(1))
        If Not strongSelf Is Nothing Then
            CampaignStartOptionsForm_closeOverlay strongSelf
        Else
            ' backup in case there's a lingering overlay screen
            Set psOverlayScreen = W_SCREEN_fromPtr(f.cap(0))
            If Not psOverlayScreen Is Nothing Then
                widgRemoveOverlayScreen psOverlayScreen
            End If
        End If

    Case FN_overlay_popoverCalcLayout           ' [psWeakSelf](WIDGET *psWidget)
        Set psWidget = args(0)
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:2480", "openOverlay", "parent is null": Exit Function
        ' (Ideally, with its left aligned with the left side of the "CampaignStartOptionsForm" widget, but ensure full visibility on-screen)
        popOverX0 = strongSelf.screenPosX() + (strongSelf.width() - psWidget.width()) \ 2
        If popOverX0 + psWidget.width() > screenWidth Then
            popOverX0 = screenWidth - psWidget.width()
        End If
        ' (Ideally, with its middle centered with the middle of the "CampaignStartOptionsForm" widget, but ensure full visibility on-screen)
        popOverY0 = maxL(strongSelf.screenPosY() + (strongSelf.height() - psWidget.height()) \ 2, 0)
        If popOverY0 + psWidget.height() > screenHeight Then
            popOverY0 = screenHeight - psWidget.height()
        End If
        psWidget.move popOverX0, popOverY0

    Case FN_tweakEditFormOverlay_onClose        ' [weakSelf]()
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not strongSelf Is Nothing Then
            ' refresh any possibly-edited data
            WzCampaignTweakOptionsSummaryWidget_update CampaignStartOptionsForm_(strongSelf).tweakOptionsSummary, CampaignStartOptionsForm_(strongSelf).tweakOptionSettings
        End If

    Case FN_startForm_titleAreaCalcLayout       ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:2613", "initialize", "No parent?": Exit Function
        psWidget.setGeometry 0, 0, psParent.width(), 155

    Case FN_startForm_subInfoBoxCalcLayout      ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = dynCast(psWidget.parent(), K_CampaignStartOptionsForm)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:2621", "initialize", "No parent?": Exit Function
        y0 = CampaignStartOptionsForm_(psParent).titleArea.y() + CampaignStartOptionsForm_(psParent).titleArea.height()
        psWidget.setGeometry 0, y0, psParent.width(), 30

    Case FN_startForm_tweakSummaryInfoClick     ' [weakSelf]()
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:2674", "initialize", "No parent?": Exit Function
        CampaignStartOptionsForm_openTweakOptionsEditFormOverlay strongSelf

    Case FN_startForm_scrollableListCalcLayout  ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        Set psParent = dynCast(psWidget.parent(), K_CampaignStartOptionsForm)
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:2690", "initialize", "No parent?": Exit Function
        y0 = CampaignStartOptionsForm_(psParent).titleArea.y() + CampaignStartOptionsForm_(psParent).titleArea.height() + 40
        psWidget.setGeometry 5, y0, psParent.width() - 10, 160

    Case FN_startForm_startButtonCalcLayout     ' [](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:2699", "initialize", "No parent?": Exit Function
        h = psWidget.height()
        y0 = psParent.height() - h - 20
        psWidget.setGeometry 30, y0, maxL(psParent.width() - 60, 50), psWidget.height()

    Case FN_startForm_startButtonClick          ' [weakSelf](W_BUTTON&)
        widgScheduleTask WzFunc_new(FN_startForm_startButtonTask, f.cap(0))

    Case FN_startForm_startButtonTask           ' [weakSelf]()
        Set strongSelf = WIDGET_fromPtr(f.cap(0))
        If Not (Not strongSelf Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:2707", "initialize", "No parent?": Exit Function
        CampaignStartOptionsForm_handleStartGame strongSelf

    Case FN_campaignBaseForm_calcLayout         ' LAMBDA_CALCLAYOUT_SIMPLE
        Set psWidget = args(0)
        psWidget.setGeometry FRONTEND_BOTFORM_WIDEX, FRONTEND_TOPFORM_WIDEY, 590, 460

    Case FN_campaignSelector_quitClick          ' [weakSelf](W_BUTTON& button)
        widgScheduleTask WzFunc_new(FN_campaignSelector_quitTask, f.cap(0))

    Case FN_campaignSelector_quitTask           ' [weakSelf]()
        Set strongTitleUI = campaignSelectorTitleUI_lock(f.cap(0))
        If Not (Not strongTitleUI Is Nothing) Then ASSERT_FAILURE "strongSelf != nullptr", "src/titleui/campaign.cpp:2771", "start", "No parent": Exit Function
        changeTitleUI strongTitleUI.getParentTitleUI() ' go back

    Case FN_campaignSelector_calcLayout         ' [campaignSelectorY0](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:2781", "start", "No parent?": Exit Function
        h = psParent.height() - f.cap(0)
        psWidget.setGeometry 1, f.cap(0), 180, h

    Case FN_loadCampaignMods_completion         ' [weakSelf](std::vector<WzCampaignModInfo> campaignMods)
        wzAsyncExecOnMainThread WzFunc_new(FN_loadCampaignMods_mainThread, args(0), f.cap(0))

    Case FN_loadCampaignMods_mainThread         ' [campaignMods, weakSelf]()
        Set campaignMods = f.cap(0)
        debug_ 2791, LOG_INFO, "start", "Found %zu campaign mods", campaignMods.Count

        Set strongTitleUI = campaignSelectorTitleUI_lock(f.cap(1))
        If strongTitleUI Is Nothing Then
            ' nothing to do - the titleUI instance that triggered this async method not longer exists
            Exit Function
        End If

        Set cached_availableCampaignMods = campaignMods

        If cached_availableCampaignMods.Count > 0 Then
            WzCampaignSelectorForm_updateList strongTitleUI.getCampaignSelector(), cached_availableCampaignMods
            ' Update the original campaign panel's list of balance options (which is currently being displayed)
            Set pDisplayCampaignStartOptions = dynCast(strongTitleUI.getDisplayedPanel(), K_CampaignStartOptionsForm)
            If Not pDisplayCampaignStartOptions Is Nothing Then
                CampaignStartOptionsForm_updateCampaignBalanceDropdown pDisplayCampaignStartOptions, cached_availableCampaignMods
            End If
        End If

    Case FN_displayedPanel_calcLayout           ' [weakSelf](WIDGET *psWidget)
        Set psWidget = args(0)
        Set psParent = psWidget.parent()
        If Not (Not psParent Is Nothing) Then ASSERT_FAILURE "psParent != nullptr", "src/titleui/campaign.cpp:2840", "displayStartOptionsForm", "No parent?": Exit Function
        Set strongTitleUI = campaignSelectorTitleUI_lock(f.cap(0))
        Set pLeftPanel = strongTitleUI.getCampaignSelector()
        x = pLeftPanel.x() + pLeftPanel.width() + 1
        w = psParent.width() - x - 1
        psWidget.setGeometry x, 1, w, psParent.height() - 2
    End Select
End Function
