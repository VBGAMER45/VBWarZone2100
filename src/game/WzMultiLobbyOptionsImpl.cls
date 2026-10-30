VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzMultiLobbyOptionsImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/multilobbyoptions.cpp: the fields of class WzMultiLobbyOptionsImpl
' (VB) data of widget class WzMultiLobbyOptionsImpl
Option Explicit

Public lobbyStatusBar As WIDGET         ' std::shared_ptr<WzMultiLobbyStatusBar>

Public gameNameEditBox As WIDGET        ' std::shared_ptr<W_EDITBOX>

Public mapNameWidget As WIDGET          ' std::shared_ptr<W_LABEL>

Public optionsList As WIDGET            ' std::shared_ptr<ScrollableListWidget>
Public matchOptionsConfigureButton As WIDGET ' std::shared_ptr<WzFrontendImageButton>
Public scavengerChoice As WIDGET        ' std::shared_ptr<LobbyMultiOptionsChoiceWidget>
Public allianceChoice As WIDGET
Public powerChoice As WIDGET
Public baseTypeChoice As WIDGET
Public technologyChoice As WIDGET
Public multiOptionButtons As Collection ' std::vector<std::shared_ptr<WzLobbyMultiOptionButton>>: the WIDGETs, 1-based

Public limitIconsWidget As WIDGET       ' std::shared_ptr<MultibuttonWidget>
Public limitIconsWidgetWrapper As WIDGET ' std::shared_ptr<WIDGET>
Public roomPasswordOptionsButton As WIDGET ' std::shared_ptr<WzFrontendImageButton>
Public lobbyOptionButtonUpdateFuncs As Collection ' std::vector<std::function<void ()>>: the WzFuncs, 1-based

Public paddingXLeft As Long             ' int paddingXLeft = 13;
Public paddingXRight As Long            ' int paddingXRight = 10;
Public paddingY As Long                 ' int paddingY = 3;

Public isChallenge As Boolean           ' bool isChallenge = false;
Public has_showTotalVotes As Boolean    ' optional<bool> showTotalVotes = nullopt;: has_value()
Public showTotalVotes As Boolean        ' its value

Public optionsPopover As WIDGET         ' std::shared_ptr<PopoverWidget>
Public cachedPasswordFieldValue As String ' WzString
Public weakSerial As Long               ' (VB) this widget's serial for weak pointers to it (widgweak_vb)
