VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzPlayerRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/lobbyplayerrow.h: the fields of class WzPlayerRow
' (VB) data of widget class WzPlayerRow
Option Explicit

Public parentTitleUI As Long            ' std::weak_ptr<WzMultiplayerOptionsTitleUI>: its address
Public playerIdx As Long                ' unsigned playerIdx = 0;
Public teamButton As WIDGET             ' std::shared_ptr<W_BUTTON>
Public colorButton As WIDGET            ' std::shared_ptr<W_BUTTON>
Public playerInfo As WIDGET             ' std::shared_ptr<W_BUTTON>
Public readyButtonContainer As WIDGET   ' std::shared_ptr<WIDGET>
Public difficultyChooserButton As WIDGET ' std::shared_ptr<W_BUTTON>
Public readyButton As WIDGET            ' std::shared_ptr<W_BUTTON>
Public readyTextLabel As WIDGET         ' std::shared_ptr<W_LABEL>
