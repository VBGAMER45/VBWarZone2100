VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzLobbyPasswordOptionPopover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/multilobbyoptions.cpp: the fields of class WzLobbyPasswordOptionPopover
' (VB) data of widget class WzLobbyPasswordOptionPopover
Option Explicit

Public weakParentMultiOptions As Long   ' std::weak_ptr<WzMultiLobbyOptionsWidgetBase>: the WzMultiLobbyOptionsImpl's serial (widgweak_vb)
Public passwordEditBox As WIDGET        ' std::shared_ptr<W_EDITBOX>
Public requiredPasswordCheckbox As WIDGET ' std::shared_ptr<WzAdvCheckbox>
Public paddingX As Long                 ' int paddingX = 15;
Public paddingY As Long                 ' int paddingY = 15;
Public innerPaddingY As Long            ' int innerPaddingY = 10;
Public editBoxHeight As Long            ' int editBoxHeight = 24;
Public weakSerial As Long               ' (VB) this widget's serial for weak pointers to it (widgweak_vb)
