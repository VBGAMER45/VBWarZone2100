VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopoverWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class PopoverWidget
' lib/widget/popover.h: PopoverWidget's fields (popover_c). std::weak_ptr<W_SCREEN> overlayScreen: see
' popover_c. Vector2i positionOffset is positionOffsetX, positionOffsetY. weakSerial: widgweak_vb.
Option Explicit

Public weakSerial As Long
Public overlayScreen As W_SCREEN
Public popoverContents As WIDGET
Public onCloseHandler As WzFunc         ' OnPopoverCloseHandlerFunc
Public style As PopoverWidget_Style
Public align As PopoverWidget_Alignment
Public positionOffsetX As Long
Public positionOffsetY As Long
