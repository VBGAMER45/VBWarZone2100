VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopoverMenuWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class PopoverMenuWidget
' lib/widget/popovermenu.h: PopoverMenuWidget's fields (popovermenu_c). std::weak_ptr<PopoverWidget>
' currentPopover is the popover's serial (widgweak_vb), 0 when empty.
Option Explicit

Public weakSerial As Long
Public menuList As WIDGET               ' std::shared_ptr<ScrollableListWidget>
Public currentPopover As Long
Public mouseDownItem As WIDGET          ' std::shared_ptr<MenuItemWrapper>
Public maxItemHeight As Long            ' int32_t = 0
