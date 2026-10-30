VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItemWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class MenuItemWrapper
' lib/widget/popovermenu.h: MenuItemWrapper's fields (popovermenu_c). std::weak_ptr<PopoverMenuWidget>
' parent is the menu's serial (widgweak_vb).
Option Explicit

Public item As WIDGET
Public parent As Long
Public closeMenuOnClick As Boolean      ' = false
Public mouseDownOnWrapper As Boolean    ' = false
