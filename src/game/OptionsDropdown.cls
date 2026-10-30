VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionsDropdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.h: the fields of class OptionsDropdown
' (VB) data of widget class OptionsDropdown
Option Explicit

Public current As OptionChoices          ' OptionChoices<T>
Public populateFunc As WzFunc            ' PopulateFunc: OptionChoices<T> (void)
Public onSelectFunc As WzFunc            ' OnSelectItemFunc: bool (const T& newValue)
Public maxListItemsShown As Long
Public needsStateUpdates As Boolean
