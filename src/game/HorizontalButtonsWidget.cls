VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorizontalButtonsWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/optionsforms.cpp: the fields of class HorizontalButtonsWidget
' (VB) data of widget class HorizontalButtonsWidget
Option Explicit

Public weakSerial As Long               ' (VB) its std::weak_ptr serial (widgweak_vb)
Public onSelectForm As WzFunc            ' OnSelectForm: void(const WzString& untranslatedTitle)
Public buttonFontID As iV_fonts
Public clickLock As Boolean              ' = false
Public horizontalButtonPadding As Long   ' = 15
Public verticalButtonPadding As Long     ' = 10
Public verticalOuterPadding As Long      ' = 0
Public has_backgroundColor As Boolean    ' optional<PIELIGHT> backgroundColor = nullopt
Public backgroundColor As Long
Public optionsFormUntranslatedTitles As Collection ' std::vector<WzString>
Public clipContainer As WIDGET           ' std::shared_ptr<ClipRectWidget>
Public optionsTitles As Collection       ' std::vector<std::shared_ptr<WzClippableButton>>
Public has_currentButtonIdx As Boolean   ' optional<size_t> currentButtonIdx = nullopt
Public currentButtonIdx As Long
Public lastMouseOverInternalWidget As WIDGET
Public scrollLeftButton As WIDGET        ' std::shared_ptr<HorizontalScrollButton>
Public scrollRightButton As WIDGET
Public scrollableWidth As Long           ' = 0
Public maxItemHeight As Long             ' = 0
Public itemSpacing As Long               ' = 0
Public clipContainerInnerPaddingX As Long ' = 28
Public currentXScrollPos As Long         ' = 0
Public has_dragStartXScrollPos As Boolean ' optional<int32_t> dragStartXScrollPos = nullopt
Public dragStartXScrollPos As Long
Public isHandlingDrag As Boolean         ' = false
Public layoutDirty As Boolean            ' = true
