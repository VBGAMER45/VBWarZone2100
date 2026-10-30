VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_EDITBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_EDITBOX
' lib/widget/editbox.h: W_EDITBOX's fields (editbox_c). EditBoxDisplayCache displayCache is wzDisplayedText
' and modeText; optional<PIELIGHT> fixedPlaceholderTextColor is has_fixedPlaceholderTextColor and
' fixedPlaceholderTextColor. The handlers are WzFuncs (Nothing = nullptr).
Option Explicit

Public state As Long                    ' UDWORD: The current edit box state
Public aText As String                  ' WzString: The text in the edit box
Public placeholderText As String
Public FontID As iV_fonts
Public blinkOffset As Long              ' Cursor should be visible at time blinkOffset.
Public maxStringSize As Long            ' max characters string will accept
Public insPos As Long                   ' The insertion point in the buffer
Public printStart As Long               ' Where in the string appears at the far left of the box
Public printChars As Long               ' The number of characters appearing in the box
Public printWidth As Long               ' The pixel width of the characters in the box
Public pBoxDisplay As WzFunc            ' WIDGET_DISPLAY: Optional callback to display the edit box background.
Public HilightAudioID As Long           ' SWORD: Audio ID for form clicked sound
Public ClickedAudioID As Long           ' SWORD: Audio ID for form hilighted sound
Public ErrorAudioID As Long             ' SWORD: Audio ID for error sound
Public AudioCallback As WzFunc          ' WIDGET_AUDIOCALLBACK: Pointer to audio callback function
Public boxColourFirst As Long           ' PIELIGHT
Public boxColourSecond As Long
Public boxColourBackground As Long
Public has_fixedPlaceholderTextColor As Boolean
Public fixedPlaceholderTextColor As Long
Public wzDisplayedText As WzText        ' displayCache.wzDisplayedText
Public modeText As WzText               ' displayCache.modeText
Public suppressAudioCallback As Boolean ' = false
Public onRetHandler As WzFunc           ' OnReturnHandler
Public onTabHandler As WzFunc           ' OnTabHandler
Public onEscHandler As WzFunc           ' OnReturnHandler
Public onEditingStoppedHandler As WzFunc ' OnReturnHandler
Public onPasteTransformFunc As WzFunc   ' OnPasteTransformFunc
Public pTip As String                   ' std::string: The tool tip for the edit box
