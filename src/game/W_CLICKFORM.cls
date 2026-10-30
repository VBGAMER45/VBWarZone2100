VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_CLICKFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_CLICKFORM
' lib/widget/form.h: W_CLICKFORM's fields (form_c); the initialisers are W_CLICKFORM_construct's.
Option Explicit

Public state As Long                    ' unsigned: Button state of the form

' private
Public pTip As String                   ' std::string: Tip for the form
Public help As WidgetHelp               ' optional<WidgetHelp>: Nothing = nullopt
Public HilightAudioID As Long           ' SWORD: Audio ID for form clicked sound
Public ClickedAudioID As Long           ' SWORD: Audio ID for form hilighted sound
Public AudioCallback As WzFunc          ' WIDGET_AUDIOCALLBACK: Pointer to audio callback function
Public has_clickDownStart As Boolean    ' optional<std::chrono::steady_clock::time_point> clickDownStart: the start time of click down on this form
Public clickDownStart As Double         ' (widgSteadyNowMs)
Public has_clickDownKey As Boolean      ' optional<WIDGET_KEY> clickDownKey
Public clickDownKey As WIDGET_KEY
