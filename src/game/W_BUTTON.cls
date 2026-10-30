VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_BUTTON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_BUTTON
' lib/widget/button.h: W_BUTTON's fields (button_c); the initialisers are W_BUTTON_construct's.
' images and progressBorder are UDTs (button_c's W_BUTTON_Images, W_BUTTON_ProgressBorder): Friend
' properties (a copy is read; W_BUTTON_setImages / W_BUTTON_setProgressBorder write them).
Option Explicit

Public state As Long                    ' unsigned: The current button state
Public pText As String                  ' WzString: The text for the button
Private images_ As W_BUTTON_Images      ' The images for the button.
Public pTip As String                   ' std::string: The tool tip for the button
Public HilightAudioID As Long           ' SWORD: Audio ID for form clicked sound
Public ClickedAudioID As Long           ' SWORD: Audio ID for form hilighted sound
Public AudioCallback As WzFunc          ' WIDGET_AUDIOCALLBACK: Pointer to audio callback function
Public FontID As iV_fonts
Public minClickInterval As Long         ' UDWORD = 0

' private
Public lastClickTime As Long            ' UDWORD = 0
Public onClickHandlers As Collection    ' std::vector<W_BUTTON_ONCLICK_FUNC>: WzFuncs
Public has_progressBorder As Boolean    ' optional<ProgressBorder> progressBorder
Private progressBorder_ As W_BUTTON_ProgressBorder
Public progressBorderColour As Long     ' PIELIGHT
Public lastClickButton As WIDGET_KEY    ' = WKEY_NONE
Public help As WidgetHelp               ' optional<WidgetHelp>: Nothing = nullopt
Public has_clickDownStart As Boolean    ' optional<std::chrono::steady_clock::time_point> clickDownStart: the start time of click down on this button
Public clickDownStart As Double         ' (widgSteadyNowMs)
Public has_clickDownKey As Boolean      ' optional<WIDGET_KEY> clickDownKey
Public clickDownKey As WIDGET_KEY

Private Sub Class_Initialize()
    Set onClickHandlers = New Collection
End Sub

Friend Property Get images() As W_BUTTON_Images
    images = images_
End Property

Friend Property Let images(v As W_BUTTON_Images)
    images_ = v
End Property

Friend Property Get progressBorder() As W_BUTTON_ProgressBorder
    progressBorder = progressBorder_
End Property

Friend Property Let progressBorder(v As W_BUTTON_ProgressBorder)
    progressBorder_ = v
End Property
