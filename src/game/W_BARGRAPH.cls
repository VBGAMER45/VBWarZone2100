VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_BARGRAPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_BARGRAPH
' lib/widget/bar.h: W_BARGRAPH's fields (bar_c); the initialisers are W_BARGRAPH_construct_init's.
Option Explicit

Public barPos As WBAR_ORIENTATION       ' Orientation of the bar on the widget
Public majorSize As Long                ' UWORD: Percentage of the main bar that is filled
Public minorSize As Long                ' UWORD: Percentage of the minor bar if there is one
Public iRange As Long                   ' UWORD: Maximum range
Public majorValue As Long               ' UWORD: Untransformed major value
Public minorValue As Long               ' UWORD: Untransformed minor value
Public denominator As Long              ' Denominator, 1 by default.
Public precision_ As Long               ' precision: Number of places after the decimal point to display, 0 by default.
Public majorCol As Long                 ' PIELIGHT: Colour for the major bar
Public minorCol As Long                 ' PIELIGHT: Colour for the minor bar
Public textCol As Long                  ' PIELIGHT: Colour for the text on the bar.
Public pTip As String                   ' std::string: The tool tip for the graph
Public text As String                   ' WzString: Text on the bar.

'private:
Public backgroundColour As Long         ' PIELIGHT
Public wzCachedText As WzText

Public sizesDirty As Boolean            ' = false

' private
Public help As WidgetHelp               ' optional<WidgetHelp>: Nothing = nullopt

Private Sub Class_Initialize()
    Set wzCachedText = New WzText
End Sub
