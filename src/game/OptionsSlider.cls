VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionsSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.h: the fields of class OptionsSlider
' (VB) data of widget class OptionsSlider
Option Explicit

Public sliderWidget As WIDGET            ' std::shared_ptr<W_SLIDER>
Public cachedIdealSliderWidth As Long    ' = 0
Public text As String
Public wzText As WzCachedText
Public FontID As iV_fonts                ' = font_regular
Public cachedIdealTextWidth As Long      ' = 0
Public cachedIdealTextLineSize As Long   ' = 0
Public lastWidgetWidth As Long           ' = 0
Public outerPaddingY As Long             ' = 8
Public horizontalPadding As Long         ' = 10
Public isDisabled As Boolean             ' = false
Public isTruncated As Boolean            ' = false
Public currentValFunc As WzFunc          ' CurrentValueFunc: int32_t (void)
Public onChangeValueFunc As WzFunc       ' OnChangeValueFunc: void (int32_t newValue)
Public minValue As Long                  ' = 0
Public maxValue As Long                  ' = 0
Public totalRange As Long                ' uint32_t = 0
Public stepValue As Long                 ' = 1
Public needsStateUpdates As Boolean
