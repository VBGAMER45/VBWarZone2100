VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParagraphElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/paragraph.cpp: struct ParagraphElement and its implementations ParagraphTextElement and
' ParagraphWidgetElement (paragraph_c): kind is PARAGRAPH_TEXT_ELEMENT or PARAGRAPH_WIDGET_ELEMENT, the
' fields of the other kind unused. The virtual methods are paragraph_c's ParagraphElement_<method>.
Option Explicit

Public kind As Long

' ParagraphTextElement
Public text As String                   ' WzString
Private m_style As ParagraphTextStyle
Public rightToLeft As Boolean           ' = false
Public fragments As Collection          ' std::vector<std::shared_ptr<ParagraphTextWidget>>

' ParagraphWidgetElement
Public widget As WIDGET
Public layoutWidth As Long              ' uint32_t = 0
Public layoutHeight As Long             ' uint32_t = 0
Public aboveBase As Long                ' int32_t

Friend Property Get style() As ParagraphTextStyle
    style = m_style
End Property

Friend Property Let style(v As ParagraphTextStyle)
    m_style = v
End Property
