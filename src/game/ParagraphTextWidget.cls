VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParagraphTextWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class ParagraphTextWidget
' lib/widget/paragraph.cpp: ParagraphTextWidget's fields (paragraph_c).
Option Explicit

Public cachedText As WzCachedText
Private m_textStyle As ParagraphTextStyle

Friend Property Get textStyle() As ParagraphTextStyle
    textStyle = m_textStyle
End Property

Friend Property Let textStyle(v As ParagraphTextStyle)
    m_textStyle = v
End Property
