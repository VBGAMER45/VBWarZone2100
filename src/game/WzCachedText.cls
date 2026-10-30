VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzCachedText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/paragraph.h: class WzCachedText, a WzText made when drawn and dropped when unused for
' cacheDurationMs. `New WzCachedText` is WzCachedText() (cacheDurationMs 100); the constructors with
' arguments are paragraph_h's WzCachedText_new / WzCachedText_new_text. operator->() (the WzText,
' made if needed, its expiry pushed back) is deref().
Option Explicit

Private text As String                  ' WzString
Private font As iV_fonts
Private has_cachedTextWidth As Boolean  ' optional<unsigned int> cachedTextWidth = nullopt
Private cachedTextWidth As Long
Private cacheDurationMs As Long         ' uint32_t
Private cachedText As WzText            ' std::unique_ptr<WzText> = nullptr
Private cacheExpireAt As Long           ' uint32_t = 0

' WzCachedText(uint32_t cacheDurationMs = 100): font(font_regular)
Private Sub Class_Initialize()
    font = font_regular
    cacheDurationMs = 100
End Sub

' (VB) the constructors' initialisers (paragraph_h)
Friend Sub init(ByVal text_ As String, ByVal font_ As iV_fonts, ByVal cacheDurationMs_ As Long)
    text = text_
    font = font_
    cacheDurationMs = cacheDurationMs_
End Sub

Public Sub tick()
    If Not cachedText Is Nothing Then
        If cacheExpireAt < realTime Then
            Set cachedText = Nothing
        End If
    End If
End Sub

Public Sub setText(ByVal text_ As String, ByVal fontID_ As iV_fonts)
    If text = text_ And font = fontID_ Then
        Exit Sub
    End If
    text = text_
    font = fontID_
    has_cachedTextWidth = False
    If Not cachedText Is Nothing Then
        cachedText.setText text, font
    End If
End Sub

Public Sub resetCachedDimensions()
    has_cachedTextWidth = False
End Sub

Public Function getText() As String
    getText = text
End Function

Public Function getFontID() As iV_fonts
    getFontID = font
End Function

Public Function getTextWidth() As Long
    If cachedText Is Nothing Then
        If Not has_cachedTextWidth Then
            cachedTextWidth = iV_GetTextWidth(text, font)
            has_cachedTextWidth = True
        End If
        getTextWidth = cachedTextWidth
        Exit Function
    End If
    getTextWidth = cachedText.width()
End Function

Public Function getTextLineSize() As Long
    If cachedText Is Nothing Then
        getTextLineSize = iV_GetTextLineSize(font)
        Exit Function
    End If
    getTextLineSize = cachedText.lineSize()
End Function

' WzText *operator ->()
Public Function deref() As WzText
    If cachedText Is Nothing Then
        Set cachedText = WzText_new(text, font)
    End If

    cacheExpireAt = realTime + (cacheDurationMs * GAME_TICKS_PER_SEC) \ 1000
    Set deref = cachedText
End Function
