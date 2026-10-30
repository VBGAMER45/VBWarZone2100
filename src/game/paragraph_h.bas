Attribute VB_Name = "paragraph_h"
' lib/widget/paragraph.h's class WzCachedText: its constructors with
' arguments (the class: WzCachedText.cls; paragraph.cpp's Paragraph is
' another module's).
Option Explicit

' WzCachedText(uint32_t cacheDurationMs = 100)
Public Function WzCachedText_new(Optional ByVal cacheDurationMs As Long = 100) As WzCachedText
    Dim t As WzCachedText
    Set t = New WzCachedText
    t.init "", font_regular, cacheDurationMs
    Set WzCachedText_new = t
End Function

' WzCachedText(WzString text, iV_fonts font, uint32_t cacheDurationMs = 100)
Public Function WzCachedText_new_text(ByVal text As String, ByVal font As iV_fonts, Optional ByVal cacheDurationMs As Long = 100) As WzCachedText
    Dim t As WzCachedText
    Set t = New WzCachedText
    t.init text, font, cacheDurationMs
    Set WzCachedText_new_text = t
End Function
