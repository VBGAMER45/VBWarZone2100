Attribute VB_Name = "piepalette_h"
' lib/ivis_opengl/piepalette.h's static inline functions, the exe's copy (WzGfx's
' piepalette_c has its own; CONVENTIONS, "a static inline function of a header used
' by several projects"). A PIELIGHT is a Long: its 32-bit rgba, r in the low byte
' (the union's byte order). The palette itself is WzGfx's: WZCOL_X is psPalette(WZCOL_X).
Option Explicit

' (VB) PIELIGHT{r, g, b, a}
Public Function pal_RGBA(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
    pal_RGBA = (r And &HFF&) Or SHL(g And &HFF&, 8) Or SHL(b And &HFF&, 16) Or SHL(a And &HFF&, 24)
End Function

Public Function pal_Colour(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    pal_Colour = pal_RGBA(r, g, b, &HFF&)
End Function

Public Function pal_SetBrightness(ByVal brightness As Long) As Long
    pal_SetBrightness = pal_RGBA(brightness, brightness, brightness, &HFF&)
End Function
