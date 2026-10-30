Attribute VB_Name = "macros_h"
' lib/framework/macros.h: MIN, MAX, ABSDIF and CLIP on signed integers
' (std::min and std::max on int too). MIN(a, b) is MIN_(a, b): the
' trailing underscore keeps them apart from locals named min and max.
' Also lib/framework/math_ext.h's clip(x, min, max) on int, as clip_.
Option Explicit

Public Function MIN_(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MIN_ = a Else MIN_ = b
End Function

Public Function MAX_(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MAX_ = a Else MAX_ = b
End Function

Public Function ABSDIF(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then ABSDIF = a - b Else ABSDIF = b - a
End Function

' CLIP(val, min, max)
Public Sub CLIP(val As Long, ByVal minVal As Long, ByVal maxVal As Long)
    If val < minVal Then
        val = minVal
    ElseIf val > maxVal Then
        val = maxVal
    End If
End Sub

' math_ext.h: Clips x to boundaries
Public Function clip_(ByVal x As Long, ByVal minVal As Long, ByVal maxVal As Long) As Long
    If x < minVal Then
        clip_ = minVal
    ElseIf x > maxVal Then
        clip_ = maxVal
    Else
        clip_ = x
    End If
End Function
