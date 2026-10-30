Attribute VB_Name = "wzrect_h"
' lib/framework/geometry.h: class WzRect's methods, for widgbase_h's WzRect
' (x, y, w, h: C++ keeps the top left and bottom right corners, so a
' setX / setY moves the left / top edge only, as C++'s does). r.x() is r.x,
' r.width() is r.w; WzRect(topLeft, bottomRight) is WzRect_corners.
Option Explicit

' WzRect(const glm::ivec2 &topLeft, const glm::ivec2 &bottomRight)
Public Function WzRect_corners(ByVal left_ As Long, ByVal top As Long, ByVal right_ As Long, ByVal bottom As Long) As WzRect
    WzRect_corners.x = left_
    WzRect_corners.y = top
    WzRect_corners.w = right_ - left_
    WzRect_corners.h = bottom - top
End Function

Public Function WzRect_left(r As WzRect) As Long
    WzRect_left = r.x
End Function

Public Function WzRect_top(r As WzRect) As Long
    WzRect_top = r.y
End Function

Public Function WzRect_right(r As WzRect) As Long
    WzRect_right = r.x + r.w
End Function

Public Function WzRect_bottom(r As WzRect) As Long
    WzRect_bottom = r.y + r.h
End Function

' bool contains(int x, int y) const
Public Function WzRect_contains(r As WzRect, ByVal x As Long, ByVal y As Long) As Boolean
    WzRect_contains = Not (x < r.x) And (x < r.x + r.w) And Not (y < r.y) And (y < r.y + r.h)
End Function

' bool contains(const WzRect& other) const
Public Function WzRect_containsRect(r As WzRect, other As WzRect) As Boolean
    If Not WzRect_contains(r, other.x, other.y) Then Exit Function
    WzRect_containsRect = WzRect_contains(r, other.x + other.w - 1, other.y + other.h - 1)
End Function

Public Sub WzRect_setHeight(r As WzRect, ByVal height As Long)
    r.h = height
End Sub

Public Sub WzRect_setWidth(r As WzRect, ByVal width As Long)
    r.w = width
End Sub

' void setX(int x) { _topLeft.x = x; } (the right edge stays)
Public Sub WzRect_setX(r As WzRect, ByVal x As Long)
    r.w = r.w - (x - r.x)
    r.x = x
End Sub

Public Sub WzRect_setY(r As WzRect, ByVal y As Long)
    r.h = r.h - (y - r.y)
    r.y = y
End Sub

Public Sub WzRect_translateBy(r As WzRect, ByVal x As Long, ByVal y As Long)
    r.x = r.x + x
    r.y = r.y + y
End Sub

Public Function WzRect_intersects(r As WzRect, other As WzRect) As Boolean
    WzRect_intersects = (r.x < other.x + other.w) And (r.x + r.w > other.x) And (r.y < other.y + other.h) And (r.y + r.h > other.y)
End Function

Public Function WzRect_intersectionWith(r As WzRect, other As WzRect) As WzRect
    Dim l As Long, t As Long, rr As Long, b As Long
    If r.x > other.x Then l = r.x Else l = other.x
    If r.y > other.y Then t = r.y Else t = other.y
    If r.x + r.w < other.x + other.w Then rr = r.x + r.w Else rr = other.x + other.w
    If r.y + r.h < other.y + other.h Then b = r.y + r.h Else b = other.y + other.h
    WzRect_intersectionWith = WzRect_corners(l, t, rr, b)
End Function

Public Function WzRect_minimumBoundingRect(r As WzRect, other As WzRect) As WzRect
    Dim l As Long, t As Long, rr As Long, b As Long
    If r.x < other.x Then l = r.x Else l = other.x
    If r.y < other.y Then t = r.y Else t = other.y
    If r.x + r.w > other.x + other.w Then rr = r.x + r.w Else rr = other.x + other.w
    If r.y + r.h > other.y + other.h Then b = r.y + r.h Else b = other.y + other.h
    WzRect_minimumBoundingRect = WzRect_corners(l, t, rr, b)
End Function

' operator==
Public Function WzRect_equal(a As WzRect, b As WzRect) As Boolean
    WzRect_equal = (a.x = b.x And a.y = b.y And a.w = b.w And a.h = b.h)
End Function
