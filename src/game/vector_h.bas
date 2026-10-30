Attribute VB_Name = "vector_h"
' lib/framework/vector.h: the vector types' helpers (the types are src/shared/gfx's glm_h).
' Position (typedef Vector3i) is a Vector3i. Rotation's fields are
' uint16_t held in Longs: an assignment from a wider value masks with
' &HFFFF& where C++ would wrap.
Option Explicit

Public Type Rotation
    direction As Long                   ' uint16_t: Object rotation in 0..64k range
    pitch As Long                       ' uint16_t
    roll As Long                        ' uint16_t
End Type

' (VB) Vector2i(x, y)
Public Function Vector2i_(ByVal x As Long, ByVal y As Long) As Vector2i
    Vector2i_.x = x
    Vector2i_.y = y
End Function

' (VB) Vector3i(x, y, z), Position(x, y, z)
Public Function Vector3i_(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Vector3i
    Vector3i_.x = x
    Vector3i_.y = y
    Vector3i_.z = z
End Function

' (VB) Rotation(int direction, int pitch, int roll)
Public Function Rotation_(ByVal direction As Long, ByVal pitch As Long, ByVal roll As Long) As Rotation
    Rotation_.direction = direction And &HFFFF&
    Rotation_.pitch = pitch And &HFFFF&
    Rotation_.roll = roll And &HFFFF&
End Function

' dot(Vector2i, Vector2i): int arithmetic, which wraps (as on x86) where VB's would overflow
Public Function dot2i(a As Vector2i, b As Vector2i) As Long
    Dim d As Double
    d = CDbl(a.x) * b.x + CDbl(a.y) * b.y
    If d >= -2147483648# And d <= 2147483647# Then dot2i = d Else dot2i = FromU32(d)
End Function

' Round direction to nearest axis-aligned direction.
Public Function snapDirection(ByVal direction As Long) As Long
    snapDirection = ((direction And &HFFFF&) + &H2000&) And &HC000&
End Function
