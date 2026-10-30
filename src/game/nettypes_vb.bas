Attribute VB_Name = "nettypes_vb"
' lib/netplay/nettypes.cpp's NETPosition, NETRotation and NETVector2i, for
' the exe's UDTs: Vector3i, Rotation and Vector2i are declared in each
' project, and UDTs do not cross the DLL boundary, so these overloads live
' in the exe and serialise the fields through WzNet's NETint32_t and
' NETuint16_t. Like WzNet's procedures, each takes the MessageReader or the
' MessageWriter As Object, the value ByRef (a reader writes it, a writer
' reads it).
Option Explicit

Public Sub NETPosition(ByVal c As Object, pos As Vector3i)
    NETint32_t c, pos.x
    NETint32_t c, pos.y
    NETint32_t c, pos.z
End Sub

Public Sub NETRotation(ByVal c As Object, rot As Rotation)
    NETuint16_t c, rot.direction
    NETuint16_t c, rot.pitch
    NETuint16_t c, rot.roll
End Sub

Public Sub NETVector2i(ByVal c As Object, vec As Vector2i)
    NETint32_t c, vec.x
    NETint32_t c, vec.y
End Sub
