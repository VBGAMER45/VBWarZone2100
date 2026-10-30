Attribute VB_Name = "glm_h"
' (VB) glm's types, as lib/framework/vector.h names them (Vector2i = glm::ivec2,
' Vector2f = glm::vec2, Vector3i = glm::ivec3, Vector3f = glm::vec3) and as this
' port names the others (Vector4f = glm::vec4, Vector4i = glm::ivec4, mat4 =
' glm::mat4). Compiled into WzGfx and the exe (src/shared/gfx), which pass them
' to each other by address (VarPtr) inside the constant buffers and vertex data
' of gfx_api.h (gfx_api_h). glm's functions on them are glm_mat_vb's; the
' float-exact vector operations of the PIE loader are WzGfx's glm_vb.
Option Explicit

Public Type Vector2i
    x As Long
    y As Long
End Type

Public Type Vector2f
    x As Single
    y As Single
End Type

Public Type Vector3i
    x As Long
    y As Long
    z As Long
End Type

Public Type Vector3f
    x As Single
    y As Single
    z As Single
End Type

Public Type Vector4f
    x As Single
    y As Single
    z As Single
    w As Single
End Type

Public Type Vector4i
    x As Long
    y As Long
    z As Long
    w As Long
End Type

' glm::mat4: column-major as glm stores it (value_ptr order): glm's m[col][row]
' is m((col) * 4 + (row))
Public Type mat4
    m(0 To 15) As Single
End Type
