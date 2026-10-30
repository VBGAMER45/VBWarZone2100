VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pipeline_state_helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/gfx_api.h: gfx_api::pipeline_state_helper<rasterizer,
' primitive, index, uniform_inputs, vertex_buffer_inputs, texture_inputs,
' shader>. One class for every instantiation: the template arguments are the
' pso_type it was made for, written out by gfx_api_pso_c's pso_create_info;
' gfx_api::X::get() is gfx_api_pso_get(X) (one object per pso_type).
' Arguments that C++ passes by reference are their addresses:
' bind_constants(VarPtr(cb), LenB(cb)); set_uniforms(VarPtr(a), LenB(a),
' VarPtr(b), LenB(b), ...) takes the blocks as address and size pairs. The
' buffers and textures are gfx_api_buffer / gfx_api_texture / ... objects
' (Nothing = nullptr).
Option Explicit

Private psoType As pso_type             ' (VB) which instantiation
Private createInfo As pipeline_create_info ' (VB) its template arguments
Private pso As gfx_api_pipeline_state_object ' = nullptr
Private nextpso As gfx_api_pipeline_state_object ' = nullptr

' pipeline_state_helper() (private constructor): recompile()
Friend Sub init(ByVal pso_ As pso_type)
    psoType = pso_
    createInfo = pso_create_info(pso_)
    recompile
End Sub

Public Sub bind()
    If Not nextpso Is Nothing Then
        If Not nextpso.broken Then
            Set pso = nextpso               ' if (this->pso != nullptr) delete this->pso;
            Set nextpso = Nothing
        End If
    End If
    gfx_api_context_bind_pipeline pso, createInfo.texture_desc_size = 0
End Sub

' (the static_assert on the number of vertex buffers is the caller's)
Public Sub bind_vertex_buffers(ParamArray args() As Variant)
    Dim buffers() As gfx_api_buffer, offsets() As Long, n As Long, i As Long
    n = UBound(args) + 1
    If n > 0 Then
        ReDim buffers(0 To n - 1)
        ReDim offsets(0 To n - 1)
        For i = 0 To n - 1
            Set buffers(i) = args(i)
        Next i
    End If
    gfx_api_context_bind_vertex_buffers 0, buffers, offsets, n
End Sub

Public Sub unbind_vertex_buffers(ParamArray args() As Variant)
    Dim buffers() As gfx_api_buffer, offsets() As Long, n As Long, i As Long
    n = UBound(args) + 1
    If n > 0 Then
        ReDim buffers(0 To n - 1)
        ReDim offsets(0 To n - 1)
        For i = 0 To n - 1
            Set buffers(i) = args(i)
        Next i
    End If
    gfx_api_context_unbind_vertex_buffers 0, buffers, offsets, n
End Sub

Public Sub bind_textures(ParamArray args() As Variant)
    Dim textures() As gfx_api_abstract_texture, n As Long, i As Long
    n = UBound(args) + 1
    If n > 0 Then
        ReDim textures(0 To n - 1)
        For i = 0 To n - 1
            Set textures(i) = args(i)
        Next i
    End If
    gfx_api_context_bind_textures createInfo.texture_desc, createInfo.texture_desc_size, textures, n
End Sub

' data: the address of the constant_buffer_type<shader> (a gfx_api_h UDT); size: its LenB
Public Sub bind_constants(ByVal data As Long, ByVal size As Long)
    ' Vulkan: Many platforms have a maxUniformBufferRange of 64k
    ' - see: https://vulkan.gpuinfo.org/displaydevicelimit.php?name=maxUniformBufferRange
    ' (static_assert(sizeof(constant_buffer_type<shader>) <= 65536, "Constant buffer size exceeds 64k"))
    gfx_api_context_set_constants data, size
End Sub

' args: address, size, address, size, ... (one pair per uniform input)
Public Sub set_uniforms(ParamArray args() As Variant)
    Dim a As Variant
    a = args
    set_uniforms_list 0, a
End Sub

Public Sub set_uniforms_at(ByVal first As Long, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    set_uniforms_list first, a
End Sub

Private Sub set_uniforms_list(ByVal first As Long, args As Variant)
    Dim blocks() As Long, sizes() As Long, n As Long, i As Long
    n = (UBound(args) + 1) \ 2
    If n > 0 Then
        ReDim blocks(0 To n - 1)
        ReDim sizes(0 To n - 1)
        For i = 0 To n - 1
            blocks(i) = args(2 * i)
            sizes(i) = args(2 * i + 1)
        Next i
    End If
    gfx_api_context_set_uniforms first, blocks, sizes, n
End Sub

Public Sub draw(ByVal count As Long, ByVal offset As Long)
    gfx_api_context_draw offset, count, createInfo.primitive
End Sub

Public Sub draw_elements(ByVal count As Long, ByVal offset As Long)
    gfx_api_context_draw_elements offset, count, createInfo.primitive, createInfo.index
End Sub

Public Sub draw_instanced(ByVal count As Long, ByVal offset As Long, ByVal instance_count As Long)
    gfx_api_context_draw_instanced offset, count, createInfo.primitive, instance_count
End Sub

Public Sub draw_elements_instanced(ByVal count As Long, ByVal offset As Long, ByVal instance_count As Long)
    gfx_api_context_draw_elements_instanced offset, count, createInfo.primitive, createInfo.index, instance_count
End Sub

Public Function recompile() As Boolean
    Set nextpso = gfx_api_context_build_pipeline(pso, createInfo)
    If Not nextpso Is Nothing Then recompile = Not nextpso.broken
End Function

Public Function isBroken() As Boolean
    If Not nextpso Is Nothing Then
        isBroken = nextpso.broken
    Else
        isBroken = pso.broken
    End If
End Function
