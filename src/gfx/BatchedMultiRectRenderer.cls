VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchedMultiRectRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/pieblitfunc.cpp/.h: class BatchedMultiRectRenderer (the
' exe's display3d draws its health and progress bars with one). The
' PIERECT_DrawRequest(_f) arguments are their fields; a projection matrix is
' an address (0 = the default argument, defaultProjectionMatrix());
' optional<size_t> rectGroup is has_rectGroup and rectGroup. The vectors are
' arrays and counts.
Option Explicit

Private Type GroupData                  ' std::vector<gfx_api::MultiRectPerInstanceInterleavedData>
    items() As MultiRectPerInstanceInterleavedData
    size As Long
End Type

Private Type GroupBufferInfo
    bufferOffset As Long                ' = 0
    instancesCount As Long              ' = 0
End Type

Private Type UploadedRectsInstanceBufferInfo
    buffer As gfx_api_buffer            ' = nullptr
    totalInstances As Long              ' = 0
    groupInfo() As GroupBufferInfo
    groupInfo_size As Long
End Type

Private useInstancedRendering As Boolean    ' = false
Private groupsData() As GroupData
Private groupsData_size As Long
Private totalAddedRects As Long         ' = 0

Private instancesData() As MultiRectPerInstanceInterleavedData
Private instancesData_size As Long
Private instanceDataBuffers() As gfx_api_buffer
Private instanceDataBuffers_size As Long
Private currInstanceBufferIdx As Long   ' = 0
Private has_currentUploadedRectInfo As Boolean
Private currentUploadedRectInfo As UploadedRectsInstanceBufferInfo

Public Sub resizeRectGroups(ByVal count As Long)
    Dim i As Long
    If count > 0 Then
        If groupsData_size = 0 Then
            ReDim groupsData(0 To count - 1)
        Else
            ReDim Preserve groupsData(0 To count - 1)
        End If
    End If
    For i = groupsData_size To count - 1
        groupsData(i).size = 0
    Next
    groupsData_size = count
End Sub

Private Sub pushRect(ByVal rectGroup As Long, data As MultiRectPerInstanceInterleavedData)
    Dim n As Long
    n = groupsData(rectGroup).size
    If n = 0 Then
        ReDim groupsData(rectGroup).items(0 To 15)
    ElseIf n > UBound(groupsData(rectGroup).items) Then
        ReDim Preserve groupsData(rectGroup).items(0 To 2 * n - 1)
    End If
    groupsData(rectGroup).items(n) = data
    groupsData(rectGroup).size = n + 1
End Sub

' rect: PIERECT_DrawRequest(x0, y0, x1, y1, color)
Public Sub addRect(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal color As Long, Optional ByVal rectGroup As Long = 0)
    Dim t As Long, center As Vector2f, matrix As mat4, data As MultiRectPerInstanceInterleavedData
    If x0 > x1 Then
        t = x0: x0 = x1: x1 = t
    End If
    If y0 > y1 Then
        t = y0: y0 = y1: y1 = t
    End If
    center = glm_vec2(x0, y0)
    matrix = glm_mat4_mul(glm_translate(glm_vec3(center.x, center.y, 0!)), glm_scale(glm_vec3(x1 - x0, y1 - y0, 1!)))

    data.TransformationMatrix = matrix
    data.offset_scale = glm_vec4(0!, 0!, 0!, 0!)
    data.colour = color                 ' rect.color.rgba()
    pushRect rectGroup, data
    totalAddedRects = totalAddedRects + 1
End Sub

' rect: PIERECT_DrawRequest_f(x0, y0, x1, y1, color)
Public Sub addRectF(ByVal x0 As Single, ByVal y0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal color As Long, Optional ByVal rectGroup As Long = 0)
    Dim t As Single, center As Vector2f, matrix As mat4, data As MultiRectPerInstanceInterleavedData
    If x0 > x1 Then
        t = x0: x0 = x1: x1 = t
    End If
    If y0 > y1 Then
        t = y0: y0 = y1: y1 = t
    End If
    center = glm_vec2(x0, y0)
    matrix = glm_mat4_mul(glm_translate(glm_vec3(center.x, center.y, 0!)), glm_scale(glm_vec3(x1 - x0, y1 - y0, 1!)))

    data.TransformationMatrix = matrix
    data.offset_scale = glm_vec4(0!, 0!, 0!, 0!)
    data.colour = color                 ' rect.color.rgba()
    pushRect rectGroup, data
    totalAddedRects = totalAddedRects + 1
End Sub

' projectionMatrix: an address (0 = defaultProjectionMatrix())
Public Sub drawAllRects(Optional ByVal projectionMatrix As Long = 0)
    Dim m As mat4
    m = defaultProjectionMatrix_or(projectionMatrix)
    drawRects_m False, 0, m
End Sub

Private Function uploadAllRectInstances() As UploadedRectsInstanceBufferInfo
    Dim result As UploadedRectsInstanceBufferInfo, i As Long, j As Long, startIdxOfGroupInstances As Long
    Dim one As MultiRectPerInstanceInterleavedData   ' (VB) for sizeof(gfx_api::MultiRectPerInstanceInterleavedData)

    instancesData_size = 0
    If totalAddedRects > 0 Then ReDim instancesData(0 To totalAddedRects - 1)   ' reserve
    If groupsData_size > 0 Then ReDim result.groupInfo(0 To groupsData_size - 1)
    For i = 0 To groupsData_size - 1
        startIdxOfGroupInstances = instancesData_size
        For j = 0 To groupsData(i).size - 1
            If instancesData_size > UBound(instancesData) Then ReDim Preserve instancesData(0 To 2 * instancesData_size - 1)
            instancesData(instancesData_size) = groupsData(i).items(j)
            instancesData_size = instancesData_size + 1
        Next
        result.groupInfo(result.groupInfo_size).bufferOffset = startIdxOfGroupInstances * LenB(one)
        result.groupInfo(result.groupInfo_size).instancesCount = groupsData(i).size
        result.groupInfo_size = result.groupInfo_size + 1
    Next

    If Not (totalAddedRects = instancesData_size) Then ASSERT_FAILURE "totalAddedRects == instancesData.size()", "lib/ivis_opengl/pieblitfunc.cpp:308", "uploadAllRectInstances", "totalAddedRects not up-to-date!"
    result.totalInstances = instancesData_size

    If result.totalInstances <> 0 Then
        If Not (instanceDataBuffers_size <> 0) Then
            ASSERT_FAILURE "!instanceDataBuffers.empty()", "lib/ivis_opengl/pieblitfunc.cpp:313", "uploadAllRectInstances", "No buffers available - unexpectedly called before init"
            Exit Function                   ' UploadedRectsInstanceBufferInfo()
        End If

        ' Upload buffer
        currInstanceBufferIdx = currInstanceBufferIdx + 1
        If currInstanceBufferIdx >= instanceDataBuffers_size Then
            currInstanceBufferIdx = 0
        End If
        instanceDataBuffers(currInstanceBufferIdx).upload instancesData_size * LenB(one), VarPtr(instancesData(0))
        Set result.buffer = instanceDataBuffers(currInstanceBufferIdx)
    Else
        Set result.buffer = Nothing
    End If

    uploadAllRectInstances = result
End Function

' rectGroup: optional<size_t> (has_rectGroup); projectionMatrix: an address (0 = defaultProjectionMatrix())
Public Sub drawRects(ByVal has_rectGroup As Boolean, ByVal rectGroup As Long, Optional ByVal projectionMatrix As Long = 0)
    Dim m As mat4
    m = defaultProjectionMatrix_or(projectionMatrix)
    drawRects_m has_rectGroup, rectGroup, m
End Sub

Private Sub drawRects_m(ByVal has_rectGroup As Boolean, ByVal rectGroup As Long, projectionMatrix As mat4)
    Dim i As Long, instanceBufferOffset As Long, instancesCount As Long
    Dim cb As constant_buffer_type_SHADER_RECT_INSTANCED
    Dim bufs(0 To 1) As gfx_api_buffer, offs(0 To 1) As Long
    If groupsData_size = 0 Then Exit Sub

    If Not useInstancedRendering Then
        If has_rectGroup Then i = rectGroup Else i = 0
        Do While i < groupsData_size
            If has_rectGroup And rectGroup <> i Then Exit Do
            If groupsData(i).size = 0 Then GoTo continue_
            pie_DrawMultiRect_NonInstanced groupsData(i).items, groupsData(i).size
continue_:
            i = i + 1
        Loop

        Exit Sub
    End If

    ' otherwise, use instanced rendering
    If Not has_currentUploadedRectInfo Then
        currentUploadedRectInfo = uploadAllRectInstances()
        has_currentUploadedRectInfo = True
    End If

    If currentUploadedRectInfo.totalInstances = 0 Then
        Exit Sub
    End If
    If Not (currentUploadedRectInfo.totalInstances = totalAddedRects) Then ASSERT_FAILURE "uploadedInfo.totalInstances == totalAddedRects", "lib/ivis_opengl/pieblitfunc.cpp:384", "drawRects", "Rects were added after instance data was uploaded - make sure to call clear() before adding rects after draw!": Exit Sub

    instanceBufferOffset = 0
    instancesCount = currentUploadedRectInfo.totalInstances

    If has_rectGroup Then
        instanceBufferOffset = currentUploadedRectInfo.groupInfo(rectGroup).bufferOffset
        instancesCount = currentUploadedRectInfo.groupInfo(rectGroup).instancesCount
    End If

    If instancesCount = 0 Then
        Exit Sub
    End If

    ' draw instances
    gfx_api_pso_get(BoxFillPSO_Instanced).bind
    cb.ProjectionMatrix = projectionMatrix
    gfx_api_pso_get(BoxFillPSO_Instanced).bind_constants VarPtr(cb), LenB(cb)
    Set bufs(0) = pie_internal_rectBuffer: offs(0) = 0
    Set bufs(1) = currentUploadedRectInfo.buffer: offs(1) = instanceBufferOffset
    gfx_api_context_bind_vertex_buffers 0, bufs, offs, 2
    gfx_api_pso_get(BoxFillPSO_Instanced).draw_instanced 4, 0, instancesCount
    gfx_api_pso_get(BoxFillPSO_Instanced).unbind_vertex_buffers pie_internal_rectBuffer, instanceDataBuffers(currInstanceBufferIdx)
End Sub

Public Function initialize() As Boolean
    Dim max_vertex_attribs As Long, max_vertex_output_components As Long
    Dim maxInstancedMeshShaderVertexAttribs As Long, i As Long
    Const min_required_vertex_output_components As Long = 11 * 4 ' from the shaders - see tcmask_instanced.vert / nolight_instanced.vert
    reset

    If Not gfx_api_context_supportsInstancedRendering() Then
        useInstancedRendering = False
        initialize = True
        Exit Function
    End If

    ' check for minimum required vertex attributes (and vertex outputs) for the instanced shaders
    max_vertex_attribs = gfx_api_context_get_context_value(context_value_MAX_VERTEX_ATTRIBS)
    max_vertex_output_components = gfx_api_context_get_context_value(context_value_MAX_VERTEX_OUTPUT_COMPONENTS)
    ' std::max({gfx_api::instance_modelMatrix, gfx_api::instance_packedValues, gfx_api::instance_Colour, gfx_api::instance_TeamColour}) + 1
    maxInstancedMeshShaderVertexAttribs = gfx_api_instance_modelMatrix
    If gfx_api_instance_packedValues > maxInstancedMeshShaderVertexAttribs Then maxInstancedMeshShaderVertexAttribs = gfx_api_instance_packedValues
    If gfx_api_instance_Colour > maxInstancedMeshShaderVertexAttribs Then maxInstancedMeshShaderVertexAttribs = gfx_api_instance_Colour
    If gfx_api_instance_TeamColour > maxInstancedMeshShaderVertexAttribs Then maxInstancedMeshShaderVertexAttribs = gfx_api_instance_TeamColour
    maxInstancedMeshShaderVertexAttribs = maxInstancedMeshShaderVertexAttribs + 1
    If max_vertex_attribs < maxInstancedMeshShaderVertexAttribs Then
        debug_ 427, LOG_INFO, "initialize", "Disabling instanced rendering - Insufficient MAX_VERTEX_ATTRIBS (%i; need: %zu)", max_vertex_attribs, maxInstancedMeshShaderVertexAttribs
        useInstancedRendering = False
        initialize = True
        Exit Function
    End If
    If max_vertex_output_components < min_required_vertex_output_components Then
        debug_ 433, LOG_INFO, "initialize", "Disabling instanced rendering - Insufficient MAX_VERTEX_OUTPUT_COMPONENTS (%i; need: %zu)", max_vertex_output_components, min_required_vertex_output_components
        useInstancedRendering = False
        initialize = True
        Exit Function
    End If

    instanceDataBuffers_size = gfx_api_context_maxFramesInFlight() + 1
    ReDim instanceDataBuffers(0 To instanceDataBuffers_size - 1)
    For i = 0 To instanceDataBuffers_size - 1
        Set instanceDataBuffers(i) = gfx_api_context_create_buffer_object(buffer_usage_vertex_buffer, buffer_storage_hint_stream_draw, "BatchedMultiRectRenderer::instanceDataBuffer[" & CStr(i) & "]")
    Next
    useInstancedRendering = True
    initialize = True
End Function

Public Sub clear()
    Dim i As Long
    For i = 0 To groupsData_size - 1
        groupsData(i).size = 0
    Next
    totalAddedRects = 0
    has_currentUploadedRectInfo = False     ' currentUploadedRectInfo.reset()
    Set currentUploadedRectInfo.buffer = Nothing
    instancesData_size = 0
End Sub

Public Sub reset()
    Dim numGroups As Long, i As Long
    clear
    numGroups = groupsData_size
    Erase groupsData
    groupsData_size = 0
    resizeRectGroups numGroups
    For i = 0 To instanceDataBuffers_size - 1
        Set instanceDataBuffers(i) = Nothing
    Next
    Erase instanceDataBuffers
    instanceDataBuffers_size = 0
End Sub
