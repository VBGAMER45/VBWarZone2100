Attribute VB_Name = "bucket3d_c"
' src/bucket3d.cpp, bucket3d.h: stores object render calls in a list, renders
' them after sorting the objects by depth. A tag's void *pObject is the
' object's handle: a DROID *, STRUCTURE *, FEATURE * or PROJECTILE * (the
' pools'), a PROXIMITY_DISPLAY * (aProxDisp), a FLAG_POSITION * (aFlagPos), an
' EFFECT * (effects_c's aEffect) or an ATPART * (atmos_c's asAtmosParts index
' + 1). std::sort (not stable) is a merge sort here: tags of equal depth keep
' their order, which C++ leaves open.
Option Explicit

Public Enum RENDER_TYPE
    RENDER_DROID
    RENDER_STRUCTURE
    RENDER_FEATURE
    RENDER_PROXMSG
    RENDER_PROJECTILE
    RENDER_EFFECT
    RENDER_DELIVPOINT
    RENDER_PARTICLE
End Enum

' Gerard - HACK Multiplied by 7 to fix clipping
' someone needs to take a good look at the radius calculation
Private Const SCALE_DEPTH As Long = FP12_MULTIPLIER * 7

Private Type BUCKET_TAG
    objectType As RENDER_TYPE           ' type of object held
    pObject As Long                     ' pointer to the object (its handle)
    actualZ As Long                     ' int32_t
End Type

Private bucketArray() As BUCKET_TAG     ' static std::vector<BUCKET_TAG> bucketArray;
Private bucketArray_size As Long
Private bucketTmp() As BUCKET_TAG       ' (VB) the merge sort's buffer

' (VB) CLIP_LEFT, CLIP_RIGHT, CLIP_TOP, CLIP_BOTTOM: 0, pie_GetVideoBufferWidth(), 0, pie_GetVideoBufferHeight()
Private Function clippedOut(ByVal px As Long, ByVal py As Long, ByVal radius As Long) As Boolean
    clippedOut = (px + radius < 0) Or (px - radius > pie_GetVideoBufferWidth()) _
              Or (py + radius < 0) Or (py - radius > pie_GetVideoBufferHeight())
End Function

Private Function bucketCalculateZ(ByVal objectType As RENDER_TYPE, ByVal pObject As Long, perspectiveViewMatrix As mat4) As Long
    Dim z As Long, radius As Long, px As Long, py As Long, position As Vector3i, droidSize As Long
    Dim psBStats As Long, pImd As Long, typ As STRUCTURE_TYPE, psMessage As Long, vd As Long, psObj As Long
    Dim pVPM As Long
    pVPM = VarPtr(perspectiveViewMatrix)

    z = 0
    Select Case objectType
    Case RENDER_PARTICLE
        position.x = Fix(asAtmosParts(pObject - 1).position.x)
        position.y = Fix(asAtmosParts(pObject - 1).position.y)
        position.z = Fix(asAtmosParts(pObject - 1).position.z)

        position.x = position.x
        position.z = -(position.z)

        ' 16 below is HACK!!!
        z = pie_RotateProjectWithPerspective(position.x, position.y, position.z, pVPM, px, py) - 16
        If z > 0 Then
            'particle use the image radius
            radius = iIMDBaseShape_radius(asAtmosParts(pObject - 1).imd)
            radius = radius * SCALE_DEPTH
            radius = radius \ z
            If clippedOut(px, py, radius) Then
                z = -1
            End If
        End If

    Case RENDER_PROJECTILE
        Select Case asWeaponStats(aProj(pObject).psWStats).weaponSubClass
        Case WSC_FLAME, WSC_COMMAND, WSC_EMP
            ' We don't do projectiles from these guys, cos there's an effect instead
            z = -1
        Case Else
            'the weapon stats holds the reference to which graphic to use
            pImd = 0
            If asWeaponStats(aProj(pObject).psWStats).pInFlightGraphic <> 0 Then pImd = iIMDBaseShape_displayModel(asWeaponStats(aProj(pObject).psWStats).pInFlightGraphic)
            position.x = aProj(pObject).pos.x
            position.z = -(aProj(pObject).pos.y)
            position.y = aProj(pObject).pos.z

            z = pie_RotateProjectWithPerspective(position.x, position.y, position.z, pVPM, px, py)
            If z > 0 And pImd <> 0 Then
                'particle use the image radius
                radius = iIMDShape_radius(pImd)
                radius = radius * SCALE_DEPTH
                radius = radius \ z
                If clippedOut(px, py, radius) Then
                    z = -1
                End If
            End If
        End Select

    Case RENDER_STRUCTURE 'not depth sorted
        position.x = aStruct(pObject).pos.x
        position.z = -(aStruct(pObject).pos.y)
        typ = asStructureStats(aStruct(pObject).pStructureType).typ
        If (typ = REF_DEFENSE) Or (typ = REF_WALL) Or (typ = REF_WALLCORNER) Then
            position.y = aStruct(pObject).pos.z + 64
            radius = iIMDBaseShape_radius(aStruct(pObject).sDisplay.imd) 'walls guntowers and tank traps clip tightly
        Else
            position.y = aStruct(pObject).pos.z
            radius = iIMDBaseShape_radius(aStruct(pObject).sDisplay.imd)
        End If
        z = pie_RotateProjectWithPerspective(position.x, position.y, position.z, pVPM, px, py)
        If z > 0 Then
            'particle use the image radius
            radius = radius * SCALE_DEPTH
            radius = radius \ z
            If clippedOut(px, py, radius) Then
                z = -1
            End If
        End If

    Case RENDER_FEATURE 'not depth sorted
        position.x = aFeature(pObject).pos.x
        position.z = -(aFeature(pObject).pos.y)
        position.y = aFeature(pObject).pos.z + 2
        z = pie_RotateProjectWithPerspective(position.x, position.y, position.z, pVPM, px, py)
        If z > 0 Then
            'particle use the image radius
            radius = iIMDBaseShape_radius(aFeature(pObject).sDisplay.imd)
            radius = radius * SCALE_DEPTH
            radius = radius \ z
            If clippedOut(px, py, radius) Then
                z = -1
            End If
        End If

    Case RENDER_DROID
        position.x = aDroid(pObject).pos.x
        position.z = -(aDroid(pObject).pos.y)
        position.y = aDroid(pObject).pos.z
        psBStats = DROID_getBodyStats(pObject)
        droidSize = iIMDBaseShape_radius(asBodyStats(psBStats).pIMD)
        z = pie_RotateProjectWithPerspective(position.x, position.y, position.z, pVPM, px, py) - (droidSize * 2)
        If z > 0 Then
            'particle use the image radius
            radius = droidSize
            radius = radius * SCALE_DEPTH
            radius = radius \ z
            If clippedOut(px, py, radius) Then
                z = -1
            End If
        End If

    Case RENDER_PROXMSG
        psMessage = aProxDisp(pObject).psMessage
        If aProxDisp(pObject).typ = POS_PROXDATA Then
            vd = aMessage(psMessage).pViewData
            position.x = aViewData(vd).prox_x
            position.z = -(aViewData(vd).prox_y)
            position.y = aViewData(vd).prox_z
        ElseIf aProxDisp(pObject).typ = POS_PROXOBJ Then
            psObj = aMessage(psMessage).psObj
            position.x = SIMPLE_OBJECT_pos(psObj).x
            position.z = -(SIMPLE_OBJECT_pos(psObj).y)
            position.y = SIMPLE_OBJECT_pos(psObj).z
        End If
        z = pie_RotateProjectWithPerspective(position.x, position.y, position.z, pVPM, px, py)
        If z > 0 Then
            'particle use the image radius
            pImd = getDisplayImdFromIndex(MI_BLIP_ENEMY) 'use MI_BLIP_ENEMY as all are same radius
            radius = iIMDShape_radius(pImd)
            radius = radius * SCALE_DEPTH
            radius = radius \ z
            If clippedOut(px, py, radius) Then
                z = -1
            End If
        End If

    Case RENDER_EFFECT
        position.x = Fix(aEffect(pObject).position.x)
        position.z = Fix(-(aEffect(pObject).position.z))
        position.y = Fix(aEffect(pObject).position.y)

        ' 16 below is HACK!!!
        z = pie_RotateProjectWithPerspective(position.x, position.y, position.z, pVPM, px, py) - 16
        If z > 0 Then
            'particle use the image radius
            pImd = aEffect(pObject).imd
            If pImd <> 0 Then
                radius = iIMDShape_radius(pImd)
                radius = radius * SCALE_DEPTH
                radius = radius \ z
                If clippedOut(px, py, radius) Then
                    z = -1
                End If
            End If
        End If

    Case RENDER_DELIVPOINT
        position.x = aFlagPos(pObject).coords.x
        position.z = -(aFlagPos(pObject).coords.y)
        position.y = aFlagPos(pObject).coords.z
        z = pie_RotateProjectWithPerspective(position.x, position.y, position.z, pVPM, px, py)
        If z > 0 Then
            'particle use the image radius
            radius = iIMDBaseShape_radius(pAssemblyPointIMDs(aFlagPos(pObject).factoryType, aFlagPos(pObject).factoryInc))
            radius = radius * SCALE_DEPTH
            radius = radius \ z
            If clippedOut(px, py, radius) Then
                z = -1
            End If
        End If

    Case Else
    End Select

    bucketCalculateZ = z
End Function

' add an object to the current render list
Public Sub bucketAddTypeToList(ByVal objectType As RENDER_TYPE, ByVal pObject As Long, perspectiveViewMatrix As mat4)
    Dim pie As Long, newTag As BUCKET_TAG, z As Long

    z = bucketCalculateZ(objectType, pObject, perspectiveViewMatrix)

    If z < 0 Then
        ' Object will not be render - has been clipped!
        If objectType = RENDER_DROID Then
            ' Won't draw selection boxes
            aDroid(pObject).sDisplay.frameNumber = 0
        ElseIf objectType = RENDER_STRUCTURE Then
            aStruct(pObject).sDisplay.frameNumber = 0
        End If
        Exit Sub
    End If

    Select Case objectType
    Case RENDER_EFFECT
        Select Case aEffect(pObject).group
        Case EFFECT_EXPLOSION, EFFECT_CONSTRUCTION, EFFECT_SMOKE, EFFECT_FIREWORK
            ' Use calculated Z
        Case EFFECT_WAYPOINT
            pie = aEffect(pObject).imd
            z = INT32_MAX - iIMDShape_texpage(pie)
        Case Else
            z = INT32_MAX - 42
        End Select
    Case RENDER_DROID
        pie = iIMDBaseShape_displayModel(asBodyStats(DROID_getBodyStats(pObject)).pIMD)
        z = INT32_MAX - iIMDShape_texpage(pie)
    Case RENDER_STRUCTURE
        pie = iIMDBaseShape_displayModel(aStruct(pObject).sDisplay.imd)
        z = INT32_MAX - iIMDShape_texpage(pie)
    Case RENDER_FEATURE
        pie = iIMDBaseShape_displayModel(aFeature(pObject).sDisplay.imd)
        z = INT32_MAX - iIMDShape_texpage(pie)
    Case RENDER_DELIVPOINT
        pie = iIMDBaseShape_displayModel(pAssemblyPointIMDs(aFlagPos(pObject).factoryType, aFlagPos(pObject).factoryInc))
        z = INT32_MAX - iIMDShape_texpage(pie)
    Case RENDER_PARTICLE
        z = 0
    Case Else
        ' Use calculated Z
    End Select

    'put the object data into the tag
    newTag.objectType = objectType
    newTag.pObject = pObject
    newTag.actualZ = z

    'add tag to bucketArray
    If bucketArray_size = 0 Then
        If Not bucketArrayAllocated() Then ReDim bucketArray(0 To 255)
    ElseIf bucketArray_size > UBound(bucketArray) Then
        ReDim Preserve bucketArray(0 To 2 * bucketArray_size - 1)
    End If
    bucketArray(bucketArray_size) = newTag
    bucketArray_size = bucketArray_size + 1
End Sub

' (VB) whether bucketArray has storage
Private Function bucketArrayAllocated() As Boolean
    On Error GoTo none
    bucketArrayAllocated = UBound(bucketArray) >= 0
    Exit Function
none:
    bucketArrayAllocated = False
End Function

' (VB) std::sort(bucketArray) by BUCKET_TAG::operator< (actualZ > b.actualZ: reverse z order): a bottom-up merge sort
Private Sub sortBucketArray()
    Dim width As Long, lo As Long, md As Long, hi As Long, i As Long, j As Long, k As Long, n As Long
    n = bucketArray_size
    If n < 2 Then Exit Sub
    ReDim bucketTmp(0 To n - 1)
    width = 1
    Do While width < n
        lo = 0
        Do While lo < n
            md = lo + width
            If md > n Then md = n
            hi = lo + 2 * width
            If hi > n Then hi = n
            i = lo
            j = md
            k = lo
            Do While i < md And j < hi
                If bucketArray(j).actualZ > bucketArray(i).actualZ Then
                    bucketTmp(k) = bucketArray(j)
                    j = j + 1
                Else
                    bucketTmp(k) = bucketArray(i)
                    i = i + 1
                End If
                k = k + 1
            Loop
            Do While i < md
                bucketTmp(k) = bucketArray(i)
                i = i + 1
                k = k + 1
            Loop
            Do While j < hi
                bucketTmp(k) = bucketArray(j)
                j = j + 1
                k = k + 1
            Loop
            lo = hi
        Loop
        For i = 0 To n - 1
            bucketArray(i) = bucketTmp(i)
        Next
        width = width * 2
    Loop
End Sub

' render Objects in list
Public Sub bucketRenderCurrentList(viewMatrix As mat4, perspectiveViewMatrix As mat4)
    Dim i As Long
    sortBucketArray
    For i = 0 To bucketArray_size - 1
        Select Case bucketArray(i).objectType
        Case RENDER_PARTICLE
            renderParticle bucketArray(i).pObject, viewMatrix
        Case RENDER_EFFECT
            renderEffect bucketArray(i).pObject, viewMatrix
        Case RENDER_DROID
            displayComponentObject bucketArray(i).pObject, viewMatrix, perspectiveViewMatrix
        Case RENDER_STRUCTURE
            renderStructure bucketArray(i).pObject, viewMatrix, perspectiveViewMatrix
        Case RENDER_FEATURE
            renderFeature bucketArray(i).pObject, viewMatrix, perspectiveViewMatrix
        Case RENDER_PROXMSG
            renderProximityMsg bucketArray(i).pObject, viewMatrix, perspectiveViewMatrix
        Case RENDER_PROJECTILE
            renderProjectile bucketArray(i).pObject, viewMatrix, perspectiveViewMatrix
        Case RENDER_DELIVPOINT
            renderDeliveryPoint bucketArray(i).pObject, False, viewMatrix, perspectiveViewMatrix
        End Select
    Next
    'reset the bucket array as we go
    bucketArray_size = 0
End Sub
