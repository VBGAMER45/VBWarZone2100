Attribute VB_Name = "paged_entity_container_h"
' lib/framework/paged_entity_container.h: PagedEntityContainer, one copy
' per type used (DroidContainer, StructContainer, FeatureContainer, all
' with ReuseSlots = false; projectile.cpp's globalProjectileStorage,
' ProjectileContainer here, with ReuseSlots = true: a freed slot is taken
' again, the most recently freed first), and their storage: the object
' pools aDroid, aStruct, aFeature and aProj (see basedef_h for the handles). A slot is taken
' once, in order, and never reused until clear(), as in C++; the C++
' pages are one growing array. emplace is <Type>Container_alloc (a zeroed
' slot; the caller runs the constructor, GlobalDroidContainer_emplace and
' the like), erase runs the destructor and leaves the slot's fields as
' the destructor left them, as C++ leaves the dead object in its page.
Option Explicit

Private Const INITIAL_CAPACITY As Long = 256

Public aDroid() As DROID
Private droidUsed As Long               ' slots taken since the last clear() (C++: _maxIndex + 1)
Private droidCap As Long
Private droidAlive() As Boolean         ' SlotMetadata::is_alive
Private droidSize As Long
Private droidEmpty As DROID

Public aStruct() As STRUCTURE
Private structUsed As Long
Private structCap As Long
Private structAlive() As Boolean
Private structSize As Long
Private structEmpty As STRUCTURE
' (VB) scratch slots: a STRUCTURE a C++ function makes on the stack (structure.cpp's buildBlueprint)
' needs a slot here to be drawn; taken once per clear(), never alive, so no list or iteration sees it
Private structScratch(0 To 3) As Long

Public aFeature() As FEATURE
Private featureUsed As Long
Private featureCap As Long
Private featureAlive() As Boolean
Private featureSize As Long
Private featureEmpty As FEATURE

Public aProj() As PROJECTILE
Private projUsed As Long                ' slots ever taken since the last clear()
Private projCap As Long
Private projAlive() As Boolean
Private projSize As Long
Private projFree() As Long              ' the freed slots, a stack
Private projFreeCount As Long
Private projEmpty As PROJECTILE

' (VB) the first handle of a type's pool
Private Function poolBase(ByVal typ As OBJECT_TYPE) As Long
    poolBase = typ * OBJ_HANDLE_UNIT + 1
End Function

' (VB) the C++ initialisers of this module's globals (Sub Main calls it)
Public Sub paged_entity_container_h_init()
    droidCap = INITIAL_CAPACITY
    ReDim aDroid(poolBase(OBJ_DROID) To poolBase(OBJ_DROID) + droidCap - 1)
    ReDim droidAlive(poolBase(OBJ_DROID) To poolBase(OBJ_DROID) + droidCap - 1)
    structCap = INITIAL_CAPACITY
    ReDim aStruct(poolBase(OBJ_STRUCTURE) To poolBase(OBJ_STRUCTURE) + structCap - 1)
    ReDim structAlive(poolBase(OBJ_STRUCTURE) To poolBase(OBJ_STRUCTURE) + structCap - 1)
    featureCap = INITIAL_CAPACITY
    ReDim aFeature(poolBase(OBJ_FEATURE) To poolBase(OBJ_FEATURE) + featureCap - 1)
    ReDim featureAlive(poolBase(OBJ_FEATURE) To poolBase(OBJ_FEATURE) + featureCap - 1)
    projCap = INITIAL_CAPACITY
    ReDim aProj(poolBase(OBJ_PROJECTILE) To poolBase(OBJ_PROJECTILE) + projCap - 1)
    ReDim projAlive(poolBase(OBJ_PROJECTILE) To poolBase(OBJ_PROJECTILE) + projCap - 1)
    ReDim projFree(0 To projCap - 1)
End Sub

' ---- DroidContainer

' (VB) emplace's allocation: a zeroed live slot
Public Function DroidContainer_alloc() As Long
    Dim h As Long
    If droidUsed = droidCap Then
        droidCap = droidCap * 2
        ReDim Preserve aDroid(poolBase(OBJ_DROID) To poolBase(OBJ_DROID) + droidCap - 1)
        ReDim Preserve droidAlive(poolBase(OBJ_DROID) To poolBase(OBJ_DROID) + droidCap - 1)
    End If
    h = poolBase(OBJ_DROID) + droidUsed
    droidUsed = droidUsed + 1
    aDroid(h) = droidEmpty
    droidAlive(h) = True
    droidSize = droidSize + 1
    DroidContainer_alloc = h
End Function

' iterator find(T& x): whether the droid is alive in the container (C++: != end())
Public Function DroidContainer_find(ByVal h As Long) As Boolean
    If h < poolBase(OBJ_DROID) Or h >= poolBase(OBJ_DROID) + droidUsed Then Exit Function
    DroidContainer_find = droidAlive(h)
End Function

Public Sub DroidContainer_erase(ByVal h As Long)
    If Not droidAlive(h) Then Exit Sub
    droidAlive(h) = False
    DROID_destruct h
    droidSize = droidSize - 1
End Sub

Public Function DroidContainer_size() As Long
    DroidContainer_size = droidSize
End Function

' begin(): the first live droid, 0 = end()
Public Function DroidContainer_begin() As Long
    DroidContainer_begin = DroidContainer_next(poolBase(OBJ_DROID) - 1)
End Function

' ++it: the next live droid after h, 0 = end()
Public Function DroidContainer_next(ByVal h As Long) As Long
    For h = h + 1 To poolBase(OBJ_DROID) + droidUsed - 1
        If droidAlive(h) Then
            DroidContainer_next = h
            Exit Function
        End If
    Next
End Function

Public Sub DroidContainer_clear()
    Dim h As Long
    h = DroidContainer_begin()
    Do While h <> 0
        DroidContainer_erase h
        h = DroidContainer_next(h)
    Loop
    droidUsed = 0
End Sub

' ---- StructContainer

Public Function StructContainer_alloc() As Long
    Dim h As Long
    If structUsed = structCap Then
        structCap = structCap * 2
        ReDim Preserve aStruct(poolBase(OBJ_STRUCTURE) To poolBase(OBJ_STRUCTURE) + structCap - 1)
        ReDim Preserve structAlive(poolBase(OBJ_STRUCTURE) To poolBase(OBJ_STRUCTURE) + structCap - 1)
    End If
    h = poolBase(OBJ_STRUCTURE) + structUsed
    structUsed = structUsed + 1
    aStruct(h) = structEmpty
    structAlive(h) = True
    structSize = structSize + 1
    StructContainer_alloc = h
End Function

Public Function StructContainer_find(ByVal h As Long) As Boolean
    If h < poolBase(OBJ_STRUCTURE) Or h >= poolBase(OBJ_STRUCTURE) + structUsed Then Exit Function
    StructContainer_find = structAlive(h)
End Function

Public Sub StructContainer_erase(ByVal h As Long)
    If Not structAlive(h) Then Exit Sub
    structAlive(h) = False
    STRUCTURE_destruct h
    structSize = structSize - 1
End Sub

Public Function StructContainer_size() As Long
    StructContainer_size = structSize
End Function

Public Function StructContainer_begin() As Long
    StructContainer_begin = StructContainer_next(poolBase(OBJ_STRUCTURE) - 1)
End Function

Public Function StructContainer_next(ByVal h As Long) As Long
    For h = h + 1 To poolBase(OBJ_STRUCTURE) + structUsed - 1
        If structAlive(h) Then
            StructContainer_next = h
            Exit Function
        End If
    Next
End Function

Public Sub StructContainer_clear()
    Dim h As Long
    For h = 0 To 3
        structScratch(h) = 0
    Next
    h = StructContainer_begin()
    Do While h <> 0
        StructContainer_erase h
        h = StructContainer_next(h)
    Loop
    structUsed = 0
End Sub

' (VB) scratch slot n (0 to 3): zeroed, not alive, valid until the next clear()
Public Function StructContainer_scratch(ByVal n As Long) As Long
    If structScratch(n) = 0 Then
        structScratch(n) = StructContainer_alloc()
        structAlive(structScratch(n)) = False
        structSize = structSize - 1
    End If
    aStruct(structScratch(n)) = structEmpty
    StructContainer_scratch = structScratch(n)
End Function

' ---- FeatureContainer

Public Function FeatureContainer_alloc() As Long
    Dim h As Long
    If featureUsed = featureCap Then
        featureCap = featureCap * 2
        ReDim Preserve aFeature(poolBase(OBJ_FEATURE) To poolBase(OBJ_FEATURE) + featureCap - 1)
        ReDim Preserve featureAlive(poolBase(OBJ_FEATURE) To poolBase(OBJ_FEATURE) + featureCap - 1)
    End If
    h = poolBase(OBJ_FEATURE) + featureUsed
    featureUsed = featureUsed + 1
    aFeature(h) = featureEmpty
    featureAlive(h) = True
    featureSize = featureSize + 1
    FeatureContainer_alloc = h
End Function

Public Function FeatureContainer_find(ByVal h As Long) As Boolean
    If h < poolBase(OBJ_FEATURE) Or h >= poolBase(OBJ_FEATURE) + featureUsed Then Exit Function
    FeatureContainer_find = featureAlive(h)
End Function

Public Sub FeatureContainer_erase(ByVal h As Long)
    If Not featureAlive(h) Then Exit Sub
    featureAlive(h) = False
    FEATURE_destruct h
    featureSize = featureSize - 1
End Sub

Public Function FeatureContainer_size() As Long
    FeatureContainer_size = featureSize
End Function

Public Function FeatureContainer_begin() As Long
    FeatureContainer_begin = FeatureContainer_next(poolBase(OBJ_FEATURE) - 1)
End Function

Public Function FeatureContainer_next(ByVal h As Long) As Long
    For h = h + 1 To poolBase(OBJ_FEATURE) + featureUsed - 1
        If featureAlive(h) Then
            FeatureContainer_next = h
            Exit Function
        End If
    Next
End Function

Public Sub FeatureContainer_clear()
    Dim h As Long
    h = FeatureContainer_begin()
    Do While h <> 0
        FeatureContainer_erase h
        h = FeatureContainer_next(h)
    Loop
    featureUsed = 0
End Sub

' ---- ProjectileContainer (globalProjectileStorage)

' (VB) emplace's allocation: a zeroed live slot (a freed one if any; the caller runs PROJECTILE_construct)
Public Function ProjectileContainer_alloc() As Long
    Dim h As Long
    If projFreeCount > 0 Then
        projFreeCount = projFreeCount - 1
        h = projFree(projFreeCount)
    Else
        If projUsed = projCap Then
            projCap = projCap * 2
            ReDim Preserve aProj(poolBase(OBJ_PROJECTILE) To poolBase(OBJ_PROJECTILE) + projCap - 1)
            ReDim Preserve projAlive(poolBase(OBJ_PROJECTILE) To poolBase(OBJ_PROJECTILE) + projCap - 1)
            ReDim Preserve projFree(0 To projCap - 1)
        End If
        h = poolBase(OBJ_PROJECTILE) + projUsed
        projUsed = projUsed + 1
    End If
    aProj(h) = projEmpty
    projAlive(h) = True
    projSize = projSize + 1
    ProjectileContainer_alloc = h
End Function

' iterator find(T& x): whether the projectile is alive in the container (C++: != end())
Public Function ProjectileContainer_find(ByVal h As Long) As Boolean
    If h < poolBase(OBJ_PROJECTILE) Or h >= poolBase(OBJ_PROJECTILE) + projUsed Then Exit Function
    ProjectileContainer_find = projAlive(h)
End Function

Public Sub ProjectileContainer_erase(ByVal h As Long)
    If Not projAlive(h) Then Exit Sub
    projAlive(h) = False
    PROJECTILE_destruct h
    projSize = projSize - 1
    projFree(projFreeCount) = h
    projFreeCount = projFreeCount + 1
End Sub

Public Function ProjectileContainer_size() As Long
    ProjectileContainer_size = projSize
End Function

Public Sub ProjectileContainer_clear()
    Dim h As Long
    For h = poolBase(OBJ_PROJECTILE) To poolBase(OBJ_PROJECTILE) + projUsed - 1
        If projAlive(h) Then
            projAlive(h) = False
            PROJECTILE_destruct h
        End If
    Next
    projSize = 0
    projUsed = 0
    projFreeCount = 0
End Sub
