Attribute VB_Name = "treap_c"
' lib/framework/treap.cpp: a treap of strings keyed by strings (strres's
' store). The nodes are a pool: a TREAP_NODE * is a 1-based index into
' treapNodes (0 = nullptr), freed nodes chain through psLeft. A
' TREAP_NODE ** (the treap) is a 1-based index into treapRoots (0 =
' nullptr). Keys compare with VB's binary compare (UTF-16 code units; C's
' strcmp on UTF-8 orders the same except for characters beyond U+FFFF).
' rand() is cstdlib_vb's.
Option Explicit

Private Type TREAP_NODE
    key As String                   ' The key to sort the node on
    priority As Long                ' Treap priority
    string_ As String               ' The string stored in the treap
    psLeft As Long                  ' The sub trees
    psRight As Long
End Type

Private treapNodes() As TREAP_NODE
Private treapNodesSize As Long
Private freeNodes As Long           ' first free node, chained through psLeft

Private treapRoots() As Long        ' *psTreap, for each treap
Private treapRootUsed() As Boolean
Private treapRootsSize As Long

' A useful comparison function - keys are char pointers
Private Function treapStringCmp(ByVal key1 As String, ByVal key2 As String) As Long
    Dim result As Long

    result = StrComp(key1, key2, vbBinaryCompare)
    If result < 0 Then
        treapStringCmp = -1
    ElseIf result > 0 Then
        treapStringCmp = 1
    Else
        treapStringCmp = 0
    End If
End Function

Public Function treapCreate() As Long
    Dim i As Long
    ' (VB) a free slot for the root
    For i = 1 To treapRootsSize
        If Not treapRootUsed(i) Then Exit For
    Next
    If i > treapRootsSize Then
        treapRootsSize = treapRootsSize + 1
        ReDim Preserve treapRoots(1 To treapRootsSize)
        ReDim Preserve treapRootUsed(1 To treapRootsSize)
    End If
    treapRootUsed(i) = True

    ' Initialise the tree to nothing
    treapRoots(i) = 0

    treapCreate = i
End Function

' Rotate a tree to the right
' (Make left sub tree the root and the root the right sub tree)
Private Sub treapRotRight(ppsRoot As Long)
    Dim psNewRoot As Long

    psNewRoot = treapNodes(ppsRoot).psLeft
    treapNodes(ppsRoot).psLeft = treapNodes(psNewRoot).psRight
    treapNodes(psNewRoot).psRight = ppsRoot
    ppsRoot = psNewRoot
End Sub

' Rotate a tree to the left
' (Make right sub tree the root and the root the left sub tree)
Private Sub treapRotLeft(ppsRoot As Long)
    Dim psNewRoot As Long

    psNewRoot = treapNodes(ppsRoot).psRight
    treapNodes(ppsRoot).psRight = treapNodes(psNewRoot).psLeft
    treapNodes(psNewRoot).psLeft = ppsRoot
    ppsRoot = psNewRoot
End Sub

' Recursive function to add an object to a tree
' (ppsRoot is a root or a node's psLeft or psRight, passed ByRef: nothing is
' allocated while it is held)
Private Sub treapAddNode(ppsRoot As Long, ByVal psNew As Long)
    If ppsRoot = 0 Then
        ' Make the node the root of the tree
        ppsRoot = psNew
    ElseIf treapStringCmp(treapNodes(psNew).key, treapNodes(ppsRoot).key) < 0 Then
        ' Node less than root, insert to the left of the tree
        treapAddNode treapNodes(ppsRoot).psLeft, psNew

        ' Sort the priority
        If treapNodes(ppsRoot).priority > treapNodes(treapNodes(ppsRoot).psLeft).priority Then
            ' rotate tree right
            treapRotRight ppsRoot
        End If
    Else
        ' Node greater than root, insert to the right of the tree
        treapAddNode treapNodes(ppsRoot).psRight, psNew

        ' Sort the priority
        If treapNodes(ppsRoot).priority > treapNodes(treapNodes(ppsRoot).psRight).priority Then
            ' rotate tree left
            treapRotLeft ppsRoot
        End If
    End If
End Sub

' Add an object to a treap
Public Function treapAdd(ByVal psTreap As Long, ByVal key As String, ByVal string_ As String) As Boolean
    Dim psNew As Long, i As Long

    ' (VB) a node from the pool (C: malloc, with the key and the string in the same chunk)
    If freeNodes = 0 Then
        ReDim Preserve treapNodes(1 To treapNodesSize * 2 + 16)
        For i = treapNodesSize + 1 To UBound(treapNodes)
            If i < UBound(treapNodes) Then treapNodes(i).psLeft = i + 1 Else treapNodes(i).psLeft = 0
        Next
        freeNodes = treapNodesSize + 1
        treapNodesSize = UBound(treapNodes)
    End If
    psNew = freeNodes
    freeNodes = treapNodes(psNew).psLeft

    treapNodes(psNew).key = key
    treapNodes(psNew).string_ = string_

    treapNodes(psNew).priority = rand()
    treapNodes(psNew).psLeft = 0
    treapNodes(psNew).psRight = 0

    treapAddNode treapRoots(psTreap), psNew

    treapAdd = True
End Function

' Recursively find an object in a treap (vbNullString: nullptr)
Private Function treapFindRec(ByVal psRoot As Long, ByVal key As String) As String
    If psRoot = 0 Then
        treapFindRec = vbNullString
        Exit Function
    End If

    Select Case treapStringCmp(key, treapNodes(psRoot).key)
    Case 0
        ' equal
        treapFindRec = treapNodes(psRoot).string_
        Exit Function
    Case -1
        treapFindRec = treapFindRec(treapNodes(psRoot).psLeft, key)
        Exit Function
    Case 1
        treapFindRec = treapFindRec(treapNodes(psRoot).psRight, key)
        Exit Function
    Case Else
        ASSERT_FAILURE "false", "lib/framework/treap.cpp:196", "treapFindRec", "treapFindRec: invalid return from comparison"
    End Select
    treapFindRec = vbNullString
End Function

' Find an object in a treap
Public Function treapFind(ByVal psTreap As Long, ByVal key As String) As String
    If Not (psTreap <> 0) Then ASSERT_FAILURE "psTreap != nullptr", "lib/framework/treap.cpp:206", "treapFind", "Invalid treap pointer!"

    treapFind = treapFindRec(treapRoots(psTreap), key)
End Function

Private Function treapFindKeyRec(ByVal psNode As Long, ByVal string_ As String) As String
    Dim key As String

    If psNode = 0 Then
        treapFindKeyRec = vbNullString
        Exit Function
    End If

    If treapNodes(psNode).string_ = string_ Then
        treapFindKeyRec = treapNodes(psNode).key
        Exit Function
    End If

    key = treapFindKeyRec(treapNodes(psNode).psLeft, string_)
    If StrPtr(key) <> 0 Then
        treapFindKeyRec = key
        Exit Function
    End If

    treapFindKeyRec = treapFindKeyRec(treapNodes(psNode).psRight, string_)
End Function

Public Function treapFindKey(ByVal psTreap As Long, ByVal string_ As String) As String
    If Not (psTreap <> 0) Then ASSERT_FAILURE "psTreap != nullptr", "lib/framework/treap.cpp:236", "treapFindKey", "Invalid treap pointer!"

    treapFindKey = treapFindKeyRec(treapRoots(psTreap), string_)
End Function

' Recursively free a treap
Private Sub treapDestroyRec(ByVal psRoot As Long)
    If psRoot = 0 Then
        Exit Sub
    End If

    ' free the sub branches
    treapDestroyRec treapNodes(psRoot).psLeft
    treapDestroyRec treapNodes(psRoot).psRight

    ' free the root
    treapNodes(psRoot).key = vbNullString
    treapNodes(psRoot).string_ = vbNullString
    treapNodes(psRoot).psRight = 0
    treapNodes(psRoot).psLeft = freeNodes
    freeNodes = psRoot
End Sub

' Destroy a treap and release all the memory associated with it
Public Sub treapDestroy(ByVal psTreap As Long)
    If Not (psTreap <> 0) Then ASSERT_FAILURE "psTreap != nullptr", "lib/framework/treap.cpp:261", "treapDestroy", "Invalid treap pointer!"

    treapDestroyRec treapRoots(psTreap)
    treapRoots(psTreap) = 0
    treapRootUsed(psTreap) = False
End Sub
