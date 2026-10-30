Attribute VB_Name = "objlist_vb"
' (VB) std::list<T *> of object handles: the per-player object lists
' (apsDroidLists, apsStructLists, ...) and the other lists of pointers. An
' OBJ_LIST is a head, a tail and a size; its nodes live in one shared pool
' (lnObj, lnNext, lnPrev; 0 = end()), so a node index is a list iterator:
' it stays valid while other nodes are erased, as in std::list.
'
'   for (DROID *psDroid : apsDroidLists[player])     it = apsDroidLists(player).head
'                                                     Do While it <> 0
'                                                         psDroid = lnObj(it)
'                                                         ...
'                                                         it = lnNext(it)
'                                                     Loop
'
' it = list.erase(it) is it = objlist_erase(list, it).
Option Explicit

Public Type OBJ_LIST
    head As Long                        ' first node, 0 = empty
    tail As Long                        ' last node
    size As Long
End Type

Public lnObj() As Long                  ' the node's value (a handle)
Public lnNext() As Long                 ' the next node, 0 = end()
Public lnPrev() As Long                 ' the previous node, 0 = none
Private lnCap As Long
Private lnFree As Long                  ' free nodes, chained through lnNext

' (VB) a free node holding obj
Private Function newNode(ByVal obj As Long) As Long
    Dim n As Long, i As Long
    If lnFree = 0 Then
        n = lnCap
        If lnCap = 0 Then lnCap = 1024 Else lnCap = lnCap * 2
        ReDim Preserve lnObj(1 To lnCap)
        ReDim Preserve lnNext(1 To lnCap)
        ReDim Preserve lnPrev(1 To lnCap)
        For i = lnCap To n + 1 Step -1
            lnNext(i) = lnFree
            lnFree = i
        Next
    End If
    newNode = lnFree
    lnFree = lnNext(lnFree)
    lnObj(newNode) = obj
    lnNext(newNode) = 0
    lnPrev(newNode) = 0
End Function

' list.emplace_front(obj)
Public Sub objlist_push_front(l As OBJ_LIST, ByVal obj As Long)
    Dim n As Long
    n = newNode(obj)
    lnNext(n) = l.head
    If l.head <> 0 Then lnPrev(l.head) = n Else l.tail = n
    l.head = n
    l.size = l.size + 1
End Sub

' list.emplace_back(obj)
Public Sub objlist_push_back(l As OBJ_LIST, ByVal obj As Long)
    Dim n As Long
    n = newNode(obj)
    lnPrev(n) = l.tail
    If l.tail <> 0 Then lnNext(l.tail) = n Else l.head = n
    l.tail = n
    l.size = l.size + 1
End Sub

' it = list.erase(it): the node after it
Public Function objlist_erase(l As OBJ_LIST, ByVal it As Long) As Long
    Dim nx As Long, pv As Long
    nx = lnNext(it)
    pv = lnPrev(it)
    If pv <> 0 Then lnNext(pv) = nx Else l.head = nx
    If nx <> 0 Then lnPrev(nx) = pv Else l.tail = pv
    l.size = l.size - 1
    lnObj(it) = 0
    lnPrev(it) = 0
    lnNext(it) = lnFree
    lnFree = it
    objlist_erase = nx
End Function

' std::find(list.begin(), list.end(), obj): the node, 0 = end()
Public Function objlist_find(l As OBJ_LIST, ByVal obj As Long) As Long
    Dim it As Long
    it = l.head
    Do While it <> 0
        If lnObj(it) = obj Then
            objlist_find = it
            Exit Function
        End If
        it = lnNext(it)
    Loop
End Function

' list.clear()
Public Sub objlist_clear(l As OBJ_LIST)
    Do While l.head <> 0
        objlist_erase l, l.head
    Loop
End Sub

' list.front()
Public Function objlist_front(l As OBJ_LIST) As Long
    objlist_front = lnObj(l.head)
End Function
