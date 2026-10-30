VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' std::vector<Structure>, std::vector<Droid>, std::vector<Feature> (no C++
' counterpart of its own): the objects of one kind, at(i) from 0 to
' size() - 1.
Option Explicit

Private items() As Object
Private itemCount As Long
Private itemCap As Long

Public Function size() As Long
    size = itemCount
End Function

Public Function at(ByVal i As Long) As Object
    Set at = items(i)
End Function

Public Sub push_back(ByVal o As Object)
    If itemCount = itemCap Then
        itemCap = 2 * itemCap + 16
        ReDim Preserve items(0 To itemCap - 1)
    End If
    Set items(itemCount) = o
    itemCount = itemCount + 1
End Sub

Public Sub reserve(ByVal n As Long)
    If n <= itemCap Then Exit Sub
    itemCap = n
    ReDim Preserve items(0 To itemCap - 1)
End Sub

Public Sub clear()
    Dim i As Long
    For i = 0 To itemCount - 1
        Set items(i) = Nothing
    Next
    itemCount = 0
End Sub
