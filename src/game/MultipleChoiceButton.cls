VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultipleChoiceButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class MultipleChoiceButton
' lib/widget/button.h: MultipleChoiceButton's fields (button_c); the initialisers are
' MultipleChoiceButton_construct's. tips is a Scripting.Dictionary (choice -> tip); imageSets
' (std::map<int, Images>) is imageSets_set / imageSets_find.
Option Explicit

Public choice As Long                   ' unsigned
Public tips As Object                   ' std::map<int, std::string>
Private imageSetKeys As Object          ' std::map<int, Images>: choice -> index into imageSetImages
Private imageSetImages() As W_BUTTON_Images
Private imageSetCount As Long

Private Sub Class_Initialize()
    Set tips = CreateObject("Scripting.Dictionary")
    Set imageSetKeys = CreateObject("Scripting.Dictionary")
End Sub

' imageSets[k] = images
Friend Sub imageSets_set(ByVal k As Long, images As W_BUTTON_Images)
    Dim i As Long
    If imageSetKeys.Exists(k) Then
        i = imageSetKeys(k)
    Else
        If imageSetCount = 0 Then
            ReDim imageSetImages(0 To 3)
        ElseIf imageSetCount > UBound(imageSetImages) Then
            ReDim Preserve imageSetImages(0 To 2 * imageSetCount - 1)
        End If
        i = imageSetCount
        imageSetCount = imageSetCount + 1
        imageSetKeys.Add k, i
    End If
    imageSetImages(i) = images
End Sub

' imageSets.find(k) != imageSets.end(): the images in images
Friend Function imageSets_find(ByVal k As Long, images As W_BUTTON_Images) As Boolean
    If Not imageSetKeys.Exists(k) Then Exit Function
    images = imageSetImages(imageSetKeys(k))
    imageSets_find = True
End Function
