VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchedImageDrawRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/ivis_opengl/pieblitfunc.cpp/.h: struct BatchedImageDrawRequests, the
' image draws iV_DrawImage and the repeat functions queue (the exe passes one
' to them; BatchedImageDrawRequests(bool deferRender) is
' BatchedImageDrawRequests_new). The queue is std::list<PieDrawImageRequest>:
' an array and a count. The queuePieImageDraw overloads take a UDT and a mat4:
' Friend methods, for pieblitfunc_c.
Option Explicit

Public deferRender As Boolean           ' = false

Private imageDrawRequests_() As PieDrawImageRequest
Private imageDrawRequests_size As Long

' void queuePieImageDraw(REND_MODE rendMode, IMAGEFILE *imageFile, int id, Vector2i size, const PIERECT& dest, PIELIGHT colour, const glm::mat4& modelViewProjection, Vector2i textureInset = Vector2i(0, 0))
Friend Sub queuePieImageDraw_args(ByVal rendMode As REND_MODE, ByVal imageFile_ As Long, ByVal id As Long, size As Vector2i, dest As PIERECT, ByVal colour As Long, modelViewProjection As mat4, textureInset As Vector2i)
    Dim request As PieDrawImageRequest
    request.rendMode = rendMode
    request.imageFile_ = imageFile_
    request.ID = id
    request.size = size
    request.dest = dest
    request.colour = colour
    request.modelViewProjection = modelViewProjection
    request.textureInset = textureInset
    queuePieImageDraw request
End Sub

Friend Sub queuePieImageDraw(request As PieDrawImageRequest)
    If imageDrawRequests_size = 0 Then
        ReDim imageDrawRequests_(0 To 15)
    ElseIf imageDrawRequests_size > UBound(imageDrawRequests_) Then
        ReDim Preserve imageDrawRequests_(0 To 2 * imageDrawRequests_size - 1)
    End If
    imageDrawRequests_(imageDrawRequests_size) = request
    imageDrawRequests_size = imageDrawRequests_size + 1
End Sub

Public Function draw(Optional ByVal force As Boolean = False) As Boolean
    If deferRender And Not force Then draw = False: Exit Function
    pie_DrawMultipleImages imageDrawRequests_, imageDrawRequests_size
    imageDrawRequests_size = 0
    draw = True
End Function

Public Sub clear()
    If Not (imageDrawRequests_size = 0) Then ASSERT_FAILURE "_imageDrawRequests.empty()", "lib/ivis_opengl/pieblitfunc.cpp:981", "clear", "Clearing a BatchedImageDrawRequests that isn't empty. Images have not been drawn!"
    imageDrawRequests_size = 0
End Sub
