Attribute VB_Name = "pieclip_c"
' lib/ivis_opengl/pieclip.cpp
Option Explicit

Private videoBufferWidth As Long        ' UDWORD = 0
Private videoBufferHeight As Long       ' UDWORD = 0

Public Function pie_SetVideoBufferWidth(ByVal width As Long) As Boolean
    videoBufferWidth = width
    pie_SetVideoBufferWidth = (True)
End Function

Public Function pie_SetVideoBufferHeight(ByVal height As Long) As Boolean
    videoBufferHeight = height
    pie_SetVideoBufferHeight = (True)
End Function

Public Function pie_GetVideoBufferWidth() As Long
    pie_GetVideoBufferWidth = (videoBufferWidth)
End Function

Public Function pie_GetVideoBufferHeight() As Long
    pie_GetVideoBufferHeight = (videoBufferHeight)
End Function
