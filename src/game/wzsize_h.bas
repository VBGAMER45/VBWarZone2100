Attribute VB_Name = "wzsize_h"
' lib/framework/geometry.h: class WzSize (a width and a height) and its
' methods, for the widgets (listwidget.h's TabSelectionStyle and ListWidget,
' dropdown.h's caret image). WzSize(w, h) is WzSize_(w, h).
Option Explicit

Public Type WzSize
    width_ As Long                      ' int _width
    height_ As Long                     ' int _height
End Type

Public Function WzSize_(ByVal width As Long, ByVal height As Long) As WzSize
    WzSize_.width_ = width
    WzSize_.height_ = height
End Function

' bool isEmpty() const { return _width <= 0 || _height <= 0; }
Public Function WzSize_isEmpty(s As WzSize) As Boolean
    WzSize_isEmpty = s.width_ <= 0 Or s.height_ <= 0
End Function
