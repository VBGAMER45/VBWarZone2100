Attribute VB_Name = "widgbase_h"
' lib/widget/widgbase.h's types and constants, lib/framework/geometry.h's
' WzRect, and (VB) the object model's helpers: weak references and WzFunc
' calls (docs/CONVENTIONS.md, "Widgets"). The classes are WIDGET.cls,
' W_SCREEN.cls and WzFunc.cls; the generated half is widget_classes_vb.bas
' (tools/gen_widgets.py).
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)

' lib/framework/geometry.h: class WzRect (x, y, width, height). Its methods are WzRect_<name>(r, ...).
Public Type WzRect
    x As Long
    y As Long
    w As Long
    h As Long
End Type

' The types of widget
Public Enum WIDGET_TYPE
    WIDG_FORM
    WIDG_LABEL
    WIDG_BUTTON
    WIDG_EDITBOX
    WIDG_BARGRAPH
    WIDG_SLIDER
    WIDG_UNSPECIFIED_TYPE
End Enum

' The keys that can be used to press a button
Public Enum WIDGET_KEY
    WKEY_NONE
    WKEY_PRIMARY
    WKEY_SECONDARY
End Enum

Public Const WIDG_HIDDEN As Long = &H8000&  ' The widget is initially hidden

' Possible states for a button or clickform.
Public Enum ButtonState
    WBUT_DISABLE = &H1                  ' Disable (grey out) a button.
    WBUT_LOCK = &H2                     ' Fix a button down.
    WBUT_CLICKLOCK = &H4                ' Fix a button down but it is still clickable.
    WBUT_FLASH = &H8                    ' Make a button flash.
    WBUT_DOWN = &H10                    ' Button is down.
    WBUT_HIGHLIGHT = &H20               ' Button is highlighted.
End Enum

' WIDGET::ChildZPos
Public Enum ChildZPos
    ChildZPos_Front
    ChildZPos_Back
End Enum

' lib/ivis_opengl/ivisdef.h: struct AtlasImage {IMAGEFILE *images; unsigned id;}: the image sheet's handle
' (0 = nullptr) and the image's index. Its methods are GfxApi's AtlasImage_width(a.images, a.id), _height,
' _xOffset, _yOffset; AtlasImage_isNull(a) is isNull().
Public Type AtlasImage
    images As Long
    id As Long
End Type

' Context information to pass into the widget functions (struct W_CONTEXT). W_CONTEXT::ZeroContext() is
' a zeroed W_CONTEXT; W_CONTEXT(xOffset, yOffset, mx, my) is W_CONTEXT_(...).
Public Type W_CONTEXT
    xOffset As Long                     ' Screen offset of the parent form
    yOffset As Long
    mx As Long                          ' mouse position on the form
    my As Long
End Type

Public Function WzRect_(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As WzRect
    WzRect_.x = x
    WzRect_.y = y
    WzRect_.w = width
    WzRect_.h = height
End Function

' AtlasImage(IMAGEFILE const *images = nullptr, unsigned id = 0)
Public Function AtlasImage_(Optional ByVal images As Long = 0, Optional ByVal id As Long = 0) As AtlasImage
    AtlasImage_.images = images
    AtlasImage_.id = id
End Function

Public Function AtlasImage_isNull(a As AtlasImage) As Boolean
    AtlasImage_isNull = (a.images = 0)
End Function

Public Function W_CONTEXT_(ByVal xOffset As Long, ByVal yOffset As Long, ByVal mx As Long, ByVal my As Long) As W_CONTEXT
    W_CONTEXT_.xOffset = xOffset
    W_CONTEXT_.yOffset = yOffset
    W_CONTEXT_.mx = mx
    W_CONTEXT_.my = my
End Function

' W_CONTEXT::convertToScreenContext()
Public Function W_CONTEXT_convertToScreenContext(c As W_CONTEXT) As W_CONTEXT
    W_CONTEXT_convertToScreenContext = c
    W_CONTEXT_convertToScreenContext.mx = c.mx + c.xOffset
    W_CONTEXT_convertToScreenContext.my = c.my + c.yOffset
    W_CONTEXT_convertToScreenContext.xOffset = 0
    W_CONTEXT_convertToScreenContext.yOffset = 0
End Function

'--------------------------------------------------------------------------
' (VB) weak references (std::weak_ptr): the object's address (ObjPtr), without a reference count. The
' object clears the weak references to it that it knows of when it dies (WIDGET's parent and screen
' pointers, W_SCREEN's focus and highlight: see WIDGET.cls); WIDGET_fromPtr(0) is Nothing.

Public Function WIDGET_fromPtr(ByVal p As Long) As WIDGET
    Dim o As WIDGET
    If p = 0 Then Exit Function
    CopyMemory o, p, 4                  ' a reference without AddRef...
    Set WIDGET_fromPtr = o              ' ...this one counts
    CopyMemory o, 0&, 4                 ' and the first goes without Release
End Function

Public Function W_SCREEN_fromPtr(ByVal p As Long) As W_SCREEN
    Dim o As W_SCREEN
    If p = 0 Then Exit Function
    CopyMemory o, p, 4
    Set W_SCREEN_fromPtr = o
    CopyMemory o, 0&, 4
End Function

'--------------------------------------------------------------------------
' (VB) WzFunc: a lambda, std::function or function pointer (docs/CONVENTIONS.md, "Widgets"): its id
' names the module procedure that runs it (a module's FN_ constant, dispatched by the generated
' WzFunc_dispatch to <module>_fn), its captures are cap(0..7).

' a new WzFunc with the id and up to 8 captured values (objects or values)
Public Function WzFunc_new(ByVal id As Long, ParamArray captures() As Variant) As WzFunc
    Dim f As WzFunc, i As Long
    Set f = New WzFunc
    f.id = id
    For i = 0 To UBound(captures)
        If IsObject(captures(i)) Then Set f.cap(i) = captures(i) Else f.cap(i) = captures(i)
    Next
    Set WzFunc_new = f
End Function

' calls f with the arguments (f Nothing: nothing happens, Empty is returned, as calling an empty
' std::function would be an error the C++ always checks for first)
Public Function WzFunc_call(ByVal f As WzFunc, ParamArray a() As Variant) As Variant
    Dim args() As Variant, i As Long
    If f Is Nothing Then Exit Function
    If UBound(a) >= 0 Then
        ReDim args(0 To UBound(a))
        For i = 0 To UBound(a)
            If IsObject(a(i)) Then Set args(i) = a(i) Else args(i) = a(i)
        Next
    Else
        ReDim args(0 To 0)
    End If
    WzFunc_assign WzFunc_call, WzFunc_dispatch(f, args)
End Function

' (assigns an object or a value)
Public Sub WzFunc_assign(dst As Variant, ByVal v As Variant)
    If IsObject(v) Then Set dst = v Else dst = v
End Sub
