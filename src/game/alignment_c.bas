Attribute VB_Name = "alignment_c"
' lib/widget/alignment.cpp, alignment.h: struct Alignment (a UDT: Alignment(v, h, r)
' is Alignment_(v, h, r), Alignment::center() Alignment_center, wrap Alignment_wrap)
' and the widget class AlignmentWidget (docs/CONVENTIONS.md, "Widgets").
' (VB) widget class AlignmentWidget : WIDGET
Option Explicit

' enum class Alignment::Vertical
Public Enum Alignment_Vertical
    Alignment_Vertical_Top
    Alignment_Vertical_Center
    Alignment_Vertical_Bottom
End Enum

' enum class Alignment::Horizontal
Public Enum Alignment_Horizontal
    Alignment_Horizontal_Left
    Alignment_Horizontal_Center
    Alignment_Horizontal_Right
End Enum

' enum class Alignment::Resizing
Public Enum Alignment_Resizing
    Alignment_Resizing_Shrink
    Alignment_Resizing_Fit
End Enum

Public Type Alignment
    vertical As Alignment_Vertical
    horizontal As Alignment_Horizontal
    resizing As Alignment_Resizing
End Type

' Alignment(Vertical vertical, Horizontal horizontal, Resizing resizing = Resizing::Shrink)
Public Function Alignment_(ByVal vertical As Alignment_Vertical, ByVal horizontal As Alignment_Horizontal, Optional ByVal resizing As Alignment_Resizing = Alignment_Resizing_Shrink) As Alignment
    Alignment_.vertical = vertical
    Alignment_.horizontal = horizontal
    Alignment_.resizing = resizing
End Function

' static Alignment center()
Public Function Alignment_center() As Alignment
    Alignment_center = Alignment_(Alignment_Vertical_Center, Alignment_Horizontal_Center)
End Function

' std::shared_ptr<AlignmentWidget> wrap(std::shared_ptr<WIDGET> widget)
Public Function Alignment_wrap(a As Alignment, ByVal widget As WIDGET) As WIDGET
    Dim wrap As WIDGET
    Set wrap = AlignmentWidget_new(a)
    wrap.attach widget
    Set Alignment_wrap = wrap
End Function

' explicit AlignmentWidget(Alignment alignment)
Public Function AlignmentWidget_new(align_ As Alignment) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_AlignmentWidget)
    AlignmentWidget_construct w, align_
    Set AlignmentWidget_new = w
End Function

Public Sub AlignmentWidget_construct(ByVal self As WIDGET, align_ As Alignment)
    WIDGET_construct self
    self.kind = K_AlignmentWidget ' (VB) virtual calls from here on reach AlignmentWidget's overrides
    AlignmentWidget_(self).alignment_ = align_
    self.setTransparentToMouse True
End Sub

Public Sub AlignmentWidget_geometryChanged(ByVal self As WIDGET)
    Dim child As WIDGET, targetWidth As Long, targetHeight As Long, fitRatio As Single, childRect As WzRect
    Dim align As Alignment
    If self.childCount() = 0 Then
        Exit Sub
    End If

    align = AlignmentWidget_(self).alignment_
    Set child = self.childAt(0)
    targetWidth = 0
    targetHeight = 0
    If align.resizing = Alignment_Resizing_Shrink Then
        targetWidth = MIN_(self.width(), child.idealWidth())
        targetHeight = MIN_(self.height(), child.idealHeight())
    ElseIf align.resizing = Alignment_Resizing_Fit Then
        fitRatio = fdiv(CSng(self.width()), CSng(child.idealWidth()))
        If fdiv(CSng(self.height()), CSng(child.idealHeight())) < fitRatio Then fitRatio = fdiv(CSng(self.height()), CSng(child.idealHeight()))
        targetWidth = ToInt(fmul(CSng(child.idealWidth()), fitRatio))
        targetHeight = ToInt(fmul(CSng(child.idealHeight()), fitRatio))
    End If

    childRect = WzRect_(0, 0, 0, 0)
    If align.horizontal = Alignment_Horizontal_Center Then
        WzRect_setX childRect, MAX_(0, self.width() - targetWidth) \ 2
    ElseIf align.horizontal = Alignment_Horizontal_Right Then
        WzRect_setX childRect, MAX_(0, self.width() - targetWidth)
    End If

    If align.vertical = Alignment_Vertical_Center Then
        WzRect_setY childRect, MAX_(0, self.height() - targetHeight) \ 2
    ElseIf align.vertical = Alignment_Vertical_Bottom Then
        WzRect_setY childRect, MAX_(0, self.height() - targetHeight)
    End If

    ' must set width/height **after** setX/setY
    WzRect_setWidth childRect, targetWidth
    WzRect_setHeight childRect, targetHeight

    child.setGeometryRect childRect
End Sub

Public Function AlignmentWidget_idealWidth(ByVal self As WIDGET) As Long
    If self.childCount() = 0 Then
        AlignmentWidget_idealWidth = WIDGET_idealWidth(self)
    Else
        AlignmentWidget_idealWidth = self.childAt(0).idealWidth()
    End If
End Function

Public Function AlignmentWidget_idealHeight(ByVal self As WIDGET) As Long
    If self.childCount() = 0 Then
        AlignmentWidget_idealHeight = WIDGET_idealHeight(self)
    Else
        AlignmentWidget_idealHeight = self.childAt(0).idealHeight()
    End If
End Function
