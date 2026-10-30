Attribute VB_Name = "infobutton_c"
' src/titleui/widgets/infobutton.cpp, infobutton.h: class WzInfoButton, a
' WzFrontendImageButton with the info circle image. WzInfoButton::make is
' WzInfoButton_make.
' (VB) widget class WzInfoButton : WzFrontendImageButton
Option Explicit

' WzInfoButton::WzInfoButton()
Public Sub WzInfoButton_construct(ByVal self As WIDGET)
    WzFrontendImageButton_construct self
    self.kind = K_WzInfoButton ' (VB) virtual calls from here on reach WzInfoButton's overrides
End Sub

Public Function WzInfoButton_make() As WIDGET
    Dim widget As WIDGET
    Set widget = WIDGET_alloc(K_WzInfoButton)
    WzInfoButton_construct widget
    WzFrontendImageButton_setImage widget, True, IMAGE_INFO_CIRCLE
    WzFrontendImageButton_setBorderDrawMode widget, BorderDrawMode_Highlighted
    WzFrontendImageButton_setPadding widget, 0, 0
    Set WzInfoButton_make = widget
End Function
