VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_SCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/widgbase.h: struct W_SCREEN, the screen structure which stores all info for a widget
' screen; its methods are widget.cpp's. W_SCREEN::make(customRootForm) is widget_c's W_SCREEN_make.
' psFocus and lastHighlight are weak (std::weak_ptr<WIDGET>): addresses, read through the property
' (its lock; Nothing if none), cleared by the widget when it dies (forgetWidget). retWidgets (the
' WidgetTriggers of widgRunScreen) is retWidgetCount() / retWidgetAt(i).
Option Explicit

Public psForm As WIDGET                 ' std::shared_ptr<W_FORM>: The root form of the screen
Private psFocus_ As Long                ' std::weak_ptr<WIDGET>: The widget that has keyboard focus
Private lastHighlight_ As Long          ' std::weak_ptr<WIDGET>: The last widget to be highlighted. This is used to track when the mouse moves off something.
Public TipFontID As iV_fonts            ' ID of the IVIS font to use for tool tips.
Private retWidgets() As WIDGET          ' WidgetTriggers: The widgets to be returned by widgRunScreen.
Private retCount As Long

' W_SCREEN(): TipFontID(font_regular)
Private Sub Class_Initialize()
    TipFontID = font_regular
End Sub

' virtual ~W_SCREEN() (VB: first the widgets' weak screen pointers expire, as C++'s have when the
' destructor runs; then psFocus's focusLost)
Private Sub Class_Terminate()
    Dim focusedWidget As WIDGET
    If Not psForm Is Nothing Then psForm.clearScreenPointerRecursive
    Set focusedWidget = WIDGET_fromPtr(psFocus_)
    If Not focusedWidget Is Nothing Then
        ' must trigger a resignation of focus
        If widgIsInitialized() Then ' do not call focusLost if widgets have already shutdown / are not initialized
            focusedWidget.focusLost
        End If
    End If
    Set focusedWidget = Nothing
    psFocus_ = 0
End Sub

' void initialize(const std::shared_ptr<W_FORM>& customRootForm)
Friend Sub initialize(ByVal customRootForm As WIDGET)
    Dim customRootFormParent As WIDGET, sInit As W_FORMINIT
    If Not customRootForm Is Nothing Then
        Set customRootFormParent = customRootForm.parent()
        If Not customRootFormParent Is Nothing Then
            customRootFormParent.detach customRootForm
        End If
        Set psForm = customRootForm
        psForm.setScreenPointer Me
        Exit Sub
    End If

    ' W_FORMINIT sInit;
    sInit.id = 0
    sInit.style = WFORM_PLAIN Or WFORM_INVISIBLE
    sInit.x = 0
    sInit.y = 0
    sInit.width = screenWidth - 1
    sInit.height = screenHeight - 1

    Set psForm = W_FORM_new_init(sInit)
    psForm.assignScreenPointer Me
End Sub

' Sets psFocus, notifying the old widget, if any.
Public Sub setFocus(ByVal widget As WIDGET)
    Dim locked As WIDGET
    Set locked = WIDGET_fromPtr(psFocus_)
    If Not locked Is Nothing Then
        If locked Is widget Then
            Exit Sub ' do nothing - no change
        End If
        psFocus_ = ObjPtr(widget)
        If widgIsInitialized() Then ' do not call focusLost if widgets have already shutdown / are not initialized
            locked.focusLost
        End If
    Else
        psFocus_ = ObjPtr(widget)
    End If
End Sub

' Adds psWidget to retWidgets.
Public Sub setReturn(ByVal psWidget As WIDGET)
    ' WidgetTrigger trigger; trigger.widget = psWidget; retWidgets.push_back(trigger);
    If retCount = 0 Then
        ReDim retWidgets(0 To 3)
    ElseIf retCount > UBound(retWidgets) Then
        ReDim Preserve retWidgets(0 To 2 * retCount - 1)
    End If
    Set retWidgets(retCount) = psWidget
    retCount = retCount + 1
End Sub

' used to handle screen resizing
Public Sub screenSizeDidChange(ByVal oldWidth As Long, ByVal oldHeight As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ' resize the top-level form
    psForm.setGeometry 0, 0, screenWidth, screenHeight

    ' inform the top-level form of the event
    psForm.screenSizeDidChange oldWidth, oldHeight, newWidth, newHeight
End Sub

Public Function getWidgetWithFocus() As WIDGET
    Set getWidgetWithFocus = WIDGET_fromPtr(psFocus_)
End Function

Public Function hasFocus(ByVal widget As WIDGET) As Boolean
    hasFocus = (psFocus_ <> 0 And psFocus_ = ObjPtr(widget))
End Function

Public Function isLastHighlight(ByVal widget As WIDGET) As Boolean
    isLastHighlight = (lastHighlight_ <> 0 And lastHighlight_ = ObjPtr(widget))
End Function

'--------------------------------------------------------------------------
' (VB) the weak members and retWidgets

' psFocus.lock()
Public Property Get psFocus() As WIDGET
    Set psFocus = WIDGET_fromPtr(psFocus_)
End Property

' psFocus = widget (no focusLost: setFocus notifies)
Public Property Set psFocus(ByVal widget As WIDGET)
    psFocus_ = ObjPtr(widget)
End Property

' psFocus.reset()
Public Sub psFocusReset()
    psFocus_ = 0
End Sub

' lastHighlight.lock()
Public Property Get lastHighlight() As WIDGET
    Set lastHighlight = WIDGET_fromPtr(lastHighlight_)
End Property

' lastHighlight = widget
Public Sub setLastHighlight(ByVal widget As WIDGET)
    lastHighlight_ = ObjPtr(widget)
End Sub

' lastHighlight.reset()
Public Sub lastHighlightReset()
    lastHighlight_ = 0
End Sub

' the widget at address p dies: the weak pointers to it expire
Friend Sub forgetWidget(ByVal p As Long)
    If psFocus_ = p Then psFocus_ = 0
    If lastHighlight_ = p Then lastHighlight_ = 0
End Sub

' retWidgets.clear()
Public Sub clearRetWidgets()
    Erase retWidgets
    retCount = 0
End Sub

' retWidgets.size()
Public Function retWidgetCount() As Long
    retWidgetCount = retCount
End Function

' retWidgets[i].widget (0-based)
Public Function retWidgetAt(ByVal i As Long) As WIDGET
    Set retWidgetAt = retWidgets(i)
End Function
