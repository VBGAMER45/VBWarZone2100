VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_FORM
' lib/widget/form.h: W_FORM's fields (form_c); the initialisers are W_FORM_construct's.
Option Explicit

Public disableChildren As Boolean       ' Disable all child widgets if true
Public userMovable As Boolean           ' = false: Whether the user can drag the form around (NOTE: should only be used with forms on overlay screens, currently)

' private
Public minimizedTitle As WIDGET         ' std::shared_ptr<W_LABEL>
Private minimizedRect_ As WzRect
Public minimizedLeftButtonWidth As Long ' = 0
Public formState As W_FORM_FormState           ' = FormState::NORMAL
Public has_dragStart As Boolean         ' optional<Vector2i> dragStart
Private dragStart_ As Vector2i
Public minimizable As Boolean           ' = false

Friend Property Get minimizedRect() As WzRect
    minimizedRect = minimizedRect_
End Property

Friend Property Let minimizedRect(v As WzRect)
    minimizedRect_ = v
End Property

Friend Property Get dragStart() As Vector2i
    dragStart = dragStart_
End Property

Friend Property Let dragStart(v As Vector2i)
    dragStart_ = v
End Property
