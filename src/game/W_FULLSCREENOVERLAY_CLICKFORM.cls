VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W_FULLSCREENOVERLAY_CLICKFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class W_FULLSCREENOVERLAY_CLICKFORM
' lib/widget/form.h: W_FULLSCREENOVERLAY_CLICKFORM's fields (form_c); the initialisers are
' W_FULLSCREENOVERLAY_CLICKFORM_construct's.
Option Explicit

Public backgroundColor As Long          ' PIELIGHT = pal_RGBA(0, 0, 0, 125)
Public onClickedFunc As WzFunc          ' std::function<void ()>
Public onCancelPressed As WzFunc        ' std::function<void ()>

' protected
' std::weak_ptr<WIDGET> cutoutWidget: (VB) held strongly (setCutoutWidget): the widget does not know of
' this pointer, so an address would dangle when it dies; the C++ callers clear it on close
Public cutoutWidget As WIDGET
