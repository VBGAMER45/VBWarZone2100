VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DropdownWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class DropdownWidget
' lib/widget/dropdown.h: DropdownWidget's fields (dropdown_c). items is a Collection (element n is item
' n + 1); optional<AtlasImage> dropdownCaretImage is has_dropdownCaretImage and dropdownCaretImage.
Option Explicit

Public weakSerial As Long
Public items As Collection              ' std::vector<std::shared_ptr<DropdownItemWrapper>>
Public itemsList As WIDGET              ' std::shared_ptr<ScrollableListWidget>
Public overlayScreen As W_SCREEN
Public selectedItem As WIDGET
Public canChange As WzFunc
Public onChange As WzFunc
Public onOpen As WzFunc
Public onClose As WzFunc
Public mouseOverItem As WIDGET
Public mouseDownItem As WIDGET
Public overlayYPosOffset As Long        ' int32_t = 0
Public has_dropdownCaretImage As Boolean
Private m_dropdownCaretImage As AtlasImage
Private m_dropdownCaretImageSize As WzSize
Private m_dropdownCaretImagePadding As Padding   ' = {}
Public menuStyle As DropdownWidget_DropdownMenuStyle ' = DropdownMenuStyle::InPlace
Public isDisabled As Boolean            ' = false

Private Sub Class_Initialize()
    Set items = New Collection
End Sub

Friend Property Get dropdownCaretImage() As AtlasImage
    dropdownCaretImage = m_dropdownCaretImage
End Property

Friend Property Let dropdownCaretImage(v As AtlasImage)
    m_dropdownCaretImage = v
End Property

Friend Property Get dropdownCaretImageSize() As WzSize
    dropdownCaretImageSize = m_dropdownCaretImageSize
End Property

Friend Property Let dropdownCaretImageSize(v As WzSize)
    m_dropdownCaretImageSize = v
End Property

Friend Property Get dropdownCaretImagePadding() As Padding
    dropdownCaretImagePadding = m_dropdownCaretImagePadding
End Property

Friend Property Let dropdownCaretImagePadding(v As Padding)
    m_dropdownCaretImagePadding = v
End Property
