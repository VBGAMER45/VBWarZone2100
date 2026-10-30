VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionsKeyBindingsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/options/controlsoptions.cpp: the fields of class OptionsKeyBindingsEdit
' (VB) data of widget class OptionsKeyBindingsEdit
Option Explicit

Public weakSerial As Long               ' (VB) its std::weak_ptr serial (widgweak_vb)
Public parentOptionsForm As Long        ' std::weak_ptr<KeyOptionsForm>: its widgweak_vb serial
Public info As Long                     ' const KeyFunctionInfo&: a handle (InputManager& inputManager: gInputManager)
Private mappings_(0 To 1) As Long       ' std::array<optional<std::reference_wrapper<KeyMapping>>, KeyMappingSlot::LAST>: handles (0 = nullopt)
Private keyMappingSlotButtons_(0 To 1) As WIDGET ' std::array<std::shared_ptr<OptionsKeyBindingWidget>, KeyMappingSlot::LAST>
Public overlayScreen As W_SCREEN
Public onChangeHandlers As Collection    ' std::vector<std::function<void(WIDGET&)>>
Public onStartEditingHandler As WzFunc
Public onEndEditingHandler As WzFunc
Public cachedIdealWidth As Long          ' = 0
Public cachedIdealHeight As Long         ' = 0

Public Property Get mappings(ByVal i As Long) As Long
    mappings = mappings_(i)
End Property

Public Property Let mappings(ByVal i As Long, ByVal v As Long)
    mappings_(i) = v
End Property

Public Property Get keyMappingSlotButtons(ByVal i As Long) As WIDGET
    Set keyMappingSlotButtons = keyMappingSlotButtons_(i)
End Property

Public Property Set keyMappingSlotButtons(ByVal i As Long, ByVal v As WIDGET)
    Set keyMappingSlotButtons_(i) = v
End Property
