VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WidgetHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/widget/widgbase.h: struct WidgetHelp, formatted help information for a UI element. `New
' WidgetHelp` is WidgetHelp(); the setters return Me, as C++'s return *this. InteractionTriggers is
' widget_c's enum; a WidgetInteractionTriggerFlagSet (std::set<InteractionTriggers>) is a mask of
' 2 ^ trigger (widget_c's WidgetHelp_triggers(t, ...) makes one). The vectors interactions and
' relatedKeybindings are interactionCount() / interactionTriggers(i) / interactionDescription(i) and
' relatedKeybindingCount() / relatedKeybinding(i) (0-based).
Option Explicit

Public title_ As String                 ' WzString title
Public description As String            ' WzString
' std::vector<InteractionDescription> interactions: {triggers, description}
Private interactionTriggers_() As Long
Private interactionDescriptions_() As String
Private interactionCount_ As Long
' std::vector<KeybindingInfo> relatedKeybindings: {keybindingName} (for lookup on display with the input manager)
Private relatedKeybindings_() As String
Private relatedKeybindingCount_ As Long

Public Function setTitle(ByVal newValue As String) As WidgetHelp
    title_ = newValue
    Set setTitle = Me
End Function

Public Function setDescription(ByVal newValue As String) As WidgetHelp
    description = newValue
    Set setDescription = Me
End Function

' WidgetHelp& addInteraction(WidgetInteractionTriggerFlagSet triggers, const WzString& effectDescription)
Public Function addInteraction(ByVal triggers As Long, ByVal effectDescription As String) As WidgetHelp
    If interactionCount_ = 0 Then
        ReDim interactionTriggers_(0 To 3)
        ReDim interactionDescriptions_(0 To 3)
    ElseIf interactionCount_ > UBound(interactionTriggers_) Then
        ReDim Preserve interactionTriggers_(0 To 2 * interactionCount_ - 1)
        ReDim Preserve interactionDescriptions_(0 To 2 * interactionCount_ - 1)
    End If
    interactionTriggers_(interactionCount_) = triggers
    interactionDescriptions_(interactionCount_) = effectDescription
    interactionCount_ = interactionCount_ + 1
    Set addInteraction = Me
End Function

Public Function addRelatedKeybinding(ByVal keybindingName As String) As WidgetHelp
    If relatedKeybindingCount_ = 0 Then
        ReDim relatedKeybindings_(0 To 3)
    ElseIf relatedKeybindingCount_ > UBound(relatedKeybindings_) Then
        ReDim Preserve relatedKeybindings_(0 To 2 * relatedKeybindingCount_ - 1)
    End If
    relatedKeybindings_(relatedKeybindingCount_) = keybindingName
    relatedKeybindingCount_ = relatedKeybindingCount_ + 1
    Set addRelatedKeybinding = Me
End Function

'--------------------------------------------------------------------------
' (VB) the vectors

Public Function interactionCount() As Long
    interactionCount = interactionCount_
End Function

' interactions[i].triggers: the mask
Public Function interactionTriggers(ByVal i As Long) As Long
    interactionTriggers = interactionTriggers_(i)
End Function

' interactions[i].triggers.count(t)
Public Function interactionHasTrigger(ByVal i As Long, ByVal t As InteractionTriggers) As Boolean
    interactionHasTrigger = (interactionTriggers_(i) And (2 ^ t)) <> 0
End Function

Public Function interactionDescription(ByVal i As Long) As String
    interactionDescription = interactionDescriptions_(i)
End Function

Public Function relatedKeybindingCount() As Long
    relatedKeybindingCount = relatedKeybindingCount_
End Function

Public Function relatedKeybinding(ByVal i As Long) As String
    relatedKeybinding = relatedKeybindings_(i)
End Function

' (VB) a copy (C++ copies a WidgetHelp by value: optional<WidgetHelp> help = _help)
Public Function copy() As WidgetHelp
    Dim c As WidgetHelp, i As Long
    Set c = New WidgetHelp
    c.title_ = title_
    c.description = description
    For i = 0 To interactionCount_ - 1
        c.addInteraction interactionTriggers_(i), interactionDescriptions_(i)
    Next
    For i = 0 To relatedKeybindingCount_ - 1
        c.addRelatedKeybinding relatedKeybindings_(i)
    Next
    Set copy = c
End Function
