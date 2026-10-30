Attribute VB_Name = "input_context_c"
' src/input/context.cpp, src/input/context.h: the input contexts the key mappings belong to, and the
' ContextManager that keeps their states.
' (VB) A ContextId (const char *) is a String (InputContext_ALWAYS_ACTIVE, ...). An InputContext is its
' index (a handle: aInputContext(index)), as operator== compares the indices; the copies C++ keeps in
' the map share the index. The one ContextManager is display.cpp's gInputManager's (InputManager's
' contexts()): its methods are gInputManager_contexts_<method>, and registerDefaultContexts takes no
' managers. A PriorityCondition (std::function<bool()>) is a WzFunc (FN_ids 10900-10909).
' InputContext::State is InputContext_State_PRIORITIZED (0), _ACTIVE (1), _INACTIVE (2), in the C++ order.
Option Explicit

Public Const InputContext_ALWAYS_ACTIVE As String = "ALWAYS_ACTIVE"
Public Const InputContext_BACKGROUND As String = "BACKGROUND"
Public Const InputContext_GAMEPLAY As String = "GAMEPLAY"
Public Const InputContext_RADAR As String = "RADAR"
Public Const InputContext_DEBUG_MISC As String = "DEBUG_MISC"
Public Const InputContext_DEBUG_LEVEL_EDITOR As String = "DEBUG_LEVEL_EDITOR"
Public Const InputContext_DEBUG_HAS_SELECTION As String = "DEBUG_HAS_SELECTION"

' enum class InputContext::State
Public Enum InputContext_State
    ' The input context is prioritized. e.g. `InputContext::RADAR` is prioritized when player hovers
    ' over the radar, prioritizing any bindings belonging to that context.
    InputContext_State_PRIORITIZED
    ' The input context is active. Any bindings will use the lower `active` priority.
    InputContext_State_ACTIVE
    ' The input context is inactive. Bindings in the context are ignored.
    InputContext_State_INACTIVE
End Enum

Private Const UINT_MAX As Long = -1     ' std::numeric_limits<unsigned int>::max()

' the PriorityConditions (the lambdas of InputContext's constructor and registerDefaultContexts)
Public Const FN_InputContext_noCondition As Long = 10900        ' []() { return false; }
Public Const FN_InputContext_radarCondition As Long = 10901     ' []() { return isMouseOverRadar() && !isInDesignScreen(); }
Public Const FN_InputContext_debugCondition As Long = 10902     ' [&dbgInputManager]() { return dbgInputManager.isDebugPrioritized(); }

' struct ContextPriority { const unsigned int prioritized; const unsigned int active; }: two fields of
' InputContext (unsigned: compared with ULT)
Private Type InputContext
    id As String                    ' const ContextId
    bIsAlwaysActive As Boolean
    priority_prioritized As Long    ' const ContextPriority priority
    priority_active As Long
    index As Long                   ' const unsigned int
    displayName As String
    defaultState As Long            ' const State
    condition As WzFunc             ' const PriorityCondition
End Type

' (VB) every InputContext made, by index (the copies C++ makes share theirs)
Private aInputContext() As InputContext
Private aInputContext_alloc As Long

Private inputCtxIndexCounter As Long    ' static unsigned int inputCtxIndexCounter = 0;

' static const InputContext NULL_CONTEXT (nullContext()'s static): made on the first call
Private NULL_CONTEXT As Long
Private NULL_CONTEXT_made As Boolean

' ContextManager (gInputManager's contextManager)
' std::unordered_map<std::string, InputContext> contexts: the contexts' indices, in the order registered
Private contexts() As Long
Private contexts_size As Long
' std::vector<std::vector<InputContext::State>> states: states(context index, level)
Private states() As Long
Private states_size As Long
Private states_width As Long            ' the inner vectors' size
Private bDirty As Boolean

Private Function isInDesignScreen() As Boolean
    isInDesignScreen = intMode = INT_DESIGN
End Function

Private Function nullContext() As Long
    If Not NULL_CONTEXT_made Then
        NULL_CONTEXT = InputContext_construct("__NULL", False, 0, 0, InputContext_State_INACTIVE, "null")
        NULL_CONTEXT_made = True
    End If
    nullContext = NULL_CONTEXT
End Function

' void registerDefaultContexts(ContextManager& contextManager, DebugInputManager& dbgInputManager): the
' InputContext initialisers {id, bIsAlwaysActive, priority, initialState, displayName[, condition]}; a
' priority { p, a } is (p, a), a single value both
Public Sub registerDefaultContexts()
    Dim MAX_ICONTEXT_PRIORITY As Long
    Dim alwaysActive As Long, background As Long, gameplay As Long, radar As Long, debug_ctx As Long, debugLvlEditor As Long, debugSelection As Long
    MAX_ICONTEXT_PRIORITY = UINT_MAX
    alwaysActive = InputContext_construct(InputContext_ALWAYS_ACTIVE, True, MAX_ICONTEXT_PRIORITY, MAX_ICONTEXT_PRIORITY, InputContext_State_ACTIVE, "Global Keys")
    background = InputContext_construct(InputContext_BACKGROUND, False, 0, 0, InputContext_State_ACTIVE, "Other Keys")
    gameplay = InputContext_construct(InputContext_GAMEPLAY, False, 1, 1, InputContext_State_ACTIVE, "Gameplay Keys")
    radar = InputContext_construct(InputContext_RADAR, False, 2, 0, InputContext_State_ACTIVE, "Radar Keys", WzFunc_new(FN_InputContext_radarCondition))
    debug_ctx = InputContext_construct(InputContext_DEBUG_MISC, False, MAX_ICONTEXT_PRIORITY, 0, InputContext_State_INACTIVE, "Debug Keys", WzFunc_new(FN_InputContext_debugCondition))
    debugLvlEditor = InputContext_construct(InputContext_DEBUG_LEVEL_EDITOR, False, MAX_ICONTEXT_PRIORITY, 0, InputContext_State_INACTIVE, "Debug (level editor)", WzFunc_new(FN_InputContext_debugCondition))
    debugSelection = InputContext_construct(InputContext_DEBUG_HAS_SELECTION, False, MAX_ICONTEXT_PRIORITY, 0, InputContext_State_INACTIVE, "Debug (selection)", WzFunc_new(FN_InputContext_debugCondition))

    gInputManager_contexts_registerContext nullContext()
    gInputManager_contexts_registerContext alwaysActive
    gInputManager_contexts_registerContext background
    gInputManager_contexts_registerContext gameplay
    gInputManager_contexts_registerContext radar
    gInputManager_contexts_registerContext debug_ctx
    gInputManager_contexts_registerContext debugLvlEditor
    gInputManager_contexts_registerContext debugSelection
End Sub

' InputContext(const ContextId id, const bool bIsAlwaysActive, const ContextPriority priority, const State initialState,
' const char* const displayName[, const PriorityCondition condition]): the new context's index (no condition: []() { return false; })
Public Function InputContext_construct(ByVal id As String, ByVal bIsAlwaysActive As Boolean, ByVal prioritized As Long, ByVal active As Long, ByVal initialState As Long, ByVal displayName As String, Optional ByVal condition As WzFunc = Nothing) As Long
    Dim index As Long
    If condition Is Nothing Then Set condition = WzFunc_new(FN_InputContext_noCondition)
    index = inputCtxIndexCounter
    inputCtxIndexCounter = inputCtxIndexCounter + 1
    If index >= aInputContext_alloc Then
        aInputContext_alloc = aInputContext_alloc + 16
        ReDim Preserve aInputContext(0 To aInputContext_alloc - 1)
    End If
    aInputContext(index).id = id
    aInputContext(index).bIsAlwaysActive = bIsAlwaysActive
    aInputContext(index).priority_prioritized = prioritized
    aInputContext(index).priority_active = active
    aInputContext(index).index = index
    aInputContext(index).displayName = displayName
    aInputContext(index).defaultState = initialState
    Set aInputContext(index).condition = condition
    InputContext_construct = index
End Function

' const std::string InputContext::getDisplayName() const
Public Function InputContext_getDisplayName(ByVal context As Long) As String
    InputContext_getDisplayName = aInputContext(context).displayName
End Function

' bool InputContext::isAlwaysActive() const
Public Function InputContext_isAlwaysActive(ByVal context As Long) As Boolean
    InputContext_isAlwaysActive = aInputContext(context).bIsAlwaysActive
End Function

' (VB) the context's id
Public Function InputContext_id(ByVal context As Long) As String
    InputContext_id = aInputContext(context).id
End Function

' (VB) gInputManager.contexts().get(contextId).getDisplayName()
Public Function gInputManager_contexts_getDisplayName(ByVal contextId As String) As String
    gInputManager_contexts_getDisplayName = InputContext_getDisplayName(gInputManager_contexts_get(contextId))
End Function

' (VB) gInputManager.contexts().get(contextId).isAlwaysActive()
Public Function gInputManager_contexts_isAlwaysActive(ByVal contextId As String) As Boolean
    gInputManager_contexts_isAlwaysActive = InputContext_isAlwaysActive(gInputManager_contexts_get(contextId))
End Function

' (VB) the contexts' map: the registered context with the id, or -1 (find()'s end())
Private Function contexts_find(ByVal contextId As String) As Long
    Dim i As Long
    For i = 0 To contexts_size - 1
        If StrComp(aInputContext(contexts(i)).id, contextId, vbBinaryCompare) = 0 Then
            contexts_find = contexts(i)
            Exit Function
        End If
    Next
    contexts_find = -1
End Function

' const InputContext& ContextManager::get(const ContextId& contextId) const
Public Function gInputManager_contexts_get(ByVal contextId As String) As Long
    Dim found_ As Long
    found_ = contexts_find(contextId)
    If found_ <> -1 Then
        gInputManager_contexts_get = found_
    Else
        debug_ 118, LOG_WARNING, "get", "Tried to get missing input context ""%s"". Returning null context instead!", contextId
        gInputManager_contexts_get = nullContext()
    End If
End Function

' void ContextManager::set(const ContextId& contextId, const InputContext::State newState)
Public Sub gInputManager_contexts_set(ByVal contextId As String, ByVal newState As Long)
    Dim context As Long, oldState As Long, back As Long
    context = gInputManager_contexts_get(contextId)

    back = states_size - 1
    oldState = states(aInputContext(context).index, back)
    If oldState <> newState Then
        states(aInputContext(context).index, back) = newState
        bDirty = True
    End If
End Sub

' bool ContextManager::isActive(const ContextId& contextId) const
Public Function gInputManager_contexts_isActive(ByVal contextId As String) As Boolean
    Dim context As Long
    context = gInputManager_contexts_get(contextId)

    If Not (states_size <> 0) Then ASSERT_FAILURE "!states.empty()", "src/input/context.cpp:140", "isActive", "No InputContext::States available": gInputManager_contexts_isActive = False: Exit Function
    gInputManager_contexts_isActive = states(aInputContext(context).index, states_size - 1) <> InputContext_State_INACTIVE
End Function

' unsigned int ContextManager::getPriority(const ContextId contextId) const
Public Function gInputManager_contexts_getPriority(ByVal contextId As String) As Long
    Dim context As Long
    context = gInputManager_contexts_get(contextId)

    Select Case states(aInputContext(context).index, states_size - 1)
    Case InputContext_State_PRIORITIZED
        gInputManager_contexts_getPriority = aInputContext(context).priority_prioritized
    Case InputContext_State_ACTIVE
        gInputManager_contexts_getPriority = aInputContext(context).priority_active
    Case Else ' InputContext::State::INACTIVE, default
        gInputManager_contexts_getPriority = 0
    End Select
End Function

' void ContextManager::resetStates()
Public Sub gInputManager_contexts_resetStates()
    Dim i As Long, context As Long
    ' auto state = std::vector<InputContext::State>(contexts.size(), InputContext::State::INACTIVE);
    ' states.clear(); states.push_back(state); (the one level, filled in place)
    states_width = contexts_size
    If states_width = 0 Then states_width = 1
    ReDim states(0 To states_width - 1, 0 To 3)
    For i = 0 To states_width - 1
        states(i, 0) = InputContext_State_INACTIVE
    Next
    For i = 0 To contexts_size - 1
        context = contexts(i)
        states(aInputContext(context).index, 0) = aInputContext(context).defaultState
    Next
    states_size = 1

    bDirty = True
End Sub

' void ContextManager::makeAllInactive()
Public Sub gInputManager_contexts_makeAllInactive()
    Dim i As Long, context As Long
    For i = 0 To contexts_size - 1
        context = contexts(i)
        If Not InputContext_isAlwaysActive(context) Then
            gInputManager_contexts_set aInputContext(context).id, InputContext_State_INACTIVE
        End If
    Next

    ' No need to set dirty as set() handles setting it if necessary
End Sub

' void ContextManager::updatePriorityStatus()
Public Sub gInputManager_contexts_updatePriorityStatus()
    Dim i As Long, context As Long
    For i = 0 To contexts_size - 1
        context = contexts(i)
        If gInputManager_contexts_isActive(aInputContext(context).id) Then
            If CBool(WzFunc_call(aInputContext(context).condition)) Then
                gInputManager_contexts_set aInputContext(context).id, InputContext_State_PRIORITIZED
            Else
                gInputManager_contexts_set aInputContext(context).id, InputContext_State_ACTIVE
            End If
        End If
    Next
End Sub

' void ContextManager::registerContext(InputContext context): contexts.insert({ context.id, context }) (an id
' already there keeps its context)
Public Sub gInputManager_contexts_registerContext(ByVal context As Long)
    If contexts_find(aInputContext(context).id) <> -1 Then Exit Sub
    ReDim Preserve contexts(0 To contexts_size)
    contexts(contexts_size) = context
    contexts_size = contexts_size + 1
End Sub

' void ContextManager::pushState()
Public Sub gInputManager_contexts_pushState()
    Dim i As Long
    ' const auto copyOfCurrentState = std::vector<InputContext::State>(states.back());
    If states_size > UBound(states, 2) Then
        ReDim Preserve states(0 To states_width - 1, 0 To states_size * 2)
    End If
    For i = 0 To states_width - 1
        states(i, states_size) = states(i, states_size - 1)
    Next
    states_size = states_size + 1
    bDirty = True
End Sub

' void ContextManager::popState()
Public Sub gInputManager_contexts_popState()
    If Not (states_size > 1) Then ASSERT_FAILURE "states.size() > 1", "src/input/context.cpp:217", "popState", "Unbalanced push / pop": Exit Sub
    states_size = states_size - 1
    bDirty = True
End Sub

' bool ContextManager::isDirty() const
Public Function gInputManager_contexts_isDirty() As Boolean
    gInputManager_contexts_isDirty = bDirty
End Function

' void ContextManager::clearDirty()
Public Sub gInputManager_contexts_clearDirty()
    bDirty = False
End Sub

' MARK: - the lambdas

Public Function input_context_c_fn(ByVal f As WzFunc, args() As Variant) As Variant
    Select Case f.id
    Case FN_InputContext_noCondition                ' []() { return false; }
        input_context_c_fn = False
    Case FN_InputContext_radarCondition             ' []() { return isMouseOverRadar() && !isInDesignScreen(); }
        input_context_c_fn = isMouseOverRadar() And Not isInDesignScreen()
    Case FN_InputContext_debugCondition             ' [&dbgInputManager]() { return dbgInputManager.isDebugPrioritized(); }
        input_context_c_fn = DebugInputManager_isDebugPrioritized()
    End Select
End Function
