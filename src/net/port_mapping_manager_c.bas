Attribute VB_Name = "port_mapping_manager_c"
' lib/netplay/port_mapping_manager.cpp, port_mapping_manager.h: the PortMappingManager singleton (its members are
' this module's variables; PortMappingManager::instance().f() is f()). A PortMappingAsyncRequestHandle /
' PortMappingAsyncRequestPtr is a PortMappingAsyncRequest object (Nothing: nullptr; dynamic_pointer_cast is the
' object itself); a std::shared_ptr<PortMappingImpl> is a PortMappingImpl object. PortMappingDiscoveryStatus is
' NetApi's enum. A SuccessCallback / FailureCallback pair (the lambdas NETaddRedirects and NEThostGame attach) is a
' callback number, PMCB_*, whose two functions pmcb_success and pmcb_failure call.
' activeDiscoveries_, successfulRequests_ and a CallbacksPerRequest (std::unordered_maps) are arrays of keys and
' requests searched by key and walked in insertion order (MSVC's hash order would only show with two requests at
' once; 4.7.0 makes one). The mutex is left out (one thread).
' (VB) No threads (docs/CONVENTIONS.md): the monitoring thread's loop is thread_monitor_pass, one iteration at the
' first pass 50 ms after the last (its sleep_for(50ms)), while the thread "runs" (monitoringThread_). A task given
' to wzAsyncExecOnMainThread waits in this module's queue and runs at the start of the next pass, as on 4.7.0's
' next main loop iteration. PortMappingManager_poll is the pass, run by NETflush (NetApi's, each frame of the title
' loop and the game loop): the queued tasks, libplum's client thread (plum_poll), the monitoring thread.
' The singleton's destructor (PortMappingManager::shutdown at the process's exit) does not run: VB has no static
' destructors; NETshutdown calls shutdown as 4.7.0's does.
Option Explicit

Public Const WZ_PORT_MAPPING_ID_INVALID As Long = -1

' enum class PortMappingInternetProtocol : uint8_t
Public Const PortMappingInternetProtocol_TCP_IPV4 As Long = &H1
Public Const PortMappingInternetProtocol_TCP_IPV6 As Long = &H2
Public Const PortMappingInternetProtocol_UDP_IPV4 As Long = &H4
Public Const PortMappingInternetProtocol_UDP_IPV6 As Long = &H8

' enum class PortMappingImpl::Type
Public Const PortMappingImpl_Type_Libplum As Long = 0
Public Const PortMappingImpl_Type_Miniupnpc As Long = 1
Public Const PortMappingImpl_Type_End As Long = 2

Public Const TIMEPOINT_MAX As Double = 1E+300   ' PortMappingAsyncRequest::TimePoint::max()

' (VB) the callback numbers: the SuccessCallback / FailureCallback pairs attach_callback is given
Public Const PMCB_NETaddRedirects As Long = 1   ' netplay.cpp NETaddRedirects's console messages
Public Const PMCB_NEThostGame As Long = 2       ' netplay.cpp NEThostGame's SetupKnownExternalIPv4Info and NETEnableAllowJoining

' std::unordered_map<ImplMappingId, PortMappingAsyncRequestPtr, ImplMappingIdHasher> activeDiscoveries_;
Private ad_key() As String
Private ad_req() As PortMappingAsyncRequest
Private ad_count As Long
' std::unordered_map<PortWithProtocol, PortMappingAsyncRequestPtr> successfulRequests_;
Private sr_key() As String
Private sr_req() As PortMappingAsyncRequest
Private sr_count As Long
Private isInit_ As Boolean                      ' bool isInit_ = false;
Private loadedImpls_ As Collection              ' std::vector<std::shared_ptr<PortMappingImpl>> loadedImpls_;
Private currentImpl_ As PortMappingImpl         ' std::shared_ptr<PortMappingImpl> currentImpl_ = nullptr;
Private currImplTypeIdx_ As Long                ' PortMappingImpl::Type currImplTypeIdx_ = PortMappingImpl::Type::Libplum;
Private monitoringThread_ As Boolean            ' std::thread monitoringThread_: (VB) its passes run
Private monitorNextPass As Double               ' (VB) the end of its sleep_for(50ms)
Private stopRequested_ As Boolean               ' std::atomic_bool stopRequested_;

' (VB) CallbacksPerRequest: the requests and their CallbackVectors
Private cpr_req() As PortMappingAsyncRequest
Private cpr_cbs() As Collection
Private cpr_count As Long

' (VB) wzAsyncExecOnMainThread's tasks (see above)
Private Const TASK_LOG As Long = 0              ' PlumLogCallback's _debug line
Private Const TASK_SUCCESS As Long = 1          ' the success callbacks of a request
Private Const TASK_FAILURE As Long = 2          ' the failure callbacks of a request
Private Type AsyncTask
    kind As Long
    level As Long                               ' TASK_LOG: the code_part
    text As String                              ' TASK_LOG: the message; TASK_SUCCESS: the external IP address
    port As Long                                ' TASK_SUCCESS: the external port
    status As Long                              ' TASK_FAILURE: the request's status
    callbacks As Collection                     ' TASK_SUCCESS, TASK_FAILURE: the callback numbers
End Type
Private tasks() As AsyncTask
Private task_count As Long

' (VB) the maps' keys
Private Function ImplMappingId_key(ByVal implType As Long, ByVal mappingId As Long) As String
    ImplMappingId_key = implType & ":" & mappingId
End Function

Private Function PortWithProtocol_key(ByVal port As Long, ByVal protocol As Long) As String
    PortWithProtocol_key = port & ":" & protocol
End Function

' (VB) activeDiscoveries_.find(key): its index, -1 for end()
Private Function ad_find(ByVal key As String) As Long
    Dim i As Long
    For i = 0 To ad_count - 1
        If ad_key(i) = key Then
            ad_find = i
            Exit Function
        End If
    Next
    ad_find = -1
End Function

' (VB) activeDiscoveries_.emplace(key, req): nothing when the key is there
Private Sub ad_emplace(ByVal key As String, ByVal req As PortMappingAsyncRequest)
    If ad_find(key) >= 0 Then Exit Sub
    If ad_count = 0 Then
        ReDim ad_key(0 To 3)
        ReDim ad_req(0 To 3)
    ElseIf ad_count > UBound(ad_key) Then
        ReDim Preserve ad_key(0 To ad_count * 2 - 1)
        ReDim Preserve ad_req(0 To ad_count * 2 - 1)
    End If
    ad_key(ad_count) = key
    Set ad_req(ad_count) = req
    ad_count = ad_count + 1
End Sub

' (VB) activeDiscoveries_.erase(it)
Private Sub ad_erase(ByVal i As Long)
    Dim j As Long
    For j = i To ad_count - 2
        ad_key(j) = ad_key(j + 1)
        Set ad_req(j) = ad_req(j + 1)
    Next
    ad_count = ad_count - 1
    Set ad_req(ad_count) = Nothing
End Sub

Private Function sr_find(ByVal key As String) As Long
    Dim i As Long
    For i = 0 To sr_count - 1
        If sr_key(i) = key Then
            sr_find = i
            Exit Function
        End If
    Next
    sr_find = -1
End Function

Private Sub sr_emplace(ByVal key As String, ByVal req As PortMappingAsyncRequest)
    If sr_find(key) >= 0 Then Exit Sub
    If sr_count = 0 Then
        ReDim sr_key(0 To 3)
        ReDim sr_req(0 To 3)
    ElseIf sr_count > UBound(sr_key) Then
        ReDim Preserve sr_key(0 To sr_count * 2 - 1)
        ReDim Preserve sr_req(0 To sr_count * 2 - 1)
    End If
    sr_key(sr_count) = key
    Set sr_req(sr_count) = req
    sr_count = sr_count + 1
End Sub

Private Sub sr_erase(ByVal i As Long)
    Dim j As Long
    For j = i To sr_count - 2
        sr_key(j) = sr_key(j + 1)
        Set sr_req(j) = sr_req(j + 1)
    Next
    sr_count = sr_count - 1
    Set sr_req(sr_count) = Nothing
End Sub

' (VB) callbacksPerRequest[req] = callbacks
Private Sub cpr_set(ByVal req As PortMappingAsyncRequest, ByVal callbacks As Collection)
    Dim i As Long
    For i = 0 To cpr_count - 1
        If cpr_req(i) Is req Then
            Set cpr_cbs(i) = callbacks
            Exit Sub
        End If
    Next
    If cpr_count = 0 Then
        ReDim cpr_req(0 To 3)
        ReDim cpr_cbs(0 To 3)
    ElseIf cpr_count > UBound(cpr_req) Then
        ReDim Preserve cpr_req(0 To cpr_count * 2 - 1)
        ReDim Preserve cpr_cbs(0 To cpr_count * 2 - 1)
    End If
    Set cpr_req(cpr_count) = req
    Set cpr_cbs(cpr_count) = callbacks
    cpr_count = cpr_count + 1
End Sub

' (VB) callbacksPerRequest.clear()
Private Sub cpr_clear()
    Dim i As Long
    For i = 0 To cpr_count - 1
        Set cpr_req(i) = Nothing
        Set cpr_cbs(i) = Nothing
    Next
    cpr_count = 0
End Sub

' (VB) wzAsyncExecOnMainThread(task): queued for the next pass
Private Sub asyncExec(task As AsyncTask)
    If task_count = 0 Then
        ReDim tasks(0 To 7)
    ElseIf task_count > UBound(tasks) Then
        ReDim Preserve tasks(0 To task_count * 2 - 1)
    End If
    tasks(task_count) = task
    task_count = task_count + 1
End Sub

' (VB) the tasks queued before this pass, in their order (a task they queue runs at the next pass)
Private Sub runAsyncTasks()
    Dim n As Long, i As Long, task As AsyncTask, blank As AsyncTask, cb As Variant
    n = task_count
    For i = 0 To n - 1
        task = tasks(i)
        Select Case task.kind
        Case TASK_LOG
            debug__ 77, task.level, "PlumLogCallback", "%s", task.text
        Case TASK_SUCCESS
            For Each cb In task.callbacks
                pmcb_success cb, task.text, task.port
            Next
        Case TASK_FAILURE
            For Each cb In task.callbacks
                pmcb_failure cb, task.status
            Next
        End Select
    Next
    ' the tasks queued while these ran move to the front
    For i = n To task_count - 1
        tasks(i - n) = tasks(i)
    Next
    For i = task_count - n To task_count - 1
        If i >= 0 Then tasks(i) = blank
    Next
    task_count = task_count - n
End Sub

' (VB) PortMappingManager's pass (see the module's comment)
Public Sub PortMappingManager_poll()
    runAsyncTasks
    plum_poll                                   ' libplum's client thread
    thread_monitor_pass
End Sub

' (VB) PlumLogCallback's wzAsyncExecOnMainThread([wz_log_level, msg]() { _debug(...); })
Public Sub asyncExecPlumLog(ByVal wz_log_level As Long, ByVal msg As String)
    Dim task As AsyncTask
    task.kind = TASK_LOG
    task.level = wz_log_level
    task.text = msg
    asyncExec task
End Sub

' (VB) cb.first(externalHost, externalPort): the SuccessCallback of callback number cb
Private Sub pmcb_success(ByVal cb As Long, ByVal externalIp As String, ByVal externalPort As Long)
    Select Case cb
    Case PMCB_NETaddRedirects: NETaddRedirects_successCb externalIp, externalPort
    Case PMCB_NEThostGame: NEThostGame_successCb externalIp, externalPort
    End Select
End Sub

' (VB) cb.second(status): the FailureCallback of callback number cb
Private Sub pmcb_failure(ByVal cb As Long, ByVal status As PortMappingDiscoveryStatus)
    Select Case cb
    Case PMCB_NETaddRedirects: NETaddRedirects_failureCb status
    Case PMCB_NEThostGame: NEThostGame_failureCb status
    End Select
End Sub

' MARK: - General callback helpers

' void PortMappingImpl_MappingCallbackSuccess(PortMappingImpl::Type type, int mappingId, const std::string& externalHost,
'                                             uint16_t externalPort)
Public Sub PortMappingImpl_MappingCallbackSuccess(ByVal typ As Long, ByVal mappingId As Long, ByVal externalHost As String, ByVal externalPort As Long)
    resolve_success_fromimpl typ, mappingId, externalHost, externalPort
End Sub

' void PortMappingImpl_MappingCallbackFailure(PortMappingImpl::Type type, int mappingId)
Public Sub PortMappingImpl_MappingCallbackFailure(ByVal typ As Long, ByVal mappingId As Long)
    resolve_failure_fromimpl typ, mappingId, PortMappingDiscoveryStatus_FAILURE
End Sub

' MARK: - PortMappingManager

' constexpr PortMappingImpl::Type nextImplType(PortMappingImpl::Type t)
Private Function nextImplType(ByVal t As Long) As Long
    If Not (t < PortMappingImpl_Type_End) Then ASSERT_FAILURE "static_cast<std::underlying_type_t<PortMappingImpl::Type>>(t) < static_cast<std::underlying_type_t<PortMappingImpl::Type>>(PortMappingImpl::Type::End)", _
        "lib/netplay/port_mapping_manager.cpp:123", "nextImplType", "Unexpected PortMappingImpl type"
    nextImplType = t + 1
End Function

' std::shared_ptr<PortMappingImpl> PortMappingManager::makeImpl(PortMappingImpl::Type t)
Private Function makeImpl(ByVal t As Long) As PortMappingImpl
    Dim impl As PortMappingImpl
    Select Case t
    Case PortMappingImpl_Type_Libplum
        Set impl = New PortMappingImpl
        impl.constructLibPlum New PortMappingImpl_LibPlum
        Set makeImpl = impl
    Case PortMappingImpl_Type_Miniupnpc
        Set impl = New PortMappingImpl
        impl.constructMiniupnpc New PortMappingImpl_Miniupnpc
        Set makeImpl = impl
    Case PortMappingImpl_Type_End
        Set makeImpl = Nothing
    End Select
End Function

' bool PortMappingManager::switchToNextWorkingImpl(PortMappingImpl::Type startingTypeIdx)
Private Function switchToNextWorkingImpl(ByVal startingTypeIdx As Long) As Boolean
    Dim newImpl As PortMappingImpl
    Do
        Set newImpl = makeImpl(startingTypeIdx)
        If newImpl Is Nothing Then Exit Do
        If newImpl.init() Then
            currImplTypeIdx_ = startingTypeIdx
            Set currentImpl_ = newImpl
            loadedImpls_.Add newImpl
            switchToNextWorkingImpl = True
            Exit Function
        End If
        startingTypeIdx = nextImplType(startingTypeIdx)
    Loop
    switchToNextWorkingImpl = False
End Function

' void PortMappingManager::init()
Public Sub PortMappingManager_init()
    If isInit_ Then Exit Sub

    If loadedImpls_ Is Nothing Then Set loadedImpls_ = New Collection
    If switchToNextWorkingImpl(currImplTypeIdx_) Then isInit_ = True
End Sub

' void PortMappingManager::shutdown()
Public Sub PortMappingManager_shutdown()
    Dim v As Variant, impl As PortMappingImpl
    If Not isInit_ Then Exit Sub

    If Not currentImpl_ Is Nothing Then currentImpl_.shutdown
    For Each v In loadedImpls_
        Set impl = v
        impl.shutdown
    Next
    Set loadedImpls_ = New Collection           ' loadedImpls_.clear();
    Set currentImpl_ = Nothing

    Do While ad_count > 0                       ' activeDiscoveries_.clear();
        ad_erase ad_count - 1
    Loop
    Do While sr_count > 0                       ' successfulRequests_.clear();
        sr_erase sr_count - 1
    Loop

    ' Request the monitoring thread to stop and wait for it to join.
    stopRequested_ = True
    monitoringThread_ = False                   ' (VB) monitoringThread_.join(): its next pass would end it
    isInit_ = False
End Sub

' PortMappingAsyncRequestHandle PortMappingManager::create_port_mapping(uint16_t port, PortMappingInternetProtocol protocol)
Public Function PortMappingManager_create_port_mapping(ByVal port As Long, ByVal protocol As Long) As PortMappingAsyncRequest
    Dim h As PortMappingAsyncRequest, hasActiveDiscoveries As Boolean, i As Long, mappingId As Long
    If Not isInit_ Then Exit Function

    hasActiveDiscoveries = False
    ' First, check if there's a request for a given <port, protocol> pair
    ' that has been successfully executed before. If so, just reuse it.
    i = sr_find(PortWithProtocol_key(port, protocol))
    If i >= 0 Then
        Set PortMappingManager_create_port_mapping = sr_req(i)
        Exit Function
    End If

    mappingId = currentImpl_.create_port_mapping(port, protocol)
    If mappingId < 0 Then
        ' TODO: Could try other backends as well
        Exit Function
    End If
    Set h = New PortMappingAsyncRequest
    h.sourcePort = port
    h.protocol = protocol
    h.set_port_mapping_in_progress steadyNowMs() + 1000# * currentImpl_.get_discovery_timeout_seconds(), mappingId, currentImpl_

    hasActiveDiscoveries = ad_count > 0
    ad_emplace ImplMappingId_key(h.impl.get_impl_type(), mappingId), h

    If Not hasActiveDiscoveries Then
        ' The monitoring thread should aready have finished processing.
        ' So, need to restart it.
        monitoringThread_ = False               ' (VB) monitoringThread_.join()
        stopRequested_ = False
        monitoringThread_ = True                ' monitoringThread_ = std::thread([this] { thread_monitor_function(); });
        monitorNextPass = 0
    End If
    Set PortMappingManager_create_port_mapping = h
End Function

' bool PortMappingManager::destroy_port_mapping(const PortMappingAsyncRequestHandle& p)
Public Function PortMappingManager_destroy_port_mapping(ByVal h As PortMappingAsyncRequest) As Boolean
    Dim i As Long, wasFailedMapping As Boolean, result As Boolean
    If Not isInit_ Then ASSERT_FAILURE "isInit_", "lib/netplay/port_mapping_manager.cpp:262", "destroy_port_mapping", "PortMappingManager is not initialized": Exit Function
    If h Is Nothing Then ASSERT_FAILURE "h != nullptr", "lib/netplay/port_mapping_manager.cpp:264", "destroy_port_mapping", "Invalid request handle": Exit Function
    If h.mappingId = WZ_PORT_MAPPING_ID_INVALID Then
        PortMappingManager_destroy_port_mapping = True
        Exit Function
    End If
    ' Remove this request from active/successful requests.
    If h.in_progress() Then
        i = ad_find(ImplMappingId_key(h.impl.get_impl_type(), h.mappingId))
        If i >= 0 Then ad_erase i
    ElseIf h.succeeded() Then
        i = sr_find(PortWithProtocol_key(h.sourcePort, h.protocol))
        If i >= 0 Then sr_erase i
    End If
    wasFailedMapping = h.failed()
    ' Destroy the internal mapping.
    result = h.destroy_mapping()
    h.mappingId = WZ_PORT_MAPPING_ID_INVALID
    If wasFailedMapping Then PortMappingManager_destroy_port_mapping = True Else PortMappingManager_destroy_port_mapping = result
End Function

' PortMappingDiscoveryStatus PortMappingManager::get_status(const PortMappingAsyncRequestHandle& p) const
Public Function PortMappingManager_get_status(ByVal h As PortMappingAsyncRequest) As PortMappingDiscoveryStatus
    If h Is Nothing Then ASSERT_FAILURE "h != nullptr", "lib/netplay/port_mapping_manager.cpp:297", "get_status", "Invalid request handle": PortMappingManager_get_status = PortMappingDiscoveryStatus_FAILURE: Exit Function
    PortMappingManager_get_status = h.status()
End Function

' void PortMappingManager::attach_callback(const PortMappingAsyncRequestHandle& p, SuccessCallback successCb,
'                                          FailureCallback failureCb): cb is the pair's callback number
Public Sub PortMappingManager_attach_callback(ByVal h As PortMappingAsyncRequest, ByVal cb As Long)
    Dim s As PortMappingDiscoveryStatus, task As AsyncTask
    If h Is Nothing Then ASSERT_FAILURE "h != nullptr", "lib/netplay/port_mapping_manager.cpp:305", "attach_callback", "Invalid request handle": Exit Sub
    s = h.s
    If s = PortMappingDiscoveryStatus_IN_PROGRESS Then
        h.callbacks.Add cb
        Exit Sub
    End If
    ' If the discovery has already finished with either success or failure, we can schedule
    ' the callback right away, because the monitoring thread may already have finished executing.
    If s = PortMappingDiscoveryStatus_SUCCESS Then
        task.kind = TASK_SUCCESS                ' wzAsyncExecOnMainThread([successCb, ip, port]() { successCb(ip, port); });
        Set task.callbacks = New Collection
        task.callbacks.Add cb
        task.text = h.discoveredIPaddress
        task.port = h.discoveredPort
        asyncExec task
        Exit Sub
    End If
    If s = PortMappingDiscoveryStatus_FAILURE Or s = PortMappingDiscoveryStatus_TIMEOUT Then
        task.kind = TASK_FAILURE                ' wzAsyncExecOnMainThread([failureCb, s]() { failureCb(s); });
        Set task.callbacks = New Collection
        task.callbacks.Add cb
        task.status = s
        asyncExec task
        Exit Sub
    End If
End Sub

' void PortMappingManager::scheduleCallbacksOnMainThread(CallbacksPerRequest cbReqMap): the map is cpr_*
Private Sub scheduleCallbacksOnMainThread()
    Dim i As Long, task As AsyncTask, blank As AsyncTask
    If cpr_count = 0 Then Exit Sub
    For i = 0 To cpr_count - 1
        task = blank
        If cpr_req(i).succeeded() And cpr_cbs(i).Count > 0 Then
            task.kind = TASK_SUCCESS
            Set task.callbacks = cpr_cbs(i)
            task.text = cpr_req(i).discoveredIPaddress
            task.port = cpr_req(i).discoveredPort
            asyncExec task
        ElseIf cpr_req(i).failed() And cpr_cbs(i).Count > 0 Then ' Failed for any reason or timed out
            task.kind = TASK_FAILURE
            Set task.callbacks = cpr_cbs(i)
            task.status = cpr_req(i).s
            asyncExec task
        ElseIf cpr_req(i).status() = PortMappingDiscoveryStatus_NOT_STARTED Then
            ' That would probably mean that the request in question has already been cancelled.
            ' Just ignore it and don't schedule any callbacks.
        Else
            ASSERT_FAILURE "false", "lib/netplay/port_mapping_manager.cpp:371", "scheduleCallbacksOnMainThread", "Should be unreachable"
        End If
    Next
End Sub

' (VB) takeCallbacksAndRemoveFromActiveList(req): the lambda of thread_monitor_function
Private Sub takeCallbacksAndRemoveFromActiveList(ByVal req As PortMappingAsyncRequest, ByVal i As Long)
    ' Move callbacks out of the finished request and remove from active list.
    cpr_set req, req.callbacks
    Set req.callbacks = New Collection
    ad_erase i
End Sub

' void PortMappingManager::thread_monitor_function()
' (VB) one iteration of its loop (see the module's comment)
Private Sub thread_monitor_pass()
    Dim i As Long, req As PortMappingAsyncRequest, requestsSwitchingImpl As Collection, r As Variant
    If Not monitoringThread_ Then Exit Sub
    If steadyNowMs() < monitorNextPass Then Exit Sub    ' (VB) still in its sleep_for(50ms)

    If stopRequested_ Then
        monitoringThread_ = False               ' the thread returns
        Exit Sub
    End If

    Set requestsSwitchingImpl = New Collection
    ' If all inflight requests have already finished, stop monitoring changes.
    If ad_count = 0 Then
        monitoringThread_ = False
        Exit Sub
    End If
    ' Go through active discovery requests and collect callbacks to be scheduled
    ' for execution at the end of the current loop iteration.
    i = 0
    Do While i < ad_count
        Set req = ad_req(i)
        If req.succeeded() Then
            sr_emplace PortWithProtocol_key(req.sourcePort, req.protocol), req
            takeCallbacksAndRemoveFromActiveList req, i
        ElseIf req.timed_out(steadyNowMs()) Then
            If resolve_failure_internal(req, PortMappingDiscoveryStatus_TIMEOUT) Then
                ' resolve_failure_internal dispatched to a different implementation and reset timeout - wait to see what it returns
                ' queue updating its key in the active list
                requestsSwitchingImpl.Add req
                ad_erase i
            Else
                ' failure is final
                takeCallbacksAndRemoveFromActiveList req, i
            End If
        ElseIf req.failed() Then
            ' Already resolved, but need to move out the callbacks and remove from active list.
            takeCallbacksAndRemoveFromActiveList req, i
        Else
            If Not req.in_progress() Then ASSERT_FAILURE "req->in_progress()", "lib/netplay/port_mapping_manager.cpp:430", "thread_monitor_function", "Unexpected request state: %d", CLng(req.s)
            i = i + 1
        End If
    Loop
    ' Schedule execution of registered callbacks for all requests, which have already
    ' transitioned to some terminal (either success or failure) state.
    scheduleCallbacksOnMainThread
    cpr_clear
    ' Move any requestsSwitchingImpl to the appropriate mapping entry in the activeDiscoveriesMap_
    For Each r In requestsSwitchingImpl
        Set req = r
        ad_emplace ImplMappingId_key(req.impl.get_impl_type(), req.mappingId), req
    Next

    ' Sleep for another 50ms and re-check the status for all active requests.
    monitorNextPass = steadyNowMs() + 50
End Sub

' void PortMappingManager::resolve_success_fromimpl(PortMappingImpl::Type type, int mappingId, std::string externalIp,
'                                                   uint16_t externalPort)
Private Sub resolve_success_fromimpl(ByVal typ As Long, ByVal mappingId As Long, ByVal externalIp As String, ByVal externalPort As Long)
    Dim req As PortMappingAsyncRequest
    Set req = active_request_for_id_internal(typ, mappingId)
    If req Is Nothing Then
        debug_ 463, LOG_NET, "resolve_success_fromimpl", "Mapping %d: no active request found in PortMappingManager, probably timed out\n", mappingId
        Exit Sub
    End If
    If Not resolve_success_internal(req, externalIp, externalPort) Then
        If req.mappingId <> WZ_PORT_MAPPING_ID_INVALID Then
            req.destroy_mapping
            req.mappingId = WZ_PORT_MAPPING_ID_INVALID
        End If
    End If
End Sub

' void PortMappingManager::resolve_failure_fromimpl(PortMappingImpl::Type type, int mappingId,
'                                                   PortMappingDiscoveryStatus failureStatus)
Private Sub resolve_failure_fromimpl(ByVal typ As Long, ByVal mappingId As Long, ByVal failureStatus As PortMappingDiscoveryStatus)
    Dim req As PortMappingAsyncRequest, i As Long
    Set req = active_request_for_id_internal(typ, mappingId)
    If req Is Nothing Then
        debug_ 487, LOG_NET, "resolve_failure_fromimpl", "Mapping %d: no active request found in PortMappingManager, probably timed out\n", mappingId
        Exit Sub
    End If
    If resolve_failure_internal(req, failureStatus) Then
        ' update its key in the active list
        i = ad_find(ImplMappingId_key(typ, mappingId))
        If i >= 0 Then ad_erase i
        ad_emplace ImplMappingId_key(req.impl.get_impl_type(), req.mappingId), req
    End If
End Sub

' MARK: - BEGIN: Unsafe functions that must only be called if mtx_ lock is already held!

' PortMappingAsyncRequestPtr PortMappingManager::active_request_for_id_internal(PortMappingImpl::Type type,
'                                                                              PortMappingAsyncRequest::MappingId id) const
Private Function active_request_for_id_internal(ByVal typ As Long, ByVal id As Long) As PortMappingAsyncRequest
    Dim i As Long
    i = ad_find(ImplMappingId_key(typ, id))
    If i >= 0 Then Set active_request_for_id_internal = ad_req(i)
End Function

' bool PortMappingManager::resolve_success_internal(PortMappingAsyncRequestPtr req, std::string externalIp,
'                                                   uint16_t externalPort)
Private Function resolve_success_internal(ByVal req As PortMappingAsyncRequest, ByVal externalIp As String, ByVal externalPort As Long) As Boolean
    If req.s <> PortMappingDiscoveryStatus_IN_PROGRESS Then
        ' Someone has already gotten ahead of us
        If req.s = PortMappingDiscoveryStatus_SUCCESS Then
            ' already resolved to success ?
            resolve_success_internal = True
        Else
            ' failed to resolve to success
            ' parent will handle cleanup as needed
            resolve_success_internal = False
        End If
        Exit Function
    End If
    req.s = PortMappingDiscoveryStatus_SUCCESS
    req.deadline = TIMEPOINT_MAX
    req.discoveredIPaddress = externalIp
    req.discoveredPort = externalPort
    resolve_success_internal = True
End Function

' bool PortMappingManager::resolve_failure_internal(PortMappingAsyncRequestPtr req,
'                                                   PortMappingDiscoveryStatus failureStatus)
Private Function resolve_failure_internal(ByVal req As PortMappingAsyncRequest, ByVal failureStatus As PortMappingDiscoveryStatus) As Boolean
    Dim finalFailure As Boolean, anotherImpl As Boolean, newMappingId As Long
    If Not (failureStatus = PortMappingDiscoveryStatus_FAILURE Or failureStatus = PortMappingDiscoveryStatus_TIMEOUT) Then ASSERT_FAILURE _
        "failureStatus == PortMappingDiscoveryStatus::FAILURE || failureStatus == PortMappingDiscoveryStatus::TIMEOUT", _
        "lib/netplay/port_mapping_manager.cpp:552", "resolve_failure_internal", "Status should either be failure or timeout"
    If req.s <> PortMappingDiscoveryStatus_IN_PROGRESS Then
        ' Someone has already gotten ahead of us - nothing to do there.
        resolve_failure_internal = False
        Exit Function
    End If

    finalFailure = False

    ' Check for another implementation
    anotherImpl = Not (req.impl Is currentImpl_)    ' (VB) || evaluates its right side only when needed
    If Not anotherImpl Then anotherImpl = switchToNextWorkingImpl(nextImplType(currImplTypeIdx_))
    If anotherImpl Then
        ' Another implementation is available - fallback to it and try again
        Do
            newMappingId = currentImpl_.create_port_mapping(req.sourcePort, req.protocol)
            If newMappingId >= 0 Then
                ' new impl is attempting request
                req.set_port_mapping_in_progress steadyNowMs() + 1000# * currentImpl_.get_discovery_timeout_seconds(), newMappingId, currentImpl_
                resolve_failure_internal = True ' attempting with different implementation
                Exit Function
            End If
        Loop While switchToNextWorkingImpl(nextImplType(currImplTypeIdx_))

        ' new impl(s) failed
        finalFailure = True
    Else
        ' no additional impl available - fail out
        finalFailure = True
    End If

    If finalFailure Then
        req.s = failureStatus
        req.deadline = TIMEPOINT_MAX
        req.discoveredIPaddress = ""
        req.discoveredPort = 0
    End If
    resolve_failure_internal = False
End Function

' MARK: - END: Unsafe functions that must only be called if mtx_ lock is already held!
