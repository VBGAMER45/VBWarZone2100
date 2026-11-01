Attribute VB_Name = "port_mapping_manager_impl_libplum_c"
' lib/netplay/port_mapping_manager_impl_libplum.cpp: its anonymous namespace's callbacks, given to libplum as
' function numbers (FN_*: plum_log_c and plum_client_c call them by number); the class PortMappingImpl_LibPlum is
' PortMappingImpl_LibPlum.cls. They run in libplum's client thread's pass (the main thread here).
Option Explicit

' (VB) the function numbers of the callbacks (a plum_log_callback_t, a plum_mapping_callback_t; 0 is NULL)
Public Const FN_PlumLogCallback As Long = 1
Public Const FN_PlumMappingCallback As Long = 2

' This function is run in its own thread managed by LibPLum! Do not call any non-threadsafe functions!
' void PlumMappingCallback(int mappingId, plum_state_t state, const plum_mapping_t* mapping)
Public Sub PlumMappingCallback(ByVal mappingId As Long, ByVal state As Long, mapping As plum_mapping_t)
    debug_ 42, LOG_NET, "PlumMappingCallback", "LibPlum cb: Port mapping %d: state=%d\n", mappingId, state
    Select Case state
    Case PLUM_STATE_SUCCESS
        debug_ 45, LOG_NET, "PlumMappingCallback", "Mapping %d: success, internal=%hu, external=%s:%hu\n", mappingId, mapping.internal_port, _
            mapping.external_host, mapping.external_port
        PortMappingImpl_MappingCallbackSuccess PortMappingImpl_Type_Libplum, mappingId, mapping.external_host, mapping.external_port

    Case PLUM_STATE_FAILURE
        debug_ 51, LOG_NET, "PlumMappingCallback", "Mapping %d: failed\n", mappingId
        PortMappingImpl_MappingCallbackFailure PortMappingImpl_Type_Libplum, mappingId

    Case Else
    End Select
End Sub

' void PlumLogCallback(plum_log_level_t level, const char* message)
Public Sub PlumLogCallback(ByVal level As Long, ByVal message As String)
    Dim wz_log_level As Long
    wz_log_level = LOG_NET
    Select Case level
    Case PLUM_LOG_LEVEL_FATAL, PLUM_LOG_LEVEL_ERROR, PLUM_LOG_LEVEL_WARN
        wz_log_level = LOG_INFO
    Case Else
    End Select
    If debugPartEnabled(wz_log_level) Then
        ' wzAsyncExecOnMainThread([wz_log_level, msg]() { _debug(__LINE__, wz_log_level, "PlumLogCallback", "%s", msg.c_str()); });
        asyncExecPlumLog wz_log_level, message
    End If
End Sub
