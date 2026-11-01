Attribute VB_Name = "plum_noprotocol_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/noprotocol.c, noprotocol.h: the "protocol" of a host with a public address. Its state (state->impl)
' is this module's impl; the mutex is left out (one thread) and noprotocol_idle's condition variable wait is
' resumable (plum_socket_h's PLUM_PENDING): it ends the client thread's pass while time is left.
Option Explicit

Private Const AF_INET As Long = 2

' typedef struct noprotocol_impl_t (the mutex and the condition variable are left out)
Private Type noprotocol_impl_t
    interrupted As Boolean
End Type

Private impl As noprotocol_impl_t               ' state->impl

' (VB) noprotocol_idle's state while it waits
Private idle_at As Long
Private idle_end_timestamp As Double

' int noprotocol_init(protocol_state_t *state)
Public Function noprotocol_init(state As protocol_state_t) As Long
    Dim zeroState As protocol_state_t, zeroImpl As noprotocol_impl_t
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Initializing no protocol state"

    state = zeroState
    state.impl = True                           ' malloc(sizeof(noprotocol_impl_t))

    impl = zeroImpl
    ' mutex_init(&impl->mutex, 0); cond_init(&impl->interrupt_cond);

    noprotocol_init = PROTOCOL_ERR_SUCCESS
End Function

' int noprotocol_cleanup(protocol_state_t *state)
Public Function noprotocol_cleanup(state As protocol_state_t) As Long
    plum_log_write PLUM_LOG_LEVEL_VERBOSE, "Cleaning up no protocol state"

    ' mutex_destroy(&impl->mutex); cond_destroy(&impl->interrupt_cond);

    state.impl = False                          ' free(state->impl); state->impl = NULL;
    noprotocol_cleanup = PROTOCOL_ERR_SUCCESS
End Function

' int noprotocol_discover(protocol_state_t *state, timediff_t duration)
Public Function noprotocol_discover(state As protocol_state_t, ByVal duration As Double) As Long
    ' Dummy
    noprotocol_discover = PROTOCOL_ERR_SUCCESS
End Function

' int noprotocol_map(protocol_state_t *state, const client_mapping_t *mapping, protocol_map_output_t *output,
'                    timediff_t duration)
Public Function noprotocol_map(state As protocol_state_t, mapping As client_mapping_t, output As protocol_map_output_t, ByVal duration As Double) As Long
    Dim zeroOutput As protocol_map_output_t, local_ As addr_record_t, local_str As String, zeroAddr As addr_record_t
    output = zeroOutput

    If net_get_default_interface(AF_INET, local_) = 0 Then
        If addr_is_public(local_.addr) Then
            If plum_log_is_enabled(PLUM_LOG_LEVEL_INFO) Then
                addr_record_to_string local_, local_str, ADDR_MAX_STRING_LEN
                plum_log_write PLUM_LOG_LEVEL_INFO, "Local address is public: %s", local_str
            End If

            output.state = PROTOCOL_MAP_STATE_SUCCESS
            output.refresh_timestamp = current_timestamp() + CLIENT_RECHECK_PERIOD
            output.external_addr = local_
            addr_set_port output.external_addr.addr, mapping.internal_port
            noprotocol_map = PROTOCOL_ERR_SUCCESS
            Exit Function
        End If
    Else
        plum_log_write PLUM_LOG_LEVEL_WARN, "Unable to get default interface address"
    End If

    output.state = PROTOCOL_MAP_STATE_FAILURE
    output.refresh_timestamp = current_timestamp() + CLIENT_RECHECK_PERIOD
    output.external_addr = zeroAddr
    noprotocol_map = PROTOCOL_ERR_SUCCESS       ' report mapping failure but keep running the protocol
End Function

' int noprotocol_unmap(protocol_state_t *state, const client_mapping_t *mapping, timediff_t duration)
Public Function noprotocol_unmap(state As protocol_state_t, mapping As client_mapping_t, ByVal duration As Double) As Long
    ' Dummy
    noprotocol_unmap = PROTOCOL_ERR_SUCCESS
End Function

' int noprotocol_idle(protocol_state_t *state, timediff_t duration)
Public Function noprotocol_idle(state As protocol_state_t, ByVal duration As Double) As Long
    Dim now As Double
    ' mutex_lock(&impl->mutex);
    If idle_at = 1 Then
        idle_at = 0                             ' (VB) resumed
    Else
        idle_end_timestamp = current_timestamp() + duration
    End If
    Do While Not impl.interrupted
        now = current_timestamp()
        If idle_end_timestamp <= now Then
            noprotocol_idle = PROTOCOL_ERR_SUCCESS
            Exit Function
        End If
        ' cond_timedwait(&impl->interrupt_cond, &impl->mutex, left): (VB) the thread's pass ends here
        idle_at = 1
        noprotocol_idle = PLUM_PENDING
        Exit Function
    Loop
    impl.interrupted = False
    noprotocol_idle = PROTOCOL_ERR_INTERRUPTED
End Function

' int noprotocol_interrupt(protocol_state_t *state, bool hard)
Public Function noprotocol_interrupt(state As protocol_state_t, ByVal hard As Boolean) As Long
    ' (void)hard; ignored, as discover and map are instantaneous
    impl.interrupted = True
    ' cond_signal(&impl->interrupt_cond);
    noprotocol_interrupt = PROTOCOL_ERR_SUCCESS
End Function
