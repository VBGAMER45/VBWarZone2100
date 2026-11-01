Attribute VB_Name = "plum_dummytls_c"
' libplum: Copyright (c) 2021 Paul-Louis Ageneau. This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0 (ref/libplum/LICENSE; https://mozilla.org/MPL/2.0/); this port of it is too.
' libplum src/dummytls.c, dummytls.h: the part a client without a DummyTLS domain reaches. 4.7.0 gives plum_init no
' dummytls_domain (NULL), so the domain stays empty: dummytls_get_host returns -1 at its first test and nothing else
' runs. The certificate requests (dummytls_set_domain's, dummytls_renew_certs, dummytls_get_cert, resolve_host) and
' plum_get_dummytls_certificate / plum_get_dummytls_host are left out. The mutex is left out (one thread).
Option Explicit

' dummytls_state_t state (its domain: the certificates are left out)
Private state_domain As String

' void dummytls_init()
Public Sub dummytls_init()
    state_domain = ""                           ' memset(&state, 0, sizeof(state));
End Sub

' void dummytls_cleanup()
Public Sub dummytls_cleanup()
    ' clear_certs(&state.certs): (VB) there are none
End Sub

' int dummytls_get_host(const struct sockaddr *sa, char *buffer, size_t size)
Public Function dummytls_get_host(sa() As Byte, buffer As String, ByVal size As Long) As Long
    If state_domain = "" Then
        dummytls_get_host = -1
        Exit Function
    End If
    ' resolve_host(sa, state.domain, buffer, size): (VB) left out, not reached (see above)
    dummytls_get_host = -1
End Function
