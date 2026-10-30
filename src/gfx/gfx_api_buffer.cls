VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gfx_api_buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' gfx_api.h: gfx_api::buffer, an abstract base that manages a single gfx buffer
' (gl_buffer and null_buffer implement it). The data is an address (VarPtr of
' the first element of a VB array or of a UDT; 0 = NULL) and a size in bytes.
' current_buffer_size is currentBufferSize. The buffer is freed when its last
' reference goes (C++'s delete).
Option Explicit

' Create a new data store for the buffer. Any existing data store will be deleted.
' The new data store is created with the specified `size` in bytes.
' If `data` is not NULL, the data store is initialized with `size` bytes of data from the pointer.
' - If data is NULL, a data store of the specified `size` is still created, but its contents may be uninitialized (and thus undefined).
'
' NOTE: A single buffer instance should only be uploaded / updated *once per frame*.
Public Sub upload(ByVal size As Long, ByVal data As Long)
End Sub

' Update `size` bytes in the existing data store, from the `start` offset, using the data at `data`.
' - The `start` offset must be within the range of the existing data store (0 to buffer_size - 1, inclusive).
' - The `size` of the data specified (+ the offset) must not attempt to write past the end of the existing data store.
' - A data store must be allocated before it can be updated - call `upload` on a `buffer` instance before `update`.
Public Sub update(ByVal start As Long, ByVal size As Long, ByVal data As Long, Optional ByVal flag As update_flag = update_flag_none)
End Sub

Public Function currentBufferSize() As Long
End Function

Public Sub bind()
End Sub
