VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/urlrequest.h: struct MemoryStruct (a std::shared_ptr<MemoryStruct> is one of these): a response's bytes,
' memory(0 .. size - 1) (memory is an empty array when size is 0).
Option Explicit

Private memory_() As Byte
Public size As Long

' char *memory: a copy of the bytes (always at least one byte long, as the C++ buffer's terminating 0)
Public Property Get memory() As Byte()
    Dim b() As Byte
    If size = 0 Then
        ReDim b(0 To 0)
        memory = b
    Else
        memory = memory_
    End If
End Property

' (VB) the transport sets the bytes b(0 .. n - 1)
Friend Sub setMemory(b() As Byte, ByVal n As Long)
    size = n
    If n > 0 Then
        memory_ = b
        If UBound(memory_) <> n - 1 Then ReDim Preserve memory_(0 To n - 1)
    Else
        Erase memory_
    End If
End Sub
