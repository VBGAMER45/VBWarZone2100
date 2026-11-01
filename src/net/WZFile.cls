VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WZFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/netplay/wzfile.h, wzfile.cpp: struct WZFile, a thin wrapper over a PhysFS file handle, used by the network
' code for downloading files. Made by NetApi's WZFile_new (the handle, the file name, the hash's 32 bytes and
' the size); the game and WzNet share the handle (WzFramework's PhysFS). The C++ moves WZFiles (no copies):
' the port shares the object. The hash crosses as its bytes (hashBytes) or its hex string (hashString).
Option Explicit

Private handle_ As Long                 ' std::unique_ptr<PHYSFS_file, void(*)(PHYSFS_file*)> handle_ (0: none)
Private filename_ As String
Private hash_ As Sha256
Private size_ As Long                   ' uint32_t
Private pos_ As Long                    ' uint32_t: Current position, the range [0; currPos[ has been sent or received already.

Friend Sub construct(ByVal handle As Long, ByVal filename As String, hashBytes() As Byte, ByVal size As Long)
    Dim k As Long
    handle_ = handle
    filename_ = filename
    For k = 0 To 31
        hash_.bytes(k) = hashBytes(k)
    Next
    size_ = size
    pos_ = 0
End Sub

' WZFile::~WZFile(): the unique_ptr's deleter physfs_file_safe_close closes a handle still held
Private Sub Class_Terminate()
    If handle_ <> 0 Then physfs_file_safe_close_impl handle_
    handle_ = 0
End Sub

Public Function closeFile() As Boolean
    Dim f As Long
    If handle_ = 0 Then
        closeFile = False
        Exit Function
    End If
    f = handle_                         ' handle_.release()
    handle_ = 0
    closeFile = physfs_file_safe_close_impl(f)
End Function

Public Function handle() As Long
    handle = handle_
End Function

Public Property Get filename() As String
    filename = filename_
End Property

Public Property Let filename(ByVal value As String)
    filename_ = value
End Property

Friend Function hash() As Sha256
    hash = hash_
End Function

Public Function hashBytes() As Byte()
    Dim b(0 To 31) As Byte, k As Long
    For k = 0 To 31
        b(k) = hash_.bytes(k)
    Next
    hashBytes = b
End Function

Public Function hashString() As String
    hashString = Sha256_toString(hash_)
End Function

Public Property Get size() As Long
    size = size_
End Property

Public Property Let size(ByVal value As Long)
    size_ = value
End Property

Public Property Get pos() As Long
    pos = pos_
End Property

Public Property Let pos(ByVal value As Long)
    pos_ = value
End Property
