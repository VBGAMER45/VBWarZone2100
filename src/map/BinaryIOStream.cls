VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryIOStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map_io.h: class WzMap::BinaryIOStream, an
' open file of an IOProvider (its openBinaryStream). The pure virtual
' methods are this interface (Implements BinaryIOStream); the read and write
' helpers that the base class implements on top of them (readULE16,
' writeULE32...) are map_io_c's BinaryIOStream_<method>(stream, ...).
' optional<size_t> is the Boolean result (has_value) and the count in the
' last parameter. close() is closeStream (Close is a VB statement, and the
' method of an interface takes no underscore). The stream closes when its
' last reference goes (the C++ destructor).
Option Explicit

Public Enum OpenMode
    OpenMode_READ = 0
    OpenMode_WRITE = 1
End Enum

' Reads up to len_ bytes into buffer(offset ..); false on a read error
Public Function readBytes(buffer() As Byte, ByVal offset As Long, ByVal len_ As Long, bytesRead As Long) As Boolean
End Function

' Writes buffer(offset .. offset + len_ - 1); false on a write error
Public Function writeBytes(buffer() As Byte, ByVal offset As Long, ByVal len_ As Long, bytesWritten As Long) As Boolean
End Function

Public Function endOfStream() As Boolean
End Function

Public Function closeStream() As Boolean
End Function
