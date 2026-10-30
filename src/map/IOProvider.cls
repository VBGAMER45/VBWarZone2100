VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' lib/wzmaplib/include/wzmaplib/map_io.h, lib/wzmaplib/src/map_io.cpp: class
' WzMap::IOProvider, the file access wzmaplib uses. This is the interface
' (Implements IOProvider): StdIOProvider is the library's own, on the
' Windows file API; the game has one on PhysFS. The base class's
' non-virtual methods (pathJoin, pathDirName, pathBaseName) and its
' default enumerateFilesRecursive / enumerateFoldersRecursive are map_io_c's
' IOProvider_<method>(io, ...).
' std::vector<char>& fileData is a Byte array with its size; the
' enumeration callbacks (enumFunc, which can stop the enumeration) are the
' caller's loop over the names returned in files(0 .. count - 1), as in
' physfs_ext_c. The base class's enumerateFiles / enumerateFolders return
' false ("Must implement in subclasses"), as a class that does not
' implement them should.
Option Explicit

Public Enum LoadFullFileResult
    LoadFullFileResult_SUCCESS = 0
    LoadFullFileResult_FAILURE_OPEN = 1
    LoadFullFileResult_FAILURE_READ = 2
    LoadFullFileResult_FAILURE_EXCEEDS_MAXFILESIZE = 3
End Enum

' Nothing when the file cannot be opened
Public Function openBinaryStream(ByVal filename As String, ByVal mode As OpenMode) As BinaryIOStream
End Function

' fileData(0 .. fileSize - 1); maxFileSize 0: no limit
Public Function loadFullFile(ByVal filename As String, fileData() As Byte, fileSize As Long, ByVal maxFileSize As Long, ByVal appendNullCharacter As Boolean) As LoadFullFileResult
End Function

Public Function writeFullFile(ByVal filename As String, fileData() As Byte, ByVal fileSize As Long) As Boolean
End Function

' Creates a directory (and any required parent directories) and returns true on success (or if the directory already exists)
Public Function makeDirectory(ByVal directoryPath As String) As Boolean
End Function

Public Function pathSeparator() As String
End Function

' Check if a file exists at a path
Public Function fileExists(ByVal filePath As String) As Boolean
End Function

' The files in basePath, in files(0 .. count - 1)
Public Function enumerateFiles(ByVal basePath As String, files() As String, count As Long) As Boolean
    ' Must implement in subclasses
    count = 0
    enumerateFiles = False
End Function

' The subfolders of basePath, in folders(0 .. count - 1)
Public Function enumerateFolders(ByVal basePath As String, folders() As String, count As Long) As Boolean
    ' Must implement in subclasses
    count = 0
    enumerateFolders = False
End Function
