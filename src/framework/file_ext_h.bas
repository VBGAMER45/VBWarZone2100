Attribute VB_Name = "file_ext_h"
' lib/framework/file_ext.h: loadFileToBufferVectorT, for T = char (a
' std::vector<char> is a Byte array plus its size). C's assert() is a no-op
' in release builds and is left out.
Option Explicit

Public Function loadFileToBufferVectorT(ByVal pFileName As String, outputBuffer() As Byte, outputSize As Long, ByVal hard_fail As Boolean, Optional ByVal appendNullCharacter As Boolean = True) As Boolean
    Dim pfile As Long, filesize As Double, length_read As Long

    If WZ_PHYSFS_isDirectory(pFileName) <> 0 Then
        Exit Function
    End If

    pfile = openLoadFile(pFileName, hard_fail)
    If pfile = 0 Then
        Exit Function
    End If

    filesize = PHYSFS_fileLength(pfile)
    If filesize < 0 Then
        PHYSFS_close pfile
        Exit Function  ' File size could not be determined. Is a directory?
    End If
    If Not (filesize < 2147483647#) Then
        PHYSFS_close pfile
        ASSERT_FAILURE "filesize < static_cast<PHYSFS_sint64>(std::numeric_limits<PHYSFS_sint32>::max())", "lib/framework/file_ext.h:49", "loadFileToBufferVectorT", """%s"" filesize >= std::numeric_limits<PHYSFS_sint32>::max()", pFileName
        Exit Function
    End If
    If Not (filesize < 4294967295#) Then
        PHYSFS_close pfile
        ASSERT_FAILURE "static_cast<PHYSFS_uint64>(filesize) < static_cast<PHYSFS_uint64>(std::numeric_limits<size_t>::max())", "lib/framework/file_ext.h:55", "loadFileToBufferVectorT", """%s"" filesize >= std::numeric_limits<size_t>::max()", pFileName
        Exit Function
    End If

    Erase outputBuffer
    outputSize = filesize
    If appendNullCharacter Then outputSize = outputSize + 1
    If outputSize > 0 Then ReDim outputBuffer(0 To outputSize - 1) Else ReDim outputBuffer(0 To 0)

    ' Load the file data
    length_read = 0
    If filesize > 0 Then length_read = PHYSFS_readBytes(pfile, outputBuffer, 0, filesize)
    If length_read <> filesize Then
        Erase outputBuffer
        outputSize = 0

        PHYSFS_close pfile
        debug_ 69, LOG_ERROR, "loadFileToBufferVectorT", "Reading %s short: %s", pFileName, WZ_PHYSFS_getLastError()
        Exit Function
    End If

    If PHYSFS_close(pfile) = 0 Then
        Erase outputBuffer
        outputSize = 0

        debug_ 78, LOG_ERROR, "loadFileToBufferVectorT", "Error closing %s: %s", pFileName, WZ_PHYSFS_getLastError()
        Exit Function
    End If

    If appendNullCharacter Then
        outputBuffer(outputSize - 1) = 0
    End If

    ' (filesize <= UDWORD max: true after the checks above)

    loadFileToBufferVectorT = True
End Function
