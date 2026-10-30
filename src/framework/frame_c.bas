Attribute VB_Name = "frame_c"
' lib/framework/frame.cpp. A char* file buffer is a Byte array with the
' terminating 0 C adds; a PHYSFS_file* is physfs_win's handle number.
' C's assert() is a no-op in release builds and is left out. Missing:
' findHashOfFile (SHA-256).
' to_string and to_display_string are overloaded in C++: here they are
' to_string_video_backend, to_display_string_video_backend and
' to_display_string_WINDOW_MODE.
Option Explicit

'*************************************************************************
'	Player globals

Private mousewarp As Boolean

Public selectedPlayer As Long  ' Current player (uint32_t)
Public realSelectedPlayer As Long

' (VB) the game's side of the framework (IFrameworkGame): frameSetGame
Public frameGame As IFrameworkGame

' Global variables for the frame rate stuff
Private frameCount As Long
Private curFrames As Double   ' Number of frames elapsed since start (uint64_t)
Private lastFrames As Double
Private curTicks As Long      ' Number of ticks since execution started (uint32_t)
Private lastTicks As Long

' InitFrameStuff - needs to be called once before frame loop commences
Private Sub InitFrameStuff()
    frameCount = 0
    curFrames = 0
    lastFrames = 0
    curTicks = 0
    lastTicks = 0
End Sub

Public Function frameRate() As Long
    frameRate = frameCount
End Function

Public Function frameGetFrameNumber() As Long
    frameGetFrameNumber = FromU32(curFrames)
End Function

' Initialise the framework library. - PC version
Public Function frameInitialise() As Boolean
    ' Initialise the trig stuff
    If Not trigInitialise() Then
        Exit Function
    End If

    ' Initialise the input system
    inputInitialise

    ' Initialise the frame rate stuff
    InitFrameStuff

    ' Initialise the resource stuff
    If Not resInitialise() Then
        Exit Function
    End If

    frameInitialise = True
End Function

' Call this each cycle to do general house keeping.
Public Sub frameUpdate()
    curTicks = wzGetTicks()
    curFrames = curFrames + 1

    ' Update the framerate only once per second (uint32_t: lastTicks + 1000 wraps)
    If U32(curTicks) >= U32(FromU32(U32(lastTicks) + 1000)) Then
        frameCount = curFrames - lastFrames
        lastTicks = curTicks
        lastFrames = curFrames
    End If
End Sub

' Cleanup framework
Public Sub frameShutDown()
    ' Shutdown the resource stuff
    debug_ 130, LOG_NEVER, "frameShutDown", "No more resources!"
    resShutDown
End Sub

'(VB) The exe registers its IFrameworkGame; Nothing at shutdown.
Public Sub frameSetGame(ByVal g As IFrameworkGame)
    Set frameGame = g
End Sub

Public Sub setMouseWarp(ByVal value As Boolean)
    mousewarp = value
End Sub

Public Function getMouseWarp() As Boolean
    getMouseWarp = mousewarp
End Function

Public Function openLoadFile(ByVal fileName As String, ByVal hard_fail As Boolean) As Long
    Dim fileHandle As Long
    fileHandle = PHYSFS_openRead(fileName)
    If debugPartEnabled(LOG_NEVER) Then debug_ 147, LOG_NEVER, "openLoadFile", "Reading...[directory: %s] %s", PHYSFS_getRealDir(fileName), fileName
    If fileHandle = 0 Then
        If hard_fail Then
            ASSERT_FAILURE "!""unable to open file""", "lib/framework/frame.cpp:152", "openLoadFile", "file %s could not be opened: %s", fileName, WZ_PHYSFS_getLastError()
        Else
            debug_ 156, LOG_WZ, "openLoadFile", "optional file %s could not be opened: %s", fileName, WZ_PHYSFS_getLastError()
        End If
    End If

    openLoadFile = fileHandle
End Function

' Load the file with name pFileName into ppFileData, with a terminating 0. If AllocateMem is true
' the array is sized for the file, else it is the caller's buffer of pFileSize bytes, and pFileSize
' is adjusted to the actual loaded file size. If hard_fail is true, we will assert and report on failures.
Private Function loadFile2(ByVal pFileName As String, ppFileData() As Byte, pFileSize As Long, ByVal AllocateMem As Boolean, ByVal hard_fail As Boolean) As Boolean
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
        Exit Function  ' File size could not be determined. Is a directory?
    End If
    If Not (filesize < 2147483647#) Then
        ASSERT_FAILURE "filesize < static_cast<PHYSFS_sint64>(std::numeric_limits<PHYSFS_sint32>::max())", "lib/framework/frame.cpp:189", "loadFile2", """%s"" filesize >= std::numeric_limits<PHYSFS_sint32>::max()", pFileName
        Exit Function
    End If

    If AllocateMem Then
        ' Allocate a buffer to store the data and a terminating zero
        ReDim ppFileData(0 To filesize)
    Else
        If filesize > pFileSize Then
            debug_ 201, LOG_ERROR, "loadFile2", "No room for file %s, buffer is too small! Got: %d Need: %ld", pFileName, pFileSize, filesize
            Exit Function
        End If
    End If

    ' Load the file data
    length_read = 0
    If filesize > 0 Then length_read = PHYSFS_readBytes(pfile, ppFileData, 0, filesize)
    If length_read <> filesize Then
        If AllocateMem Then
            Erase ppFileData
        End If

        debug_ 218, LOG_ERROR, "loadFile2", "Reading %s short: %s", pFileName, WZ_PHYSFS_getLastError()
        Exit Function
    End If

    If PHYSFS_close(pfile) = 0 Then
        If AllocateMem Then
            Erase ppFileData
        End If

        debug_ 231, LOG_ERROR, "loadFile2", "Error closing %s: %s", pFileName, WZ_PHYSFS_getLastError()
        Exit Function
    End If

    ' Add the terminating zero
    If UBound(ppFileData) >= filesize Then ppFileData(filesize) = 0

    ' always set to correct size
    pFileSize = filesize

    loadFile2 = True
End Function

Public Function openSaveFile(ByVal fileName As String) As Long
    Dim fileHandle As Long, found As String
    fileHandle = PHYSFS_openWrite(fileName)
    If fileHandle = 0 Then
        found = PHYSFS_getRealDir(fileName)

        debug_ 258, LOG_ERROR, "openSaveFile", "%s could not be opened: %s", fileName, WZ_PHYSFS_getLastError()
        If StrPtr(found) <> 0 Then
            debug_ 261, LOG_ERROR, "openSaveFile", "%s found as %s", fileName, found
        End If

        Exit Function
    End If

    openSaveFile = fileHandle
End Function

' Save the data in the buffer into the given file.
Public Function saveFile(ByVal pFileName As String, pFileData() As Byte, ByVal fileSize As Long) As Boolean
    Dim pfile As Long, size As Long

    size = fileSize
    debug_ 279, LOG_WZ, "saveFile", "We are to write (%s) of size %d", pFileName, fileSize
    pfile = openSaveFile(pFileName)
    If pfile = 0 Then
        ASSERT_FAILURE "false", "lib/framework/frame.cpp:283", "saveFile", "Couldn't save file %s (%s)?", pFileName, WZ_PHYSFS_getLastError()
        Exit Function
    End If

    If size > 0 Then
        If PHYSFS_writeBytes(pfile, pFileData, 0, size) <> size Then
            debug_ 289, LOG_ERROR, "saveFile", "%s could not write: %s", pFileName, WZ_PHYSFS_getLastError()
            Exit Function
        End If
    End If
    If PHYSFS_close(pfile) = 0 Then
        debug_ 295, LOG_ERROR, "saveFile", "Error closing %s: %s", pFileName, WZ_PHYSFS_getLastError()
        Exit Function
    End If

    If StrPtr(PHYSFS_getRealDir(pFileName)) = 0 Then
        ' weird
        debug_ 303, LOG_ERROR, "saveFile", "PHYSFS_getRealDir(%s) returns NULL (%s)?!", pFileName, WZ_PHYSFS_getLastError()
    Else
        debug_ 307, LOG_WZ, "saveFile", "Successfully wrote to %s%s with %d bytes", PHYSFS_getRealDir(pFileName), pFileName, size
    End If
    saveFile = True
End Function

Public Function loadFile(ByVal pFileName As String, ppFileData() As Byte, pFileSize As Long, Optional ByVal hard_fail As Boolean = True) As Boolean
    loadFile = loadFile2(pFileName, ppFileData, pFileSize, True, hard_fail)
End Function

' load a file from disk into a fixed memory buffer
Public Function loadFileToBuffer(ByVal pFileName As String, pFileBuffer() As Byte, ByVal bufferSize As Long, pSize As Long) As Boolean
    pSize = bufferSize
    loadFileToBuffer = loadFile2(pFileName, pFileBuffer, pSize, False, True)
End Function

' as above but returns quietly if no file found
Public Function loadFileToBufferNoError(ByVal pFileName As String, pFileBuffer() As Byte, ByVal bufferSize As Long, pSize As Long) As Boolean
    pSize = bufferSize
    loadFileToBufferNoError = loadFile2(pFileName, pFileBuffer, pSize, False, False)
End Function

' C: loadFileToBufferVector (pFileName, std::vector<char>& outputBuffer, hard_fail, appendNullCharacter)
Public Function loadFileToBufferVector(ByVal pFileName As String, outputBuffer() As Byte, outputSize As Long, ByVal hard_fail As Boolean, ByVal appendNullCharacter As Boolean) As Boolean
    loadFileToBufferVector = loadFileToBufferVectorT(pFileName, outputBuffer, outputSize, hard_fail, appendNullCharacter)
End Function

' C: PHYSFS_printf (file, format, ...): the text written as UTF-8. args: the arguments as a Variant array.
Public Function PHYSFS_printfV(ByVal file As Long, ByVal format_ As String, args As Variant) As Boolean
    Dim vaBuffer As String, b() As Byte, n As Long
    vaBuffer = Left$(vsprintf(format_, args), 260 - 1)    ' char vaBuffer[PATH_MAX]
    n = utf8_encode(vaBuffer, b)
    If n = 0 Then
        PHYSFS_printfV = True
        Exit Function
    End If
    PHYSFS_printfV = (PHYSFS_writeBytes(file, b, 0, n) <> 0)
End Function

Private Function video_backend_names(ByVal i As Long) As String
    Select Case i
    Case 0: video_backend_names = "opengl"
    Case 1: video_backend_names = "opengles"
    Case 2: video_backend_names = "vulkan"
    Case Else: video_backend_names = "invalid" ' Must be last!
    End Select
End Function

Public Function video_backend_from_str(ByVal str_ As String, output_backend As Long) As Boolean
    Dim i As Long
    For i = 0 To video_backend_num_backends - 1
        If StrComp(video_backend_names(i), str_, vbTextCompare) = 0 Then
            output_backend = i
            video_backend_from_str = True
            Exit Function
        End If
    Next
End Function

Public Function to_string_video_backend(ByVal backend As Long) As String
    to_string_video_backend = video_backend_names(backend)
End Function

Public Function to_display_string_video_backend(ByVal backend As Long) As String
    Select Case backend
    Case video_backend_opengl
        to_display_string_video_backend = "OpenGL"
    Case video_backend_opengles
        to_display_string_video_backend = "OpenGL ES"
    Case video_backend_vulkan
        to_display_string_video_backend = "Vulkan"
    Case video_backend_num_backends
        debug_ 404, LOG_FATAL, "to_display_string", "Should never happen"
        to_display_string_video_backend = "n/a"
    Case Else
        to_display_string_video_backend = "n/a"
    End Select
End Function

Public Function to_display_string_WINDOW_MODE(ByVal mode As Long) As String
    Select Case mode
    Case WINDOW_MODE_windowed
        to_display_string_WINDOW_MODE = "Windowed"
    Case WINDOW_MODE_fullscreen
        to_display_string_WINDOW_MODE = "Fullscreen"
    Case Else
        to_display_string_WINDOW_MODE = "n/a" ' silence warning
    End Select
End Function
