Attribute VB_Name = "openal_error_c"
' lib/sound/openal_error.cpp, openal_error.h: OpenAL's errors in the log. The macros sound_GetError()
' and sound_GetContextError(device) pass AT_MACRO: here the callers pass the "file:line" location
' themselves (sound_GetError "lib/sound/openal_track.cpp:412"); the functions are the C's
' __sound_GetError and __sound_GetContextError (VB names can't start with an underscore).
Option Explicit

Public Function sound_GetError(ByVal location_description As String) As Long
    Dim errorString As String, haveString As Boolean, error_ As Long
    error_ = alGetError()

    If error_ = AL_NO_ERROR Then
        sound_GetError = error_
        Exit Function
    End If

    haveString = True
    Select Case error_
    Case AL_INVALID_NAME
        errorString = "AL_INVALID_NAME: Invalid name parameter passed"

    Case AL_INVALID_ENUM
        errorString = "AL_INVALID_ENUM: Invalid enum value"

    Case AL_INVALID_VALUE
        errorString = "AL_INVALID_VALUE: Invalid parameter value"

    Case AL_INVALID_OPERATION
        errorString = "AL_INVALID_OPERATION: Illegal call"

    Case AL_OUT_OF_MEMORY
        errorString = "AL_OUT_OF_MEMORY: OpenAL ran out of memory"

    Case Else
        haveString = False
        debug_ 63, LOG_SOUND, "__sound_GetError", "OpenAL raised an unknown error code (%d), at %s. Please report this number " & _
                  "(along with the fact that it is an ""unknown OpenAL error code""): 0x%x", error_, location_description, error_
    End Select

    If haveString Then
        debug_ 70, LOG_SOUND, "__sound_GetError", "OpenAL raised an error: ""%s""; at %s", errorString, location_description
    End If

    sound_GetError = error_
End Function

Public Function sound_GetContextError(ByVal device As Long, ByVal location_description As String) As Long
    Dim errorString As String, haveString As Boolean, error_ As Long
    error_ = alcGetError(device)

    If error_ = ALC_NO_ERROR Then
        sound_GetContextError = error_
        Exit Function
    End If

    haveString = True
    Select Case error_
    Case ALC_INVALID_DEVICE
        errorString = "ALC_INVALID_DEVICE: Invalid or no device selected"

    Case ALC_INVALID_CONTEXT
        errorString = "ALC_INVALID_CONTEXT: Invalid or no context selected"

    Case ALC_INVALID_ENUM
        errorString = "ALC_INVALID_ENUM: Invalid enum value"

    Case ALC_INVALID_VALUE
        errorString = "ALC_INVALID_VALUE: Invalid parameter value"

    Case ALC_OUT_OF_MEMORY
        errorString = "ALC_OUT_OF_MEMORY: OpenAL ran out of memory"

    Case Else
        haveString = False
        debug_ 110, LOG_SOUND, "__sound_GetContextError", "OpenAL raised an unknown context error code (%d), at %s. Please report this number " & _
                  "(along with the fact that it is an ""unknown OpenAL error code""): 0x%x", error_, location_description, error_
    End Select

    If haveString Then
        debug_ 117, LOG_SOUND, "__sound_GetContextError", "OpenAL raised a context error: ""%s""; at %s", errorString, location_description
    End If

    sound_GetContextError = error_
End Function
