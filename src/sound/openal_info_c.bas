Attribute VB_Name = "openal_info_c"
' lib/sound/openal_info.cpp, openal_info.h: the OpenAL devices and the device's information, for the log
' (sound_InitLibrary with LOG_SOUND on). The OpenALInfo functions return their text (the C++ hands it
' to an outputHandler: the caller does what its lambda did). A list of strings ("name\0name\0\0") is a
' String with vbNullChar separators; a nullptr string is vbNullString. The #ifndef ALC_ENUMERATE_ALL_EXT
' names are al_win's. Output_ALInfo and checkALErrors (nothing calls them) are not ported.
Option Explicit

Private Const MAX_WIDTH As Long = 80

Private Function printList(ByVal list As String, ByVal separator As String) As String
    Dim buf As String, col As Long, length As Long, indent As String, nxt As Long, p As Long

    col = MAX_WIDTH
    indent = "    "

    If Len(list) = 0 Then
        printList = vbLf & indent & "!!! none !!!"
        Exit Function
    End If

    p = 1
    Do
        nxt = InStr(p, list, separator)
        If nxt <> 0 Then
            length = nxt - p
            Do
                nxt = nxt + 1
            Loop While Mid$(list, nxt, 1) = separator
        Else
            length = Len(list) - p + 1
        End If

        If length + col + 2 >= MAX_WIDTH Then
            buf = buf & vbLf & indent
            col = Len(indent)
        Else
            buf = buf & " "
            col = col + 1
        End If

        buf = buf & Mid$(list, p, length)
        col = col + length

        If nxt = 0 Or nxt > Len(list) Then Exit Do
        buf = buf & " "
        col = col + 1

        p = nxt
    Loop

    printList = buf
End Function

Private Function printDeviceList(ByVal list As String) As String
    Dim buf As String, p As Long, e As Long

    If Len(list) = 0 Or Left$(list, 1) = vbNullChar Then
        buf = "    !!! none !!!" & vbLf
    Else
        p = 1
        Do
            e = InStr(p, list, vbNullChar)
            If e = 0 Then e = Len(list) + 1
            buf = buf & "    " & Mid$(list, p, e - p) & vbLf
            p = e + 1
        Loop While p <= Len(list) And Mid$(list, p, 1) <> vbNullChar
    End If

    printDeviceList = buf
End Function

' #define checkALCErrors(buf, x) checkALCErrors(buf, (x),__LINE__): linenum is the caller's line
Private Function checkALCErrors(buf As String, ByVal device As Long, ByVal linenum As Long) As Long
    Dim err_ As Long, pErrStr As String, h As String
    err_ = alcGetError(device)
    If err_ <> ALC_NO_ERROR Then
        pErrStr = alcGetString(device, err_)
        If StrPtr(pErrStr) = 0 Then pErrStr = "<n/a>"
        h = Hex$(err_)
        If Len(h) < 4 Then h = String$(4 - Len(h), "0") & h
        buf = buf & "ALC Error: " & pErrStr & " (0x" & h & "), @ " & linenum & vbLf
    End If
    checkALCErrors = err_
End Function

Public Function Output_ALCInfo(ByVal device As Long) As String
    Dim buf As String, major As Long, minor As Long, devname As String, srate As Long

    If device <> 0 Then
        devname = vbNullString
        buf = buf & vbLf
        If alcIsExtensionPresent(device, "ALC_ENUMERATE_ALL_EXT") <> AL_FALSE Then
            devname = alcGetString(device, ALC_ALL_DEVICES_SPECIFIER)
        End If
        If checkALCErrors(buf, device, 188) <> ALC_NO_ERROR Or StrPtr(devname) = 0 Then
            devname = alcGetString(device, ALC_DEVICE_SPECIFIER)
        End If
        If StrPtr(devname) = 0 Then devname = "<Name not found>"
        buf = buf & "** Info for device """ & devname & """ **" & vbLf
    End If
    alcGetIntegerv device, ALC_MAJOR_VERSION, 1, major
    alcGetIntegerv device, ALC_MINOR_VERSION, 1, minor
    If checkALCErrors(buf, device, 196) = ALC_NO_ERROR Then
        buf = buf & "ALC version: " & major & "." & minor & vbLf
    End If
    If device <> 0 Then
        buf = buf & "ALC extensions:" & printList(alcGetString(device, ALC_EXTENSIONS), " ") & vbLf
        checkALCErrors buf, device, 203

        alcGetIntegerv device, ALC_FREQUENCY, 1, srate
        buf = buf & "Output frequency: " & srate & " Hz " & vbLf

        alcGetIntegerv device, ALC_REFRESH, 1, srate
        buf = buf & "Update rate: " & srate & " Hz " & vbLf
    End If

    Output_ALCInfo = buf
End Function

Public Function Output_PlaybackDevices() As String
    Dim buf As String, enumerateDevicesSpecifier As Long, defaultDeviceParam As Long, pDefaultDeviceStr As String

    buf = "Available playback devices:" & vbLf
    enumerateDevicesSpecifier = ALC_DEVICE_SPECIFIER
    If alcIsExtensionPresent(0, "ALC_ENUMERATE_ALL_EXT") <> AL_FALSE Then
        enumerateDevicesSpecifier = ALC_ALL_DEVICES_SPECIFIER
    End If
    buf = buf & printDeviceList(alcGetString(0, enumerateDevicesSpecifier))

    defaultDeviceParam = ALC_DEFAULT_DEVICE_SPECIFIER
    If alcIsExtensionPresent(0, "ALC_ENUMERATE_ALL_EXT") <> AL_FALSE Then
        defaultDeviceParam = ALC_DEFAULT_ALL_DEVICES_SPECIFIER
    End If
    pDefaultDeviceStr = alcGetString(0, defaultDeviceParam)
    If StrPtr(pDefaultDeviceStr) = 0 Then pDefaultDeviceStr = "<Not found>"
    buf = buf & "Default playback device: " & pDefaultDeviceStr & vbLf

    buf = buf & Output_ALCInfo(0)

    Output_PlaybackDevices = buf
End Function
