VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameworkApi"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' WzFramework.dll's public interface. GlobalMultiUse: the exe and the other
' DLLs call these procedures as if they were their own. The procedures
' forward to the DLL's modules, named in front (debug_c.debug_init), so that
' they never reach the class's own methods of the same name.
Option Explicit

' debug.h
Public Enum code_part
    LOG_ALL         ' special: sets all to on
    LOG_MAIN
    LOG_SOUND
    LOG_VIDEO
    LOG_WZ
    LOG_3D
    LOG_TEXTURE
    LOG_NET
    LOG_MEMORY
    LOG_WARNING     ' special; on in debug mode
    LOG_ERROR       ' special; on by default
    LOG_NEVER       ' if too verbose for anything but dedicated debugging...
    LOG_SCRIPT
    LOG_MOVEMENT
    LOG_ATTACK
    LOG_FOG
    LOG_SENSOR
    LOG_GUI
    LOG_MAP
    LOG_SAVE
    LOG_SYNC
    LOG_DEATH
    LOG_LIFE
    LOG_GATEWAY
    LOG_MSG
    LOG_INFO        ' special; on by default, for both debug & release builds
    LOG_TERRAIN
    LOG_FEATURE
    LOG_FATAL       ' special; on by default, for both debug & release builds
    LOG_INPUT       ' mouse / keyboard events
    LOG_POPUP       ' special, on by default, for both debug & release builds (used for OS dependent popup code)
    LOG_CONSOLE     ' send console messages to file
    LOG_ACTIVITY
    LOG_RESEARCH
    LOG_SAVEGAME
    LOG_REPAIRS
    LOG_LOBBY
    LOG_LAST        ' _must_ be last!
End Enum

' debug.h's callback function pointers, as numbers (CB_<function>; CB_NULL = nullptr)
Public Enum debug_callback_id
    CB_NULL = 0
    CB_debug_callback_stderr = 1
    CB_debug_callback_file = 2
    CB_debug_callback_file_init = 3
    CB_debug_callback_file_exit = 4
    CB_debug_callback_capture = 5   ' (VB) keeps the lines for debug_capture_take (test runs)
End Enum

' frame.h's player counts
Public Enum frame_h_players
    MAX_PLAYERS = 11            ' Maximum number of players in the game.
    MAX_PLAYER_SLOTS = 13       ' Max players plus 1 baba and 1 reserved for features.
End Enum

' frame.h
Public Enum video_backend
    video_backend_opengl = 0
    video_backend_opengles = 1
    video_backend_vulkan = 2
    video_backend_num_backends = 3  ' Must be last!
End Enum

Public Enum WINDOW_MODE
    WINDOW_MODE_windowed = 0
    WINDOW_MODE_fullscreen = 1
End Enum

' wzapp.h
Public Enum DialogType
    Dialog_Error
    Dialog_Warning
    Dialog_Information
End Enum

' physfs_ext.h
Public Enum physfs_mount_order
    PHYSFS_PREPEND = 0
    PHYSFS_APPEND = 1
End Enum

' json.hpp: nlohmann::detail::value_t (nlohmann_json.type_)
Public Enum value_t
    value_t_null = 0
    value_t_object = 1
    value_t_array = 2
    value_t_string = 3
    value_t_boolean = 4
    value_t_number_integer = 5
    value_t_number_unsigned = 6
    value_t_number_float = 7
    value_t_binary = 8
    value_t_discarded = 9
End Enum

' json.hpp: basic_json::error_handler_t (dump's handling of invalid text)
Public Enum error_handler_t
    error_handler_t_strict = 0
    error_handler_t_replace = 1
    error_handler_t_ignore = 2
End Enum

' wzconfig.h: WzConfig::warning
Public Enum WzConfig_warning
    WzConfig_ReadAndWrite = 0
    WzConfig_ReadOnly = 1
    WzConfig_ReadOnlyAndRequired = 2
End Enum

' physfs.h: PHYSFS_ErrorCode
Public Enum PHYSFS_ErrorCode
    PHYSFS_ERR_OK = 0
    PHYSFS_ERR_OTHER_ERROR = 1
    PHYSFS_ERR_OUT_OF_MEMORY = 2
    PHYSFS_ERR_NOT_INITIALIZED = 3
    PHYSFS_ERR_IS_INITIALIZED = 4
    PHYSFS_ERR_ARGV0_IS_NULL = 5
    PHYSFS_ERR_UNSUPPORTED = 6
    PHYSFS_ERR_PAST_EOF = 7
    PHYSFS_ERR_FILES_STILL_OPEN = 8
    PHYSFS_ERR_INVALID_ARGUMENT = 9
    PHYSFS_ERR_NOT_MOUNTED = 10
    PHYSFS_ERR_NOT_FOUND = 11
    PHYSFS_ERR_SYMLINK_FORBIDDEN = 12
    PHYSFS_ERR_NO_WRITE_DIR = 13
    PHYSFS_ERR_OPEN_FOR_READING = 14
    PHYSFS_ERR_OPEN_FOR_WRITING = 15
    PHYSFS_ERR_NOT_A_FILE = 16
    PHYSFS_ERR_READ_ONLY = 17
    PHYSFS_ERR_CORRUPT = 18
    PHYSFS_ERR_SYMLINK_LOOP = 19
    PHYSFS_ERR_IO = 20
    PHYSFS_ERR_PERMISSION = 21
    PHYSFS_ERR_NO_SPACE = 22
    PHYSFS_ERR_BAD_FILENAME = 23
    PHYSFS_ERR_BUSY = 24
    PHYSFS_ERR_DIR_NOT_EMPTY = 25
    PHYSFS_ERR_OS_ERROR = 26
    PHYSFS_ERR_DUPLICATE = 27
    PHYSFS_ERR_BAD_PASSWORD = 28
    PHYSFS_ERR_APP_CALLBACK = 29
End Enum

' physfs.h: PHYSFS_FileType
Public Enum PHYSFS_FileType
    PHYSFS_FILETYPE_REGULAR = 0
    PHYSFS_FILETYPE_DIRECTORY = 1
    PHYSFS_FILETYPE_SYMLINK = 2
    PHYSFS_FILETYPE_OTHER = 3
End Enum

' Name and version of this DLL, for the log header.
Public Function FrameworkBuildInfo() As String
    FrameworkBuildInfo = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
End Function

'--------------------------------------------------------------------------
' debug.h / debug.cpp (see docs/CONVENTIONS.md, "Logging and asserts")

' The debug(part, ...) macro: debug_ <C++ line>, <part>, "<function>", fmt, args...
Public Sub debug_(ByVal line_ As Long, ByVal part As code_part, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    If debug_c.enabled_debug(part) Then
        a = args
        debug_c.debugV line_, part, function_, str_, a
    End If
End Sub

' _debug(): logs whether or not the part is enabled
Public Sub debug__(ByVal line_ As Long, ByVal part As code_part, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    debug_c.debugV line_, part, function_, str_, a
End Sub

Public Sub wz_info(ByVal line_ As Long, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    debug_c.debugV line_, LOG_INFO, function_, str_, a
End Sub

Public Sub debug_multiline(ByVal line_ As Long, ByVal part As code_part, ByVal function_ As String, ByVal multilineString As String)
    debug_c.debug_multiline line_, part, function_, multilineString
End Sub

' _debugFromGfxCallback()
Public Sub debugFromGfxCallback_(ByVal line_ As Long, ByVal part As code_part, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    debug_c.debugFromGfxCallback_ line_, part, function_, str_, a
End Sub

' ASSERT/ASSERT_OR_RETURN failed: ASSERT_FAILURE "<C++ expression>", "<file>:<line>", "<function>", fmt, args...
Public Sub ASSERT_FAILURE(ByVal expr_string As String, ByVal location_description As String, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    debug_c.ASSERT_FAILUREV expr_string, location_description, function_, str_, a
End Sub

' ASSERT_HELPER with a location and function passed on failed (checkDroid and the like):
' ASSERT_HELPER_FAILURE <line of the ASSERT_HELPER>, "<C++ expression>", location, function, fmt, args...
Public Sub ASSERT_HELPER_FAILURE(ByVal line_ As Long, ByVal expr_string As String, ByVal location_description As String, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    debug_c.ASSERT_HELPER_FAILUREV line_, expr_string, location_description, function_, str_, a
End Sub

' objTrace(id, ...): logs when id is the traced object
Public Sub objTrace(ByVal id As Long, ByVal function_ As String, ByVal str_ As String, ParamArray args() As Variant)
    Dim a As Variant
    If id = debug_c.traceID Then
        a = args
        debug_c.realObjTrace_ id, function_, str_, a
    End If
End Sub

Public Sub objTraceEnable(ByVal id As Long)
    debug_c.traceID = id
End Sub

Public Sub objTraceDisable()
    debug_c.traceID = -1
End Sub

Public Sub debug_init()
    debug_c.debug_init
End Sub

Public Sub debug_exit()
    debug_c.debug_exit
End Sub

Public Sub debugFlushStderr()
    debug_c.debugFlushStderr
End Sub

' vbNullString (C: nullptr) when there is no error waiting
Public Function debugLastError() As String
    debugLastError = debug_c.debugLastError()
End Function

' data: the file name for CB_debug_callback_file_init
Public Sub debug_register_callback(ByVal callback As debug_callback_id, ByVal init As debug_callback_id, ByVal exit_ As debug_callback_id, ByVal data As Variant)
    debug_c.debug_register_callback callback, init, exit_, data
End Sub

Public Function debug_enable_switch(ByVal str_ As String) As Boolean
    debug_enable_switch = debug_c.debug_enable_switch(str_)
End Function

Public Function debugPartEnabled(ByVal codePart As code_part) As Boolean
    debugPartEnabled = debug_c.debugPartEnabled(codePart)
End Function

Public Sub debugDisableAssert()
    debug_c.debugDisableAssert
End Sub

Public Property Get last_called_script_event() As String
    last_called_script_event = debug_c.last_called_script_event
End Property

Public Property Let last_called_script_event(ByVal s As String)
    debug_c.last_called_script_event = Left$(s, 99)
End Property

'--------------------------------------------------------------------------
' wzapp.h, the window-free part (wzapp_win)

Public Function wzGetTicks() As Long
    wzGetTicks = wzapp_win.wzGetTicks()
End Function

' (VB) Scripted test runs: wzGetTicks gives ticks while enabled
Public Sub wzSetTestTicks(ByVal enabled As Boolean, ByVal ticks As Long)
    wzapp_win.wzSetTestTicks enabled, ticks
End Sub

Public Sub wzDelay(ByVal delay As Long)
    wzapp_win.wzDelay delay
End Sub

Public Sub wzDisplayDialog(ByVal typ As DialogType, ByVal title As String, ByVal message As String)
    wzapp_win.wzDisplayDialog typ, title, message
End Sub

' (VB) WzPlatform registers its IFrameworkPlatform; Nothing at shutdown.
Public Sub frameSetPlatform(ByVal p As IFrameworkPlatform)
    wzapp_win.frameSetPlatform p
End Sub

' (VB) Scripted test runs log dialogs instead of showing them.
Public Sub wzSetDialogsSuppressed(ByVal suppressed As Boolean)
    wzapp_win.wzSetDialogsSuppressed suppressed
End Sub

'--------------------------------------------------------------------------
' PhysicsFS (physfs_win)

Public Function PHYSFS_init(ByVal argv0 As String) As Long
    PHYSFS_init = physfs_win.PHYSFS_init(argv0)
End Function

Public Function PHYSFS_deinit() As Long
    PHYSFS_deinit = physfs_win.PHYSFS_deinit()
End Function

Public Function PHYSFS_setWriteDir(ByVal newDir As String) As Long
    PHYSFS_setWriteDir = physfs_win.PHYSFS_setWriteDir(newDir)
End Function

Public Function PHYSFS_getWriteDir() As String
    PHYSFS_getWriteDir = physfs_win.PHYSFS_getWriteDir()
End Function

Public Function PHYSFS_mkdir(ByVal dirName As String) As Long
    PHYSFS_mkdir = physfs_win.PHYSFS_mkdir(dirName)
End Function

Public Function PHYSFS_mount(ByVal newDir As String, ByVal mountPoint As String, ByVal appendToPath As physfs_mount_order) As Long
    PHYSFS_mount = physfs_win.PHYSFS_mount(newDir, mountPoint, appendToPath)
End Function

' file: a PHYSFS_File the search path owns once mounted
Public Function PHYSFS_mountHandle(ByVal file As Long, ByVal newDir As String, ByVal mountPoint As String, ByVal appendToPath As physfs_mount_order) As Long
    PHYSFS_mountHandle = physfs_win.PHYSFS_mountHandle(file, newDir, mountPoint, appendToPath)
End Function

Public Function PHYSFS_getBaseDir() As String
    PHYSFS_getBaseDir = physfs_win.PHYSFS_getBaseDir()
End Function

Public Function PHYSFS_getDirSeparator() As String
    PHYSFS_getDirSeparator = physfs_win.PHYSFS_getDirSeparator()
End Function

Public Sub PHYSFS_permitSymbolicLinks(ByVal allow As Long)
    physfs_win.PHYSFS_permitSymbolicLinks allow
End Sub

Public Function PHYSFS_getPrefDir(ByVal org As String, ByVal app As String) As String
    PHYSFS_getPrefDir = physfs_win.PHYSFS_getPrefDir(org, app)
End Function

Public Function WZ_PHYSFS_getLastError() As String
    WZ_PHYSFS_getLastError = physfs_win.WZ_PHYSFS_getLastError()
End Function

Public Function PHYSFS_getLastErrorCode() As PHYSFS_ErrorCode
    PHYSFS_getLastErrorCode = physfs_win.PHYSFS_getLastErrorCode()
End Function

Public Function PHYSFS_getErrorByCode(ByVal code As PHYSFS_ErrorCode) As String
    PHYSFS_getErrorByCode = physfs_win.PHYSFS_getErrorByCode(code)
End Function

' json.hpp (json_hpp, nlohmann_json): json::parse of a String (as UTF-8) or
' of bytes; json::object(), json::array(); json(value). ordered: an
' nlohmann::ordered_json.
Public Function json_parse(ByVal s As String, Optional ByVal ordered As Boolean = False) As nlohmann_json
    Set json_parse = json_hpp.json_parse(s, ordered)
End Function

Public Function json_parse_bytes(data() As Byte, ByVal first As Long, ByVal size As Long, Optional ByVal ordered As Boolean = False) As nlohmann_json
    Set json_parse_bytes = json_hpp.json_parse_bytes(data, first, size, ordered)
End Function

Public Function json_object(Optional ByVal ordered As Boolean = False) As nlohmann_json
    Set json_object = json_hpp.json_object(ordered)
End Function

Public Function json_array(Optional ByVal ordered As Boolean = False) As nlohmann_json
    Set json_array = json_hpp.json_array(ordered)
End Function

Public Function json_value(ByVal v As Variant, Optional ByVal ordered As Boolean = False) As nlohmann_json
    Set json_value = json_hpp.json_value(v, ordered)
End Function

Public Function json_value_unsigned(ByVal v As Variant, Optional ByVal ordered As Boolean = False) As nlohmann_json
    Set json_value_unsigned = json_hpp.json_value_unsigned(v, ordered)
End Function

' wzconfig.cpp: WzConfig(name, warning), json_variant(value) (Empty: json_variant()), json_getValue
Public Function WzConfig_new(ByVal name_ As String, ByVal warning As WzConfig_warning) As WzConfig
    Set WzConfig_new = wzconfig_c.WzConfig_new(name_, warning)
End Function

Public Function json_variant_new(Optional ByVal v As Variant) As json_variant
    Set json_variant_new = wzconfig_c.json_variant_new(v)
End Function

Public Function json_getValue(ByVal json As nlohmann_json, ByVal key As String, Optional ByVal defaultValue As json_variant = Nothing) As json_variant
    Set json_getValue = wzconfig_c.json_getValue(json, key, defaultValue)
End Function

Public Function json_getValue_idx(ByVal json As nlohmann_json, ByVal idx As Long, Optional ByVal defaultValue As json_variant = Nothing) As json_variant
    Set json_getValue_idx = wzconfig_c.json_getValue_idx(json, idx, defaultValue)
End Function

' wzjsonhelpers.cpp: optional<nlohmann::json>, Nothing for nullopt
Public Function wzLoadJsonObjectFromFile(ByVal filePath As String, Optional ByVal quiet As Boolean = False) As nlohmann_json
    Set wzLoadJsonObjectFromFile = wzjsonhelpers_c.wzLoadJsonObjectFromFile(filePath, quiet)
End Function

' wzpaths.cpp
Public Function getWZInstallPrefix() As String
    getWZInstallPrefix = wzpaths_c.getWZInstallPrefix()
End Function

Public Function PHYSFS_delete(ByVal fname As String) As Long
    PHYSFS_delete = physfs_win.PHYSFS_delete(fname)
End Function

Public Function PHYSFS_unmount(ByVal oldDir As String) As Long
    PHYSFS_unmount = physfs_win.PHYSFS_unmount(oldDir)
End Function

Public Function PHYSFS_getSearchPath(out() As String) As Long
    PHYSFS_getSearchPath = physfs_win.PHYSFS_getSearchPath(out)
End Function

Public Function PHYSFS_getMountPoint(ByVal dir_ As String) As String
    PHYSFS_getMountPoint = physfs_win.PHYSFS_getMountPoint(dir_)
End Function

Public Function PHYSFS_openRead(ByVal fname As String) As Long
    PHYSFS_openRead = physfs_win.PHYSFS_openRead(fname)
End Function

Public Function PHYSFS_openWrite(ByVal fname As String) As Long
    PHYSFS_openWrite = physfs_win.PHYSFS_openWrite(fname)
End Function

Public Function PHYSFS_openAppend(ByVal fname As String) As Long
    PHYSFS_openAppend = physfs_win.PHYSFS_openAppend(fname)
End Function

Public Function PHYSFS_readBytes(ByVal f As Long, buffer() As Byte, ByVal ofs As Long, ByVal length As Long) As Long
    PHYSFS_readBytes = physfs_win.PHYSFS_readBytes(f, buffer, ofs, length)
End Function

Public Function PHYSFS_writeBytes(ByVal f As Long, buffer() As Byte, ByVal ofs As Long, ByVal length As Long) As Long
    PHYSFS_writeBytes = physfs_win.PHYSFS_writeBytes(f, buffer, ofs, length)
End Function

Public Function PHYSFS_fileLength(ByVal f As Long) As Double
    PHYSFS_fileLength = physfs_win.PHYSFS_fileLength(f)
End Function

Public Function PHYSFS_tell(ByVal f As Long) As Double
    PHYSFS_tell = physfs_win.PHYSFS_tell(f)
End Function

Public Function PHYSFS_seek(ByVal f As Long, ByVal pos As Double) As Long
    PHYSFS_seek = physfs_win.PHYSFS_seek(f, pos)
End Function

Public Function PHYSFS_eof(ByVal f As Long) As Long
    PHYSFS_eof = physfs_win.PHYSFS_eof(f)
End Function

Public Function PHYSFS_flush(ByVal f As Long) As Long
    PHYSFS_flush = physfs_win.PHYSFS_flush(f)
End Function

Public Function PHYSFS_close(ByVal f As Long) As Long
    PHYSFS_close = physfs_win.PHYSFS_close(f)
End Function

Public Function PHYSFS_exists(ByVal fname As String) As Long
    PHYSFS_exists = physfs_win.PHYSFS_exists(fname)
End Function

Public Function PHYSFS_stat(ByVal fname As String, filesize As Double, modtime As Double, filetype As Long, readonly_ As Long) As Long
    PHYSFS_stat = physfs_win.PHYSFS_stat(fname, filesize, modtime, filetype, readonly_)
End Function

Public Function WZ_PHYSFS_isDirectory(ByVal fname As String) As Long
    WZ_PHYSFS_isDirectory = physfs_win.WZ_PHYSFS_isDirectory(fname)
End Function

Public Function PHYSFS_getRealDir(ByVal fname As String) As String
    PHYSFS_getRealDir = physfs_win.PHYSFS_getRealDir(fname)
End Function

Public Function PHYSFS_enumerateFiles(ByVal dirPath As String, out() As String) As Long
    PHYSFS_enumerateFiles = physfs_win.PHYSFS_enumerateFiles(dirPath, out)
End Function

'--------------------------------------------------------------------------
' i18n.cpp

Public Function getCompileDate() As String
    getCompileDate = i18n_c.getCompileDate()
End Function

Public Function getLanguage() As String
    getLanguage = i18n_c.getLanguage()
End Function

Public Function getLanguageName() As String
    getLanguageName = i18n_c.getLanguageName()
End Function

Public Function setLanguage(ByVal language As String) As Boolean
    setLanguage = i18n_c.setLanguage(language)
End Function

' C: std::vector<locale_info> getLocales(): code(0 ..) and nm(0 ..), and the count
Public Function getLocales(code() As String, nm() As String) As Long
    getLocales = i18n_c.getLocales(code, nm)
End Function

Public Sub initI18n()
    i18n_c.initI18n
End Sub

'--------------------------------------------------------------------------
' frame.cpp

Public Property Get selectedPlayer() As Long
    selectedPlayer = frame_c.selectedPlayer
End Property

Public Property Let selectedPlayer(ByVal player As Long)
    frame_c.selectedPlayer = player
End Property

Public Property Get realSelectedPlayer() As Long
    realSelectedPlayer = frame_c.realSelectedPlayer
End Property

Public Property Let realSelectedPlayer(ByVal player As Long)
    frame_c.realSelectedPlayer = player
End Property

Public Function frameInitialise() As Boolean
    frameInitialise = frame_c.frameInitialise()
End Function

Public Sub frameUpdate()
    frame_c.frameUpdate
End Sub

Public Sub frameShutDown()
    frame_c.frameShutDown
End Sub

Public Function frameRate() As Long
    frameRate = frame_c.frameRate()
End Function

Public Function frameGetFrameNumber() As Long
    frameGetFrameNumber = frame_c.frameGetFrameNumber()
End Function

Public Sub setMouseWarp(ByVal value As Boolean)
    frame_c.setMouseWarp value
End Sub

Public Function getMouseWarp() As Boolean
    getMouseWarp = frame_c.getMouseWarp()
End Function

Public Function openLoadFile(ByVal fileName As String, ByVal hard_fail As Boolean) As Long
    openLoadFile = frame_c.openLoadFile(fileName, hard_fail)
End Function

Public Function openSaveFile(ByVal fileName As String) As Long
    openSaveFile = frame_c.openSaveFile(fileName)
End Function

' ppFileData gets the file and a terminating 0; pFileSize its size.
Public Function loadFile(ByVal pFileName As String, ppFileData() As Byte, pFileSize As Long, Optional ByVal hard_fail As Boolean = True) As Boolean
    loadFile = frame_c.loadFile(pFileName, ppFileData, pFileSize, hard_fail)
End Function

Public Function loadFileToBuffer(ByVal pFileName As String, pFileBuffer() As Byte, ByVal bufferSize As Long, pSize As Long) As Boolean
    loadFileToBuffer = frame_c.loadFileToBuffer(pFileName, pFileBuffer, bufferSize, pSize)
End Function

Public Function loadFileToBufferNoError(ByVal pFileName As String, pFileBuffer() As Byte, ByVal bufferSize As Long, pSize As Long) As Boolean
    loadFileToBufferNoError = frame_c.loadFileToBufferNoError(pFileName, pFileBuffer, bufferSize, pSize)
End Function

' std::vector<char>: outputBuffer and its outputSize
Public Function loadFileToBufferVector(ByVal pFileName As String, outputBuffer() As Byte, outputSize As Long, ByVal hard_fail As Boolean, ByVal appendNullCharacter As Boolean) As Boolean
    loadFileToBufferVector = frame_c.loadFileToBufferVector(pFileName, outputBuffer, outputSize, hard_fail, appendNullCharacter)
End Function

Public Function saveFile(ByVal pFileName As String, pFileData() As Byte, ByVal fileSize As Long) As Boolean
    saveFile = frame_c.saveFile(pFileName, pFileData, fileSize)
End Function

Public Function PHYSFS_printf(ByVal file As Long, ByVal format_ As String, ParamArray args() As Variant) As Boolean
    Dim a As Variant
    a = args
    PHYSFS_printf = frame_c.PHYSFS_printfV(file, format_, a)
End Function

Public Function video_backend_from_str(ByVal str_ As String, output_backend As Long) As Boolean
    video_backend_from_str = frame_c.video_backend_from_str(str_, output_backend)
End Function

Public Function to_string_video_backend(ByVal backend As video_backend) As String
    to_string_video_backend = frame_c.to_string_video_backend(backend)
End Function

Public Function to_display_string_video_backend(ByVal backend As video_backend) As String
    to_display_string_video_backend = frame_c.to_display_string_video_backend(backend)
End Function

Public Function to_display_string_WINDOW_MODE(ByVal mode As WINDOW_MODE) As String
    to_display_string_WINDOW_MODE = frame_c.to_display_string_WINDOW_MODE(mode)
End Function

' (VB) The exe registers its IFrameworkGame; Nothing at shutdown.
Public Sub frameSetGame(ByVal g As IFrameworkGame)
    frame_c.frameSetGame g
End Sub

'--------------------------------------------------------------------------
' frameresource.h

Public Sub resSetLoadCallback(ByVal funcToCall As Long)
    frameresource_c.resSetLoadCallback funcToCall
End Sub

Public Sub resDoResLoadCallback()
    frameresource_c.resDoResLoadCallback
End Sub

Public Sub resSetBaseDir(ByVal pResDir As String)
    frameresource_c.resSetBaseDir pResDir
End Sub

Public Sub resForceBaseDir(ByVal pResDir As String)
    frameresource_c.resForceBaseDir pResDir
End Sub

Public Function resLoad(ByVal pResFile As String, ByVal blockID As Long) As Boolean
    resLoad = frameresource_c.resLoad(pResFile, blockID)
End Function

Public Sub resReleaseAll()
    frameresource_c.resReleaseAll
End Sub

Public Sub resReleaseBlockData(ByVal blockID As Long)
    frameresource_c.resReleaseBlockData blockID
End Sub

Public Sub resReleaseAllData()
    frameresource_c.resReleaseAllData
End Sub

' buffLoad, release: the game's function numbers (IFrameworkGame), 0 for nullptr
Public Function resAddBufferLoad(ByVal pType As String, ByVal buffLoad As Long, ByVal release As Long) As Boolean
    resAddBufferLoad = frameresource_c.resAddBufferLoad(pType, buffLoad, release)
End Function

Public Function resAddFileLoad(ByVal pType As String, ByVal fileLoad As Long, ByVal release As Long) As Boolean
    resAddFileLoad = frameresource_c.resAddFileLoad(pType, fileLoad, release)
End Function

Public Function resLoadFile(ByVal pType As String, ByVal pFile As String) As Boolean
    resLoadFile = frameresource_c.resLoadFile(pType, pFile)
End Function

Public Function resGetDataFromHash(ByVal pType As String, ByVal HashedID As Long) As Variant
    Dim v As Variant
    VarAssign v, frameresource_c.resGetDataFromHash(pType, HashedID)
    If IsObject(v) Then Set resGetDataFromHash = v Else resGetDataFromHash = v
End Function

Public Function resGetData(ByVal pType As String, ByVal pID As String) As Variant
    Dim v As Variant
    VarAssign v, frameresource_c.resGetData(pType, pID)
    If IsObject(v) Then Set resGetData = v Else resGetData = v
End Function

Public Function resPresent(ByVal pType As String, ByVal pID As String) As Boolean
    resPresent = frameresource_c.resPresent(pType, pID)
End Function

Public Sub resToLower(pStr As String)
    frameresource_c.resToLower pStr
End Sub

Public Function resGetHashfromData(ByVal pType As String, pData As Variant, pHash As Long) As Boolean
    resGetHashfromData = frameresource_c.resGetHashfromData(pType, pData, pHash)
End Function

Public Function resGetNamefromData(ByVal typ As String, data As Variant) As String
    resGetNamefromData = frameresource_c.resGetNamefromData(typ, data)
End Function

Public Function GetLastResourceFilename() As String
    GetLastResourceFilename = frameresource_c.GetLastResourceFilename()
End Function

Public Sub SetLastResourceFilename(ByVal pName As String)
    frameresource_c.SetLastResourceFilename pName
End Sub

'--------------------------------------------------------------------------
' physfs_ext.h: the names in files(0 .. count - 1) (the callback is the caller's loop)

Public Function WZ_PHYSFS_enumerateFiles(ByVal dirPath As String, files() As String, count As Long) As Boolean
    WZ_PHYSFS_enumerateFiles = physfs_ext_c.WZ_PHYSFS_enumerateFiles(dirPath, files, count)
End Function

Public Function WZ_PHYSFS_enumerateFolders(ByVal dirPath As String, folders() As String, count As Long) As Boolean
    WZ_PHYSFS_enumerateFolders = physfs_ext_c.WZ_PHYSFS_enumerateFolders(dirPath, folders, count)
End Function

' physfs_ext.h: PHYSFS_readBytes / PHYSFS_writeBytes as a PHYSFS_sint64
Public Function WZ_PHYSFS_readBytes(ByVal handle As Long, buf() As Byte, ByVal first As Long, ByVal size As Long) As Double
    WZ_PHYSFS_readBytes = physfs_ext_c.WZ_PHYSFS_readBytes(handle, buf, first, size)
End Function

Public Function WZ_PHYSFS_writeBytes(ByVal handle As Long, buf() As Byte, ByVal first As Long, ByVal size As Long) As Double
    WZ_PHYSFS_writeBytes = physfs_ext_c.WZ_PHYSFS_writeBytes(handle, buf, first, size)
End Function

Public Function filenameEndWithExtension(ByVal filename As String, ByVal extension As String) As Boolean
    filenameEndWithExtension = physfs_ext_c.filenameEndWithExtension(filename, extension)
End Function

' WZ_PHYSFS_cleanupOldFilesInFolder(path, extension, fileLimit, deleteFileFunction): up to the deleting, which is
' the caller's loop with its deleteFileFunction over oldest(0 .. count - 1) (see physfs_ext_c); -1, 0 or 1
Public Function WZ_PHYSFS_cleanupOldFilesInFolder(ByVal path As String, ByVal extension As String, ByVal fileLimit As Long, nfiles As Long, oldest() As String, count As Long) As Long
    WZ_PHYSFS_cleanupOldFilesInFolder = physfs_ext_c.WZ_PHYSFS_cleanupOldFilesInFolder(path, extension, fileLimit, nfiles, oldest, count)
End Function

' physfs.h: 32-bit values in big-endian order; nonzero on success
Public Function PHYSFS_writeSBE32(ByVal f As Long, ByVal value As Long) As Long
    PHYSFS_writeSBE32 = physfs_win.PHYSFS_writeSBE32(f, value)
End Function

Public Function PHYSFS_writeUBE32(ByVal f As Long, ByVal value As Long) As Long
    PHYSFS_writeUBE32 = physfs_win.PHYSFS_writeUBE32(f, value)
End Function

Public Function PHYSFS_readSBE32(ByVal f As Long, value As Long) As Long
    PHYSFS_readSBE32 = physfs_win.PHYSFS_readSBE32(f, value)
End Function

Public Function PHYSFS_readUBE32(ByVal f As Long, value As Long) As Long
    PHYSFS_readUBE32 = physfs_win.PHYSFS_readUBE32(f, value)
End Function

'--------------------------------------------------------------------------
' (VB) the lines CB_debug_callback_capture kept since the last call, each ending with a newline

Public Function debug_capture_take() As String
    debug_capture_take = debug_c.debug_capture_take()
End Function

'--------------------------------------------------------------------------
' strres.h (struct STR_RES * is the class STR_RES; a missing string is vbNullString)

Public Function strresCreate() As STR_RES
    Set strresCreate = strres_c.strresCreate()
End Function

Public Sub strresDestroy(ByVal psRes As STR_RES)
    strres_c.strresDestroy psRes
End Sub

Public Function strresGetString(ByVal psRes As STR_RES, ByVal ID As String) As String
    strresGetString = strres_c.strresGetString(psRes, ID)
End Function

Public Function strresLoad(ByVal psRes As STR_RES, ByVal fileName As String) As Boolean
    strresLoad = strres_c.strresLoad(psRes, fileName)
End Function

Public Function strresGetIDfromString(ByVal psRes As STR_RES, ByVal pString As String) As String
    strresGetIDfromString = strres_c.strresGetIDfromString(psRes, pString)
End Function
