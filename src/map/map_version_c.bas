Attribute VB_Name = "map_version_c"
' lib/wzmaplib/src/map_version.cpp, with the values CMake writes into
' wzmaplib_internal/map_config_internal.h (lib/wzmaplib/CMakeLists.txt:
' WZMAPLIB_VERSION_MAJOR 1, _MINOR 4, _REV 6).
Option Explicit

Private Const WZMAPLIB_VERSION_MAJOR_ As Long = 1
Private Const WZMAPLIB_VERSION_MINOR_ As Long = 4
Private Const WZMAPLIB_VERSION_REV_ As Long = 6
Private Const WZMAPLIB_VERSION_STRING_ As String = "1.4.6"

Public Function wzmaplib_version_string() As String
    wzmaplib_version_string = WZMAPLIB_VERSION_STRING_
End Function

Public Function wzmaplib_version_major() As Long
    wzmaplib_version_major = WZMAPLIB_VERSION_MAJOR_
End Function

Public Function wzmaplib_version_minor() As Long
    wzmaplib_version_minor = WZMAPLIB_VERSION_MINOR_
End Function

Public Function wzmaplib_version_rev() As Long
    wzmaplib_version_rev = WZMAPLIB_VERSION_REV_
End Function
