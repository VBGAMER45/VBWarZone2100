Attribute VB_Name = "wzphysfszipioprovider_c"
' src/wzphysfszipioprovider.cpp: WzZipIOPHYSFSSourceReadProvider::make, and the ZipIOProvider plugin's
' WzMapZipIO::openZipArchiveReadIOProvider (lib/wzmaplib/plugins/ZipIOProvider/src/ZipIOProvider.cpp): a .wz map
' package opened for wzmaplib (WzMapZipIO.cls: PhysFS's zip archiver in libzip's place).
Option Explicit

Private nextZipIOUid As Long                    ' (VB) the WzMapZipIO mounts' unique names

' static std::shared_ptr<WzZipIOPHYSFSSourceReadProvider> make(const std::string& path)
Public Function WzZipIOPHYSFSSourceReadProvider_make(ByVal path As String) As WzZipIOPHYSFSSourceReadProvider
    Dim result As WzZipIOPHYSFSSourceReadProvider
    Set result = New WzZipIOPHYSFSSourceReadProvider
    If Not result.init(path) Then
        Set WzZipIOPHYSFSSourceReadProvider_make = Nothing
        Exit Function
    End If
    Set WzZipIOPHYSFSSourceReadProvider_make = result
End Function

' static std::shared_ptr<WzMapZipIO> openZipArchiveReadIOProvider(std::shared_ptr<WzZipIOSourceReadProvider>
' zipSourceProvider, WzMap::LoggingProtocol* pCustomLogger = nullptr, bool extraConsistencyChecks = false): Nothing
' when the source is not a zip archive (logged through pCustomLogger with PhysFS's error, as libzip's)
Public Function WzMapZipIO_openZipArchiveReadIOProvider(ByVal zipSourceProvider As WzZipIOPHYSFSSourceReadProvider, _
        Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As WzMapZipIO
    Dim result As WzMapZipIO, pErrorStr As String
    If zipSourceProvider Is Nothing Then
        Set WzMapZipIO_openZipArchiveReadIOProvider = Nothing
        Exit Function
    End If
    nextZipIOUid = nextZipIOUid + 1
    Set result = New WzMapZipIO
    If Not result.openFrom(zipSourceProvider, nextZipIOUid) Then
        ' Failed to open from source
        If Not pCustomLogger Is Nothing Then
            pErrorStr = WZ_PHYSFS_getLastError()
            If Len(pErrorStr) = 0 Then pErrorStr = "<n/a>"
            pCustomLogger.printLog LogLevel_Error, "openZipArchiveReadIOProvider", 540, pErrorStr
        End If
        Set WzMapZipIO_openZipArchiveReadIOProvider = Nothing
        Exit Function
    End If
    Set WzMapZipIO_openZipArchiveReadIOProvider = result
End Function
