Attribute VB_Name = "data_c"
' src/data.cpp, data.h: the data loading functions the resource module
' (lib/framework/frameresource.cpp) calls for each type of a .wrf file.
' Their function pointers are numbers (enum RES_FUNCTION): dataInitLoadFuncs
' registers them, and the framework calls them back through FrameworkGame
' (IFrameworkGame), which calls dataBufferLoad, dataFileLoad and
' dataRelease here. void **ppData is a Variant (Empty: nullptr; (void *)1
' is 1). The calcDataHash overload for a WzConfig is calcDataHash_ini.
' What the loaders call is partly not ported yet: the images and sounds,
' the scripts, the view data (stubs_game).
Option Explicit

' data.h: used for multiplayer data integrity check.
Public Const DATA_SWEAPON As Long = 0
Public Const DATA_SBODY As Long = 1
Public Const DATA_SCONSTR As Long = 2
Public Const DATA_SECM As Long = 3
Public Const DATA_SPROP As Long = 4
Public Const DATA_SSENSOR As Long = 5
Public Const DATA_SREPAIR As Long = 6
Public Const DATA_SBRAIN As Long = 7
Public Const DATA_SPROPTY As Long = 8
Public Const DATA_STERRT As Long = 9
Public Const DATA_SWEAPMOD As Long = 10
Public Const DATA_STEMP As Long = 11
Public Const DATA_SSTRUCT As Long = 12
Public Const DATA_SSTRWEAP As Long = 13
Public Const DATA_STRFUNC As Long = 14
Public Const DATA_SSTRMOD As Long = 15
Public Const DATA_SFEAT As Long = 16
Public Const DATA_SFUNC As Long = 17
Public Const DATA_RESCH As Long = 18
Public Const DATA_RPREREQ As Long = 19
Public Const DATA_RCOMPRED As Long = 20
Public Const DATA_RSTRREQ As Long = 21
Public Const DATA_RCOMPRES As Long = 22
Public Const DATA_RSTRRED As Long = 23
Public Const DATA_RSTRRES As Long = 24
Public Const DATA_RFUNC As Long = 25
Public Const DATA_SCRIPT As Long = 26
Public Const DATA_SCRIPTVAL As Long = 27
Public Const DATA_MAXDATA As Long = 28      ' must be last

' The functions of the resource types (RES_BUFFERLOAD, RES_FILELOAD, RES_FREE): 0 is nullptr
Public Enum RES_FUNCTION
    fn_bufferSMSGLoad = 1
    fn_dataSMSGRelease
    fn_bufferSFEATLoad
    fn_dataSFEATRelease
    fn_bufferSTEMPLLoad
    fn_dataSTEMPLRelease
    fn_dataAudioLoad
    fn_sound_ReleaseTrack
    fn_bufferSWEAPONLoad
    fn_dataReleaseStats
    fn_bufferSBPIMDLoad
    fn_bufferSBRAINLoad
    fn_bufferSSENSORLoad
    fn_bufferSECMLoad
    fn_bufferSREPAIRLoad
    fn_bufferSCONSTRLoad
    fn_bufferSPROPLoad
    fn_bufferSPROPTYPESLoad
    fn_bufferSTERRTABLELoad
    fn_bufferSBODYLoad
    fn_bufferSWEAPMODLoad
    fn_bufferSPROPSNDLoad
    fn_dataAudioCfgLoad
    fn_dataImageLoad
    fn_dataImageRelease
    fn_dataTERTILESLoad
    fn_dataIMGLoad
    fn_dataIMGRelease
    fn_dataStrResLoad
    fn_dataStrResRelease
    fn_dataResearchMsgLoad
    fn_bufferSSTRMODLoad
    fn_jsLoad
    fn_bufferSSTRUCTLoad
    fn_dataSSTRUCTRelease
    fn_bufferRESCHLoad
    fn_dataRESCHRelease
    fn_dataProximityMsgLoad
    fn_dataFlicMsgLoad
End Enum

' whether a save game is currently being loaded
Private saveFlag As Boolean

Public DataHash(0 To DATA_MAXDATA - 1) As Long  ' uint32_t

'
'   hashBuffer()
'   \param pData pointer to our buffer (always a text file)
'   \param size the size of the buffer
'   \return hash calculated from the buffer.
'
'   Note, this is obviously not a very complex hash routine.  This most likely has many collisions possible.
'   This is almost the same routine that Pumpkin had, minus the ugly bug :)
'   And minus the old algorithm and debugging trace, replaced with a simple CRC...
'
Private Sub hashBuffer(pData() As Byte, ByVal size As Long, crc As Long)
    Dim nl(0 To 0) As Byte, i As Long, j As Long, lines As Long, bytes As Long
    nl(0) = 10
    lines = 0
    bytes = 0
    i = 0
    Do While i < size
        j = i
        Do While j < size
            If pData(j) = 10 Or pData(j) = 13 Then Exit Do
            j = j + 1
        Loop

        If i <> j Then  ' CRC non-empty lines only.
            crc = wz_crc_update(crc, pData, i, j - i)  ' CRC the line.
            crc = wz_crc_update(crc, nl, 0, 1)         ' CRC the line ending.

            lines = lines + 1
            bytes = bytes + j - i + 1
        End If
        i = j + 1
    Loop
    debug_ 86, LOG_NET, "hashBuffer", "The size of the old buffer (%u bytes - %d stripped), New buffer size of %u bytes, %u non-empty lines.", size, size - bytes, bytes, lines
End Sub

' create the hash for that data block.
' Data should be converted to Network byte order
Public Sub calcDataHash(pBuffer() As Byte, ByVal size As Long, ByVal index As Long)
    If Not bMultiPlayer Then
        Exit Sub
    End If

    hashBuffer pBuffer, size, DataHash(index)

    debug_ 100, LOG_NET, "calcDataHash", "DataHash[%2u] = %08x", index, DataHash(index)
End Sub

' calcDataHash(const WzConfig &ini, uint32_t index)
Private Sub calcDataHash_ini(ByVal ini As WzConfig, ByVal index As Long)
    Dim jsonDump As String, b() As Byte, n As Long
    jsonDump = ini.compactStringRepresentation()
    n = utf8_encode(jsonDump, b)
    calcDataHash b, n, index
End Sub

Public Sub resetDataHash()
    Dim i As Long
    For i = 0 To DATA_MAXDATA - 1
        DataHash(i) = wz_crc_init()
    Next
    debug_ 118, LOG_NET, "resetDataHash", "== Hash is reset =="
End Sub

'*********************************************************

Public Sub dataSetSaveFlag()
    saveFlag = True
End Sub

Public Sub dataClearSaveFlag()
    saveFlag = False
End Sub

' Load the body stats
Private Function bufferSBODYLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SBODY

    If Not loadBodyStats(ini) Then Exit Function
    If Not allocComponentList(COMP_BODY, asBodyStats_size) Then Exit Function
    ppData = 1&
    bufferSBODYLoad = True
End Function

Private Sub dataReleaseStats(pData As Variant)
    ' FIXME, huge hack, is called many times!
    freeComponentLists
    statsShutDown
End Sub

' Load the weapon stats
Private Function bufferSWEAPONLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SWEAPON

    If Not loadWeaponStats(ini) Then Exit Function
    If Not allocComponentList(COMP_WEAPON, asWeaponStats_size) Then Exit Function

    ' not interested in this value
    ppData = Empty
    bufferSWEAPONLoad = True
End Function

' Load the constructor stats
Private Function bufferSCONSTRLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SCONSTR

    If Not loadConstructStats(ini) Then Exit Function
    If Not allocComponentList(COMP_CONSTRUCT, asConstructStats_size) Then Exit Function

    'not interested in this value
    ppData = Empty
    bufferSCONSTRLoad = True
End Function

' Load the ECM stats
Private Function bufferSECMLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SECM

    If Not loadECMStats(ini) Then Exit Function
    If Not allocComponentList(COMP_ECM, asECMStats_size) Then Exit Function

    'not interested in this value
    ppData = Empty
    bufferSECMLoad = True
End Function

' Load the Propulsion stats
Private Function bufferSPROPLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SPROP

    If Not loadPropulsionStats(ini) Then Exit Function
    If Not allocComponentList(COMP_PROPULSION, asPropulsionStats_size) Then Exit Function

    'not interested in this value
    ppData = 1&
    bufferSPROPLoad = True
End Function

Private Function bufferSSENSORLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SSENSOR

    If Not loadSensorStats(ini) Then Exit Function
    If Not allocComponentList(COMP_SENSOR, asSensorStats_size) Then Exit Function

    'not interested in this value
    ppData = Empty
    bufferSSENSORLoad = True
End Function

' Load the Repair stats
Private Function bufferSREPAIRLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SREPAIR

    If Not loadRepairStats(ini) Then Exit Function
    If Not allocComponentList(COMP_REPAIRUNIT, asRepairStats_size) Then Exit Function

    'not interested in this value
    ppData = Empty
    bufferSREPAIRLoad = True
End Function

' Load the Brain stats
Private Function bufferSBRAINLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SBRAIN

    If Not loadBrainStats(ini) Then Exit Function
    If Not allocComponentList(COMP_BRAIN, asBrainStats_size) Then Exit Function
    'not interested in this value
    ppData = Empty
    bufferSBRAINLoad = True
End Function

' Load the PropulsionType stats
Private Function bufferSPROPTYPESLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SPROPTY

    If Not loadPropulsionTypes(ini) Then
        Exit Function
    End If

    'not interested in this value
    ppData = Empty
    bufferSPROPTYPESLoad = True
End Function

' Load the propulsion type sound stats
Private Function bufferSPROPSNDLoad(ByVal fileName As String, ppData As Variant) As Boolean
    If Not loadPropulsionSounds(fileName) Then
        Exit Function
    End If

    'not interested in this value
    ppData = Empty
    bufferSPROPSNDLoad = True
End Function

' Load the STERRTABLE stats
Private Function bufferSTERRTABLELoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_STERRT

    If Not loadTerrainTable(ini) Then
        Exit Function
    End If

    'not interested in this value
    ppData = Empty
    bufferSTERRTABLELoad = True
End Function

' Load the body/propulsion IMDs stats -- FIXME, REMOVE IT, NOT USED
Private Function bufferSBPIMDLoad(ByVal fileName As String, ppData As Variant) As Boolean
    'not interested in this value
    ppData = Empty
    bufferSBPIMDLoad = True
End Function

' Load the Weapon Effect modifier stats
Private Function bufferSWEAPMODLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SWEAPMOD

    If Not loadWeaponModifiers(ini) Then
        Exit Function
    End If

    'not interested in this value
    ppData = Empty
    bufferSWEAPMODLoad = True
End Function

' Load the Template stats
Private Function bufferSTEMPLLoad(ByVal fileName As String, ppData As Variant) As Boolean
    If Not loadDroidTemplates(fileName) Then
        Exit Function
    End If

    ' set a dummy value so the release function gets called
    ppData = 1&
    bufferSTEMPLLoad = True
End Function

' release the templates
Private Sub dataSTEMPLRelease(pData As Variant)
    'free the storage allocated to the droid templates
    droidTemplateShutDown
End Sub

' Load the Structure stats
Private Function bufferSSTRUCTLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SSTRUCT

    If Not loadStructureStats(ini) Then
        Exit Function
    End If

    If Not allocStructLists() Then
        Exit Function
    End If

    ' set a dummy value so the release function gets called
    ppData = 1&
    bufferSSTRUCTLoad = True
End Function

' release the structure stats
Private Sub dataSSTRUCTRelease(pData As Variant)
    freeStructureLists
    structureStatsShutDown
End Sub

' Load the Structure strength modifier stats
Private Function bufferSSTRMODLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SSTRMOD

    If Not loadStructureStrengthModifiers(ini) Then
        Exit Function
    End If

    'not interested in this value
    ppData = Empty
    bufferSSTRMODLoad = True
End Function

' Load the Feature stats
Private Function bufferSFEATLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig
    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_SFEAT

    If Not loadFeatureStats(ini) Then
        Exit Function
    End If
    ' set a dummy value so the release function gets called
    ppData = 1&
    bufferSFEATLoad = True
End Function

' free the feature stats
Private Sub dataSFEATRelease(pData As Variant)
    featureStatsShutDown
End Sub

' release the research stats
Private Sub dataRESCHRelease(pData As Variant)
    'free the storage allocated to the stats
    ResearchShutDown
End Sub

' Load the Research stats
Private Function bufferRESCHLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ini As WzConfig, nullData As Variant
    'check to see if already loaded
    If asResearch_size <> 0 Then
        'release previous data before loading in the new
        dataRESCHRelease nullData
    End If

    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    calcDataHash_ini ini, DATA_RESCH

    If Not loadResearch(ini) Then
        Exit Function
    End If

    bufferRESCHLoad = True
End Function

' Load the message viewdata
Private Function bufferSMSGLoad(pBuffer() As Byte, ByVal size As Long, ppData As Variant) As Boolean
    Dim ptr As String
    If Not loadViewData(pBuffer, size, ptr) Then
        Exit Function
    End If

    ' set the pointer so the release function gets called with it
    ppData = ptr
    bufferSMSGLoad = True
End Function

' Load research message viewdata
Private Function dataResearchMsgLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ptr As String
    If Not loadResearchViewData(fileName, ptr) Then
        Exit Function
    End If

    ' set the pointer so the release function gets called with it
    ppData = ptr
    dataResearchMsgLoad = True
End Function

Private Function dataProximityMsgLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ptr As String
    If Not loadProximityViewData(fileName, ptr) Then
        Exit Function
    End If

    ' set the pointer so the release function gets called with it
    ppData = ptr
    dataProximityMsgLoad = True
End Function

Private Function dataFlicMsgLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim ptr As String
    If Not loadFlicViewData(fileName, ptr) Then
        Exit Function
    End If

    ' set the pointer so the release function gets called with it
    ppData = ptr
    dataFlicMsgLoad = True
End Function

' release the message viewdata
Private Sub dataSMSGRelease(pData As Variant)
    If Not (Not IsEmpty(pData)) Then
        ASSERT_FAILURE "pData", "src/data.cpp:509", "dataSMSGRelease", "pData unexpectedly null"
        Exit Sub
    End If
    viewDataShutDown CStr(pData)
End Sub

'
' Load an image from file
'
Private Function dataImageLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim psSprite As iV_Image
    Set psSprite = New iV_Image
    If psSprite Is Nothing Then
        Exit Function
    End If

    If Not iV_loadImage_PNG(fileName, psSprite) Then
        debug_ 528, LOG_ERROR, "dataImageLoad", "IMGPAGE load failed"
        Set psSprite = Nothing
        Exit Function
    End If

    Set ppData = psSprite

    dataImageLoad = True
End Function

' Tertiles (terrain tiles) loader.
Private Function dataTERTILESLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim status As Boolean
    status = texLoad(fileName)
    If Not status Then
        ASSERT_FAILURE "status", "src/data.cpp:545", "dataTERTILESLoad", "Error loading tertiles!"
        Exit Function
    End If
    debug_ 546, LOG_TEXTURE, "dataTERTILESLoad", "HW Tiles loaded"

    ppData = Empty    ' don't bother calling cleanup

    dataTERTILESLoad = True
End Function

' (VB) an IMAGEFILE * is WzGfx's handle (a Long; 0 = nullptr)
Private Function dataIMGLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim img As Long
    img = iV_LoadImageFile(fileName)
    If img = 0 Then
        ppData = Empty
        Exit Function
    End If
    ppData = img

    dataIMGLoad = True
End Function

Private Sub dataIMGRelease(pData As Variant)
    iV_FreeImageFile CLng(pData)
End Sub

'
' Release an Image
'
Private Sub dataImageRelease(pData As Variant)
    ' (delete psSprite: the last reference goes with the resource)
End Sub

' Load an audio file
Private Function dataAudioLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim track As Object
    If audio_Disabled() Then
        ppData = Empty
        ' No error occurred (sound is just disabled), so we return true
        dataAudioLoad = True
        Exit Function
    End If

    ' Load the track from a file
    Set track = sound_LoadTrackFromFile(fileName)
    If track Is Nothing Then
        ppData = Empty
    Else
        Set ppData = track
    End If

    dataAudioLoad = Not track Is Nothing
End Function

' Load an audio file
Private Function dataAudioCfgLoad(ByVal fileName As String, ppData As Variant) As Boolean
    Dim strName As String, ini As WzConfig
    strName = fileName
    If Not (InStr(1, strName, ".json", vbBinaryCompare) > 0) Then
        ASSERT_FAILURE "strName.find("".json"") != std::string::npos", "src/data.cpp:604", "dataAudioCfgLoad", "Audio effect file must be JSON format!"
        Exit Function
    End If
    If audio_Disabled() Then
        dataAudioCfgLoad = True
        Exit Function
    End If

    Set ini = WzConfig_new(fileName, WzConfig_ReadOnlyAndRequired)
    If Not loadAudioEffectFileData(ini) Then
        Exit Function
    End If

    ppData = 1&
    dataAudioCfgLoad = True
End Function

' Load a string resource file
Private Function dataStrResLoad(ByVal fileName As String, ppData As Variant) As Boolean
    ' recreate the string resource if it was freed by a WRF release
    If psStringRes Is Nothing Then
        If Not stringsInitialise() Then
            Exit Function
        End If
    End If

    If Not strresLoad(psStringRes, fileName) Then
        Exit Function
    End If

    Set ppData = psStringRes
    dataStrResLoad = True
End Function

Private Sub dataStrResRelease(pData As Variant)
    If Not psStringRes Is Nothing Then
        strresDestroy psStringRes
        Set psStringRes = Nothing
    End If
End Sub

Private Function jsLoad(ByVal fileName As String, ppData As Variant) As Boolean
    debug_ 652, LOG_WZ, "jsLoad", "jsload: %s", fileName
    ppData = Empty
    jsLoad = loadGlobalScript(fileName)
End Function

' (VB) The framework's call of a RES_BUFFERLOAD
Public Function dataBufferLoad(ByVal fn As Long, pBuffer() As Byte, ByVal size As Long, ppData As Variant) As Boolean
    Select Case fn
    Case fn_bufferSMSGLoad: dataBufferLoad = bufferSMSGLoad(pBuffer, size, ppData)
    End Select
End Function

' (VB) The framework's call of a RES_FILELOAD
Public Function dataFileLoad(ByVal fn As Long, ByVal fileName As String, ppData As Variant) As Boolean
    Select Case fn
    Case fn_bufferSFEATLoad: dataFileLoad = bufferSFEATLoad(fileName, ppData)
    Case fn_bufferSTEMPLLoad: dataFileLoad = bufferSTEMPLLoad(fileName, ppData)
    Case fn_dataAudioLoad: dataFileLoad = dataAudioLoad(fileName, ppData)
    Case fn_bufferSWEAPONLoad: dataFileLoad = bufferSWEAPONLoad(fileName, ppData)
    Case fn_bufferSBPIMDLoad: dataFileLoad = bufferSBPIMDLoad(fileName, ppData)
    Case fn_bufferSBRAINLoad: dataFileLoad = bufferSBRAINLoad(fileName, ppData)
    Case fn_bufferSSENSORLoad: dataFileLoad = bufferSSENSORLoad(fileName, ppData)
    Case fn_bufferSECMLoad: dataFileLoad = bufferSECMLoad(fileName, ppData)
    Case fn_bufferSREPAIRLoad: dataFileLoad = bufferSREPAIRLoad(fileName, ppData)
    Case fn_bufferSCONSTRLoad: dataFileLoad = bufferSCONSTRLoad(fileName, ppData)
    Case fn_bufferSPROPLoad: dataFileLoad = bufferSPROPLoad(fileName, ppData)
    Case fn_bufferSPROPTYPESLoad: dataFileLoad = bufferSPROPTYPESLoad(fileName, ppData)
    Case fn_bufferSTERRTABLELoad: dataFileLoad = bufferSTERRTABLELoad(fileName, ppData)
    Case fn_bufferSBODYLoad: dataFileLoad = bufferSBODYLoad(fileName, ppData)
    Case fn_bufferSWEAPMODLoad: dataFileLoad = bufferSWEAPMODLoad(fileName, ppData)
    Case fn_bufferSPROPSNDLoad: dataFileLoad = bufferSPROPSNDLoad(fileName, ppData)
    Case fn_dataAudioCfgLoad: dataFileLoad = dataAudioCfgLoad(fileName, ppData)
    Case fn_dataImageLoad: dataFileLoad = dataImageLoad(fileName, ppData)
    Case fn_dataTERTILESLoad: dataFileLoad = dataTERTILESLoad(fileName, ppData)
    Case fn_dataIMGLoad: dataFileLoad = dataIMGLoad(fileName, ppData)
    Case fn_dataStrResLoad: dataFileLoad = dataStrResLoad(fileName, ppData)
    Case fn_dataResearchMsgLoad: dataFileLoad = dataResearchMsgLoad(fileName, ppData)
    Case fn_bufferSSTRMODLoad: dataFileLoad = bufferSSTRMODLoad(fileName, ppData)
    Case fn_jsLoad: dataFileLoad = jsLoad(fileName, ppData)
    Case fn_bufferSSTRUCTLoad: dataFileLoad = bufferSSTRUCTLoad(fileName, ppData)
    Case fn_bufferRESCHLoad: dataFileLoad = bufferRESCHLoad(fileName, ppData)
    Case fn_dataProximityMsgLoad: dataFileLoad = dataProximityMsgLoad(fileName, ppData)
    Case fn_dataFlicMsgLoad: dataFileLoad = dataFlicMsgLoad(fileName, ppData)
    End Select
End Function

' (VB) The framework's call of a RES_FREE
Public Sub dataRelease(ByVal fn As Long, pData As Variant)
    Select Case fn
    Case fn_dataSMSGRelease: dataSMSGRelease pData
    Case fn_dataSFEATRelease: dataSFEATRelease pData
    Case fn_dataSTEMPLRelease: dataSTEMPLRelease pData
    Case fn_sound_ReleaseTrack: sound_ReleaseTrack pData
    Case fn_dataReleaseStats: dataReleaseStats pData
    Case fn_dataImageRelease: dataImageRelease pData
    Case fn_dataIMGRelease: dataIMGRelease pData
    Case fn_dataStrResRelease: dataStrResRelease pData
    Case fn_dataSSTRUCTRelease: dataSSTRUCTRelease pData
    Case fn_dataRESCHRelease: dataRESCHRelease pData
    End Select
End Sub

' Pass all the data loading functions to the framework library
' (BufferResourceTypes and FileResourceTypes, in their order: {aType, load, release})
Public Function dataInitLoadFuncs() As Boolean
    ' init the data integrity hash;
    resetDataHash

    ' iterate through buffer load functions
    If Not resAddBufferLoad("SMSG", fn_bufferSMSGLoad, fn_dataSMSGRelease) Then GoTo fail
    If Not resAddBufferLoad("IMD", 0, 0) Then GoTo fail ' ignored

    ' iterate through file load functions
    If Not resAddFileLoad("SFEAT", fn_bufferSFEATLoad, fn_dataSFEATRelease) Then GoTo fail          'feature stats file
    If Not resAddFileLoad("STEMPL", fn_bufferSTEMPLLoad, fn_dataSTEMPLRelease) Then GoTo fail       'template and associated files
    If Not resAddFileLoad("WAV", fn_dataAudioLoad, fn_sound_ReleaseTrack) Then GoTo fail
    If Not resAddFileLoad("SWEAPON", fn_bufferSWEAPONLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SBPIMD", fn_bufferSBPIMDLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SBRAIN", fn_bufferSBRAINLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SSENSOR", fn_bufferSSENSORLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SECM", fn_bufferSECMLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SREPAIR", fn_bufferSREPAIRLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SCONSTR", fn_bufferSCONSTRLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SPROP", fn_bufferSPROPLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SPROPTYPES", fn_bufferSPROPTYPESLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("STERRTABLE", fn_bufferSTERRTABLELoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SBODY", fn_bufferSBODYLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SWEAPMOD", fn_bufferSWEAPMODLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("SPROPSND", fn_bufferSPROPSNDLoad, fn_dataReleaseStats) Then GoTo fail
    If Not resAddFileLoad("AUDIOCFG", fn_dataAudioCfgLoad, 0) Then GoTo fail
    If Not resAddFileLoad("IMGPAGE", fn_dataImageLoad, fn_dataImageRelease) Then GoTo fail
    If Not resAddFileLoad("TERTILES", fn_dataTERTILESLoad, 0) Then GoTo fail
    If Not resAddFileLoad("IMG", fn_dataIMGLoad, fn_dataIMGRelease) Then GoTo fail
    If Not resAddFileLoad("TEXPAGE", 0, 0) Then GoTo fail ' ignored
    If Not resAddFileLoad("TCMASK", 0, 0) Then GoTo fail ' ignored
    If Not resAddFileLoad("STR_RES", fn_dataStrResLoad, fn_dataStrResRelease) Then GoTo fail
    If Not resAddFileLoad("RESEARCHMSG", fn_dataResearchMsgLoad, fn_dataSMSGRelease) Then GoTo fail
    If Not resAddFileLoad("SSTRMOD", fn_bufferSSTRMODLoad, 0) Then GoTo fail
    If Not resAddFileLoad("JAVASCRIPT", fn_jsLoad, 0) Then GoTo fail
    If Not resAddFileLoad("SSTRUCT", fn_bufferSSTRUCTLoad, fn_dataSSTRUCTRelease) Then GoTo fail   'structure stats and associated files
    If Not resAddFileLoad("RESCH", fn_bufferRESCHLoad, fn_dataRESCHRelease) Then GoTo fail         'research stats files
    If Not resAddFileLoad("PROX", fn_dataProximityMsgLoad, fn_dataSMSGRelease) Then GoTo fail
    If Not resAddFileLoad("FLIC", fn_dataFlicMsgLoad, fn_dataSMSGRelease) Then GoTo fail

    dataInitLoadFuncs = True
    Exit Function
fail:
    dataInitLoadFuncs = False ' error whilst adding a load function
End Function
