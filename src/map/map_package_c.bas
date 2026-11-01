Attribute VB_Name = "map_package_c"
' lib/wzmaplib/src/map_package.cpp: the helpers, the level details (loadLevelDetails: a classic .addon.lev /
' .xplayers.lev file, or the new level.json), the .gam files (the old binary <map>.gam, the new gam.json) and
' MapPackage::loadPackage (the MapPackage class is MapPackage.cls). optional<LevelDetails> / optional<GamInfo>
' are the object or Nothing. The JSON's try / catch (std::exception): nlohmann_json's errors (Err.Source
' "nlohmann::json") and e.what() their text.
' Missing (no caller in the game): exportLevelDetails (_LEV, _JSON), the .gam writers, getCurrentWzMapLibGeneratorName.
Option Explicit

Private Const MaxBinaryLevFileSize As Long = 2 * 1024& * 1024&
Private Const OLD_MAX_LEVEL_SIZE As Long = 20

' MARK: - Various helper functions

Public Function gamFilePathFromMapFolderPath(ByVal mapFolderPath As String, Optional ByVal pathSeparator As String = "/") As String
    Dim gamFilePath As String
    ' Remove any trailing slash(es) and append .gam on the end of the mapFolderPath
    gamFilePath = mapFolderPath
    Do While Len(gamFilePath) > 0 And strEndsWith(mapFolderPath, pathSeparator)
        gamFilePath = Left$(gamFilePath, Len(gamFilePath) - 1)
    Loop
    If Len(gamFilePath) > 0 Then
        gamFilePath = gamFilePath & ".gam"
    End If
    gamFilePathFromMapFolderPath = gamFilePath
End Function

' "<name>-T<digit>" -> "<name>" (isdigit: the C locale's 0-9), in place
Public Sub trimTechLevelFromMapName(mapName As String)
    Dim length As Long, c As Long
    length = Len(mapName)
    If length > 2 Then
        c = AscW(Mid$(mapName, length, 1))
        If Mid$(mapName, length - 2, 1) = "-" And Mid$(mapName, length - 1, 1) = "T" And c >= 48 And c <= 57 Then
            mapName = Left$(mapName, length - 3)
        End If
    End If
End Sub

' std::isspace in the C locale
Private Function isSpaceChar(ByVal ch As String) As Boolean
    Select Case AscW(ch)
    Case 9, 10, 11, 12, 13, 32: isSpaceChar = True
    End Select
End Function

Private Sub trim_whitespace(str_ As String)
    Dim a As Long, b As Long
    a = 1
    Do While a <= Len(str_)
        If Not isSpaceChar(Mid$(str_, a, 1)) Then Exit Do
        a = a + 1
    Loop
    b = Len(str_)
    Do While b >= a
        If Not isSpaceChar(Mid$(str_, b, 1)) Then Exit Do
        b = b - 1
    Loop
    str_ = Mid$(str_, a, b - a + 1)
End Sub

' optional<MapType>: True with the type
Public Function mapTypeFromString(ByVal typeStr As String, ByVal pCustomLogger As LoggingProtocol, mapType_ As MapType) As Boolean
    If typeStr = "skirmish" Then
        mapType_ = MapType_SKIRMISH
        mapTypeFromString = True
    Else
        printLog_ pCustomLogger, 116, LogLevel_Error, "mapTypeFromString", "Unknown ""type"" value: %s", typeStr
        mapTypeFromString = False
    End If
End Function

Public Function MapType_to_string(ByVal mapType_ As MapType) As String
    Select Case mapType_
        Case MapType_SKIRMISH: MapType_to_string = "skirmish"
        Case MapType_CAMPAIGN: MapType_to_string = "campaign"
        Case MapType_SAVEGAME: MapType_to_string = "savegame"
        Case Else: MapType_to_string = ""      ' silence warning
    End Select
End Function

' optional<MAP_TILESET>: True with the tileset
Public Function levelTilesetFromString(ByVal datasetStr As String, ByVal pCustomLogger As LoggingProtocol, tileset As MAP_TILESET) As Boolean
    If datasetStr = "arizona" Then
        tileset = MAP_TILESET_ARIZONA
        levelTilesetFromString = True
    ElseIf datasetStr = "urban" Then
        tileset = MAP_TILESET_URBAN
        levelTilesetFromString = True
    ElseIf datasetStr = "rockies" Then
        tileset = MAP_TILESET_ROCKIES
        levelTilesetFromString = True
    Else
        printLog_ pCustomLogger, 151, LogLevel_Error, "levelTilesetFromString", "Unknown ""tileset"" value: %s", datasetStr
        levelTilesetFromString = False
    End If
End Function

Public Function MAP_TILESET_to_string(ByVal mapTileset As MAP_TILESET) As String
    Select Case mapTileset
        Case MAP_TILESET_ARIZONA: MAP_TILESET_to_string = "arizona"
        Case MAP_TILESET_URBAN: MAP_TILESET_to_string = "urban"
        Case MAP_TILESET_ROCKIES: MAP_TILESET_to_string = "rockies"
        Case Else: MAP_TILESET_to_string = ""  ' silence warning
    End Select
End Function

Public Function internal_baseDir(ByVal fullPathToFile As String, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol) As String
    Dim pathSeparator As String, lastSlash As Long, k As Long
    pathSeparator = mapIO.pathSeparator()
    If StrPtr(pathSeparator) = 0 Then
        printLog_ pCustomLogger, 175, LogLevel_Error, "internal_baseDir", "IOProvider::pathSeparator returned null???"
        internal_baseDir = ""
        Exit Function
    End If

    ' Find the position of the last slash in the full path (find_last_of: any of the separator's characters)
    lastSlash = 0
    For k = 1 To Len(pathSeparator)
        If InStrRev(fullPathToFile, Mid$(pathSeparator, k, 1)) > lastSlash Then lastSlash = InStrRev(fullPathToFile, Mid$(pathSeparator, k, 1))
    Next
    If lastSlash = 0 Then
        ' Did not find a path separator
        printLog_ pCustomLogger, 184, LogLevel_Info_Verbose, "internal_baseDir", "Did not find parent path in ""%s"" (path separator: ""%s"")", fullPathToFile, pathSeparator
        ' Default to an empty path
        internal_baseDir = ""
        Exit Function
    End If

    ' Trim off the last path component
    internal_baseDir = Left$(fullPathToFile, lastSlash - 1)
End Function

' !strcasecmp(file + (len - Len(suffix)), suffix): an ASCII case-insensitive suffix
Private Function endsWithNoCase(ByVal s As String, ByVal suffix As String) As Boolean
    If Len(s) > Len(suffix) Then endsWithNoCase = (StrComp(strAsciiToLower(Right$(s, Len(suffix))), strAsciiToLower(suffix), vbBinaryCompare) = 0)
End Function

' Load levels details from either:
' - a classic <map>.addon.lev / <map>.xplayers.lev file
' - the new (in-folder) level.json file
' NOTE: This function does *not* handle all of the special properties for campaign level types that are contained in .lev files (currently) - such as gamedesc.lev - it is intended for individual .lev files used by skirmish / multiplay maps
Public Function loadLevelDetails(ByVal levelFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As LevelDetails
    If strEndsWith(levelFile, ".json") Then
        Set loadLevelDetails = loadLevelDetails_JSON(levelFile, mapIO, pCustomLogger)
        Exit Function
    ' Must do this strcasecmp comparison to match old behavior
    ElseIf endsWithNoCase(levelFile, ".addon.lev") Or endsWithNoCase(levelFile, ".xplayers.lev") Then
        Set loadLevelDetails = loadLevelDetails_LEV(levelFile, mapIO, pCustomLogger)
        Exit Function
    End If
    printLog_ pCustomLogger, 209, LogLevel_Error, "loadLevelDetails", "loadLevelDetails failure - unsupported file: %s", levelFile
    Set loadLevelDetails = Nothing
End Function

Private Function convertDateSlashesToHyphens(dateString As String) As Boolean
    ' YYYY/MM/DD (...) to YYYY-MM-DD (...)
    If Len(dateString) < 10 Then
        convertDateSlashesToHyphens = False
        Exit Function
    End If

    If Mid$(dateString, 5, 1) = "/" And Mid$(dateString, 8, 1) = "/" Then
        Mid$(dateString, 5, 1) = "-"
        Mid$(dateString, 8, 1) = "-"
        convertDateSlashesToHyphens = True
        Exit Function
    End If

    convertDateSlashesToHyphens = False
End Function

Private Function strAsciiToLower(ByVal str_ As String) As String
    Dim result As String, k As Long, c As Long
    result = str_
    For k = 1 To Len(result)
        c = AscW(Mid$(result, k, 1))
        If c >= 65 And c <= 90 Then Mid$(result, k, 1) = ChrW$(c + 32)
    Next
    strAsciiToLower = result
End Function

' str.rfind(prefix, 0) == 0
Private Function startsWith(ByVal str_ As String, ByVal prefix As String) As Boolean
    startsWith = (Left$(str_, Len(prefix)) = prefix)
End Function

' (VB) "CC BY-SA 3.0 + GPL v2+" ...: licenseToSPDX's mapping ("": none)
Private Function licenseToSPDX(ByVal licenseIdentifier As String) As String
    Select Case licenseIdentifier
        '   CC BY-SA 3.0 + GPL v2+
        Case "CC BY-SA 3.0 + GPL v2+": licenseToSPDX = "CC-BY-SA-3.0 OR GPL-2.0-or-later"
        '   CC BY 3.0 + GPL v2+
        Case "CC BY 3.0 + GPL v2+": licenseToSPDX = "CC-BY-3.0 OR GPL-2.0-or-later"
        '   GPL 2+
        Case "GPL 2+": licenseToSPDX = "GPL-2.0-or-later"
        '   CC0
        Case "CC0": licenseToSPDX = "CC0-1.0"
    End Select
End Function

' std::stoi(arg): the leading integer (whitespace, a sign, digits); False when there is none (std::invalid_argument)
' or it overflows an int (std::out_of_range)
Private Function stoi(ByVal arg As String, value As Long) As Boolean
    Dim i As Long, neg As Boolean, digits As String, v As Double
    i = 1
    Do While i <= Len(arg)
        If Not isSpaceChar(Mid$(arg, i, 1)) Then Exit Do
        i = i + 1
    Loop
    If Mid$(arg, i, 1) = "-" Or Mid$(arg, i, 1) = "+" Then
        neg = (Mid$(arg, i, 1) = "-")
        i = i + 1
    End If
    Do While i <= Len(arg) And Mid$(arg, i, 1) >= "0" And Mid$(arg, i, 1) <= "9"
        digits = digits & Mid$(arg, i, 1)
        i = i + 1
    Loop
    If Len(digits) = 0 Then Exit Function
    If Len(digits) > 12 Then Exit Function
    v = CDbl(digits)
    If neg Then v = -v
    If v < -2147483648# Or v > 2147483647# Then Exit Function
    value = CLng(v)
    stoi = True
End Function

Public Function loadLevelDetails_LEV(ByVal levelFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As LevelDetails
    Dim fileData() As Byte, fileSize As Long, loadFileResult As LoadFullFileResult
    Dim details() As LevelDetails, detailsCount As Long, detailsFatalError() As Boolean
    Dim detectedGenerator As Long, generatorVersion As String, hasGeneratorVersion As Boolean   ' KnownGenerators (0: none)
    Dim createdDate As String, uploadedDate As String, author As String, license As String
    Dim evKind() As Long, evA() As String, evB() As String, evCount As Long, e As Long
    Dim command As String, arg As String, playersValue As Long, typeValue As Long, tileset As MAP_TILESET
    Dim comment As String, licenseIdentifier As String, licenseMapping As String
    Dim firstValidLevelEntryIdx As Long, hasFirstValid As Boolean, firstMapFolderPath As String, idx As Long
    Const KnownGenerators_FlaME As Long = 1
    Const KnownGenerators_euphobos As Long = 2

    ' 1.) Load file
    loadFileResult = mapIO.loadFullFile(levelFile, fileData, fileSize, MaxBinaryLevFileSize, False)
    Select Case loadFileResult
        Case LoadFullFileResult_SUCCESS
        Case LoadFullFileResult_FAILURE_OPEN
            Set loadLevelDetails_LEV = Nothing
            Exit Function
        Case LoadFullFileResult_FAILURE_READ
            ' file exists but can't be read
            printLog_ pCustomLogger, 258, LogLevel_Error, "loadLevelDetails_LEV", "Failed to read file: %s", levelFile
            Set loadLevelDetails_LEV = Nothing
            Exit Function
        Case LoadFullFileResult_FAILURE_EXCEEDS_MAXFILESIZE
            printLog_ pCustomLogger, 261, LogLevel_Error, "loadLevelDetails_LEV", "File is too large: %s", levelFile
            Set loadLevelDetails_LEV = Nothing
            Exit Function
    End Select
    If fileSize = 0 Then
        printLog_ pCustomLogger, 266, LogLevel_Error, "loadLevelDetails_LEV", "Empty file: %s", levelFile
        Set loadLevelDetails_LEV = Nothing
        Exit Function
    End If

    If Not levParseBasic(fileData, fileSize, mapIO, evKind, evA, evB, evCount, pCustomLogger) Then
        printLog_ pCustomLogger, 508, LogLevel_Error, "loadLevelDetails_LEV", "Failed to parse .lev file: %s", levelFile
        Set loadLevelDetails_LEV = Nothing
        Exit Function
    End If

    For e = 0 To evCount - 1
        If evKind(e) = LEVEV_COMMAND Then
            ' handleCommand(command, arg)
            command = evA(e)
            arg = evB(e)
            If command = "level" Then
                ReDim Preserve details(0 To detailsCount)
                ReDim Preserve detailsFatalError(0 To detailsCount)
                Set details(detailsCount) = New LevelDetails
                detailsFatalError(detailsCount) = False
                details(detailsCount).name_ = arg
                details(detailsCount).createdDate = createdDate
                details(detailsCount).uploadedDate = uploadedDate
                details(detailsCount).author = author
                details(detailsCount).license = license
                details(detailsCount).generator = generatorVersion
                details(detailsCount).hasGenerator = hasGeneratorVersion
                detailsCount = detailsCount + 1
                GoTo nextEvent
            End If

            If detailsCount = 0 Then
                printLog_ pCustomLogger, 312, LogLevel_Warning, "loadLevelDetails_LEV", "LEV File Parse Error: `Syntax Error`: Received an unexpected token (prior to a `level` token): %s %s", command, arg
                GoTo nextEvent
            End If

            If command = "players" Then
                playersValue = 0
                If Not stoi(arg, playersValue) Then
                    printLog_ pCustomLogger, 323, LogLevel_Error, "loadLevelDetails_LEV", "LEV File Parse Error: ""players"" value (""%s"") must be a valid number; ignoring level[%zu](""%s"")", arg, detailsCount - 1, details(detailsCount - 1).name_
                    detailsFatalError(detailsCount - 1) = True
                    GoTo nextEvent
                End If
                If playersValue < 0 Or playersValue > 255 Then     ' static_cast<size_t>(playersValue) > UINT8_MAX
                    printLog_ pCustomLogger, 329, LogLevel_Error, "loadLevelDetails_LEV", "LEV File Parse Error: ""players"" value (""%s"") is unsupported; ignoring level[%zu](""%s"")", arg, detailsCount - 1, details(detailsCount - 1).name_
                    detailsFatalError(detailsCount - 1) = True
                    GoTo nextEvent
                End If
                details(detailsCount - 1).players = playersValue
            ElseIf command = "type" Then
                typeValue = 0
                If Not stoi(arg, typeValue) Then
                    printLog_ pCustomLogger, 342, LogLevel_Error, "loadLevelDetails_LEV", "LEV File Parse Error: ""type"" value (""%s"") must be a valid number; ignoring level[%zu](""%s"")", arg, detailsCount - 1, details(detailsCount - 1).name_
                    detailsFatalError(detailsCount - 1) = True
                    GoTo nextEvent
                End If
                ' Check for valid / supported type numbers from old .lev files
                If typeValue = SUPPORTED_LEVEL_TYPES_SKIRMISH Or typeValue = SUPPORTED_LEVEL_TYPES_MULTI_SKIRMISH2 _
                        Or typeValue = SUPPORTED_LEVEL_TYPES_MULTI_SKIRMISH3 Or typeValue = SUPPORTED_LEVEL_TYPES_MULTI_SKIRMISH4 Then
                    details(detailsCount - 1).type_ = MapType_SKIRMISH
                Else
                    printLog_ pCustomLogger, 356, LogLevel_Error, "loadLevelDetails_LEV", "LEV File Parse Error: ""type"" value (""%s"") is unknown / unsupported; ignoring level[%zu](""%s"")", arg, detailsCount - 1, details(detailsCount - 1).name_
                    detailsFatalError(detailsCount - 1) = True
                    GoTo nextEvent
                End If
            ElseIf command = "dataset" Then
                If Not convertLevMapTilesetType(arg, tileset) Then
                    printLog_ pCustomLogger, 366, LogLevel_Error, "loadLevelDetails_LEV", "LEV File Parse Error: ""dataset"" value (""%s"") is unknown / unsupported; ignoring level[%zu](""%s"")", arg, detailsCount - 1, details(detailsCount - 1).name_
                    detailsFatalError(detailsCount - 1) = True
                    GoTo nextEvent
                End If
                details(detailsCount - 1).tileset = tileset
            ElseIf command = "game" Then
                ' Remove the ".gam" from the end of the filename
                details(detailsCount - 1).mapFolderPath = arg
                If strEndsWith(arg, ".gam") Then
                    details(detailsCount - 1).mapFolderPath = Left$(details(detailsCount - 1).mapFolderPath, Len(details(detailsCount - 1).mapFolderPath) - 4)
                Else
                    printLog_ pCustomLogger, 382, LogLevel_Warning, "loadLevelDetails_LEV", "LEV File Parse Warning: Unexpected: ""game"" command argument (""%s"") does not end with .gam", arg
                End If
            ElseIf command = "data" Then
                ' explicitly ignore `data` lines
                GoTo nextEvent
            Else
                printLog_ pCustomLogger, 392, LogLevel_Warning, "loadLevelDetails_LEV", "LEV File Parse Error: `Syntax Error`: Unknown / unsupported token: %s", command
            End If
        Else
            ' handleCommentLine(commentLine)
            ' Trim all whitespace
            comment = evA(e)
            trim_whitespace comment

            ' Detect generator comment
            If startsWith(strAsciiToLower(comment), "made with flame") Then
                If detectedGenerator = 0 Then
                    detectedGenerator = KnownGenerators_FlaME
                    generatorVersion = Mid$(comment, 11)   ' after "Made with "
                    hasGeneratorVersion = True
                Else
                    ' multiple generator comments??
                    printLog_ pCustomLogger, 413, LogLevel_Warning, "loadLevelDetails_LEV", "LEV File Parse Warning: Additional comment detected that could signify a generator (already found one) - ignoring: ""%s""", comment
                End If
                GoTo nextEvent
            ElseIf startsWith(comment, "Generated by map database at ") Then
                If detectedGenerator = 0 Then
                    detectedGenerator = KnownGenerators_euphobos
                Else
                    ' multiple generator comments??
                    printLog_ pCustomLogger, 426, LogLevel_Warning, "loadLevelDetails_LEV", "LEV File Parse Warning: Additional comment detected that could signify a generator (already found one) - ignoring: ""%s""", comment
                End If
                GoTo nextEvent
            End If

            ' Handle extracting additional info from comments
            If detectedGenerator = KnownGenerators_FlaME Then
                ' Comment fields of interest are:
                ' Date: YYYY/MM/DD HH:MM:SS
                ' Author: AUTHORNAME
                ' License: (see known options in licenseToSPDX above)
                If startsWith(comment, "Date: ") Then
                    If Len(createdDate) > 0 Then
                        printLog_ pCustomLogger, 442, LogLevel_Error, "loadLevelDetails_LEV", "Duplicate `Date` comment in .lev file: ""%s""", comment
                        GoTo nextEvent
                    End If
                    createdDate = Mid$(comment, 7)
                    trim_whitespace createdDate

                    ' Attempt to convert from YYYY/MM/DD HH:MM:SS to YYYY-MM-DD HH:MM:SS format (where HH:MM:SS is optional)
                    convertDateSlashesToHyphens createdDate
                    GoTo nextEvent
                End If
                ' additional fields handled below
            ElseIf detectedGenerator = KnownGenerators_euphobos Then
                ' Comment fields of interest are:
                ' Upload date: YYYY-MM-DD HH:MM:SS +0000
                ' Author: AUTHORNAME
                ' License: (see known options in licenseToSPDX above)

                If startsWith(comment, "Upload date: ") Then
                    If Len(uploadedDate) > 0 Then
                        printLog_ pCustomLogger, 465, LogLevel_Error, "loadLevelDetails_LEV", "Duplicate `Upload date` comment in .lev file: ""%s""", comment
                        GoTo nextEvent
                    End If
                    uploadedDate = Mid$(comment, 14)
                    trim_whitespace uploadedDate
                    GoTo nextEvent
                End If
                ' additional fields handled below
            End If

            If startsWith(comment, "Author: ") Then
                If Len(author) > 0 Then
                    printLog_ pCustomLogger, 479, LogLevel_Error, "loadLevelDetails_LEV", "Duplicate `Author` comment in .lev file: ""%s""", comment
                    GoTo nextEvent
                End If
                author = Mid$(comment, 9)
                trim_whitespace author
            ElseIf startsWith(comment, "License: ") Then
                If Len(license) > 0 Then
                    printLog_ pCustomLogger, 489, LogLevel_Error, "loadLevelDetails_LEV", "Duplicate `License` comment in .lev file: ""%s""", comment
                    GoTo nextEvent
                End If
                licenseIdentifier = Mid$(comment, 10)
                trim_whitespace licenseIdentifier
                licenseMapping = licenseToSPDX(licenseIdentifier)
                If Len(licenseMapping) > 0 Then
                    license = licenseMapping
                Else
                    ' (4.7.0 prints license, still empty here, not licenseIdentifier)
                    printLog_ pCustomLogger, 501, LogLevel_Error, "loadLevelDetails_LEV", "Unknown / unsupported license in .lev file: ""%s""", license
                End If
            End If
        End If
nextEvent:
    Next

    If detailsCount = 0 Then
        Set loadLevelDetails_LEV = Nothing
    ElseIf detailsCount = 1 Then
        ' Remove any tech level suffix from name of first entry
        Dim nm As String
        nm = details(0).name_
        trimTechLevelFromMapName nm
        details(0).name_ = nm

        Set loadLevelDetails_LEV = details(0)
    Else
        ' more than one level entry in the file

        ' find first level entry that is valid
        For idx = 0 To detailsCount - 1
            If Not detailsFatalError(idx) Then
                firstValidLevelEntryIdx = idx
                hasFirstValid = True
                Exit For
            End If
        Next
        If Not hasFirstValid Then
            ' no valid level entries - all had fatal errors associated with them!
            printLog_ pCustomLogger, 540, LogLevel_Error, "loadLevelDetails_LEV", "No valid / supported levels found in .lev file"
            Set loadLevelDetails_LEV = Nothing
            Exit Function
        End If

        ' check that they all point to the same map (we don't support multi-map packs here, which have long been deprecated)
        firstMapFolderPath = details(firstValidLevelEntryIdx).mapFolderPath
        For idx = 0 To detailsCount - 1
            If detailsFatalError(idx) Then
                GoTo continue_ ' skip problematic levels
            End If
            If details(idx).mapFolderPath <> firstMapFolderPath Then
                ' Found map pack - unsupported!
                printLog_ pCustomLogger, 556, LogLevel_Error, "loadLevelDetails_LEV", "Found map pack (unsupported!) - level map path (%s) does not match first level map path: (%s)", details(idx).mapFolderPath, firstMapFolderPath
                Set loadLevelDetails_LEV = Nothing
                Exit Function
            End If
continue_:
        Next

        ' otherwise, everything was pointing to the same map, so these were probably the old format that listed all the tech levels separately
        ' (which is no longer needed)

        ' Remove any tech level suffix from name of first entry
        nm = details(firstValidLevelEntryIdx).name_
        trimTechLevelFromMapName nm
        details(firstValidLevelEntryIdx).name_ = nm

        Set loadLevelDetails_LEV = details(firstValidLevelEntryIdx)
    End If
End Function

' output_priorGenerator: the "prior-generator" value when the caller asks (hasPriorGenerator) and there is one
Public Function loadLevelDetails_JSON(ByVal levelJsonFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing, _
        Optional output_priorGenerator As String = vbNullString, Optional ByVal wantPriorGenerator As Boolean = False) As LevelDetails
    Dim mRoot As nlohmann_json, details As LevelDetails, it As nlohmann_json, nameIt As nlohmann_json, i As Long
    Dim mt As MapType, playersNum As Long, ts As MAP_TILESET, otherAuthor As nlohmann_json, what As String

    Set mRoot = loadJsonObjectFromFile(levelJsonFile, mapIO, pCustomLogger)
    If mRoot Is Nothing Then
        ' Failed to load JSON - rely on loadJsonFromFile to handle output of errors
        Set loadLevelDetails_JSON = Nothing
        Exit Function
    End If

    printLog_ pCustomLogger, 580, LogLevel_Info, "loadLevelDetails_JSON", "Loading: %s", levelJsonFile

    Set details = New LevelDetails

    On Error GoTo catch_
    ' Get required properties
    details.name_ = mRoot.at("name").get_string()
    If Not mapTypeFromString(mRoot.at("type").get_string(), pCustomLogger, mt) Then
        ' levelTypeFromString already handled logging an error
        Set loadLevelDetails_JSON = Nothing
        Exit Function
    End If
    details.type_ = mt
    playersNum = mRoot.at("players").get_uint()
    If playersNum < 0 Or playersNum > 255 Then          ' (unsigned) > UINT8_MAX
        printLog_ pCustomLogger, 620, LogLevel_Error, "loadLevelDetails_JSON", "Invalid ""players"" value: %u", playersNum
        Set loadLevelDetails_JSON = Nothing
        Exit Function
    End If
    details.players = playersNum
    If Not levelTilesetFromString(mRoot.at("tileset").get_string(), pCustomLogger, ts) Then
        ' levelTypeFromString already handled logging an error
        Set loadLevelDetails_JSON = Nothing
        Exit Function
    End If
    details.tileset = ts

    ' mapFolderPath = the path to this levelJsonFile with the last path component (the json filename) removed
    details.mapFolderPath = internal_baseDir(levelJsonFile, mapIO, pCustomLogger)
    On Error GoTo 0

    ' Get optional info
    Set it = mRoot.find("author")
    If Not it Is Nothing Then
        Set nameIt = it.find("name")
        If Not nameIt Is Nothing Then
            If nameIt.is_string() Then
                details.author = nameIt.get_string()
            End If
        End If
    End If
    Set it = mRoot.find("created-date")
    If Not it Is Nothing Then
        If it.is_string() Then
            details.createdDate = it.get_string()
        End If
    End If
    Set it = mRoot.find("spdx-license-id")
    If Not it Is Nothing Then
        If it.is_string() Then
            details.license = it.get_string()
        End If
    End If
    Set it = mRoot.find("generator")
    If Not it Is Nothing Then
        If it.is_string() Then
            details.generator = it.get_string()
            details.hasGenerator = True
        End If
    End If
    Set it = mRoot.find("additional-authors")
    If Not it Is Nothing Then
        If it.is_array() Then
            For i = 0 To it.size() - 1
                Set otherAuthor = it.at_idx(i)
                If Not otherAuthor.is_object() Then GoTo continue_
                Set nameIt = otherAuthor.find("name")
                If Not nameIt Is Nothing Then
                    If nameIt.is_string() Then
                        details.additionalAuthors_push_back nameIt.get_string()
                    End If
                End If
continue_:
            Next
        Else
            printLog_ pCustomLogger, 697, LogLevel_Warning, "loadLevelDetails_JSON", "JSON document has additional-authors key that is not an array: %s", levelJsonFile
        End If
    End If
    If wantPriorGenerator Then
        Set it = mRoot.find("prior-generator")
        If Not it Is Nothing Then
            If it.is_string() Then
                output_priorGenerator = it.get_string()
            End If
        End If
    End If

    Set loadLevelDetails_JSON = details
    Exit Function

catch_:
    If Err.Source <> "nlohmann::json" Then Err.Raise Err.Number, Err.Source, Err.Description
    what = Err.Description
    On Error GoTo 0
    printLog_ pCustomLogger, 636, LogLevel_Error, "loadLevelDetails_JSON", "JSON document from %s is invalid: %s", levelJsonFile, what
    Set loadLevelDetails_JSON = Nothing
End Function

' SDL_Swap32
Private Function Swap32(ByVal v As Long) As Long
    Dim u As Double, b0 As Long, b1 As Long, b2 As Long, b3 As Long, r As Double
    u = v
    If u < 0 Then u = u + 4294967296#
    b0 = u - Int(u / 256#) * 256#
    u = Int(u / 256#)
    b1 = u - Int(u / 256#) * 256#
    u = Int(u / 256#)
    b2 = u - Int(u / 256#) * 256#
    u = Int(u / 256#)
    b3 = u
    r = ((b0 * 256# + b1) * 256# + b2) * 256# + b3
    If r >= 2147483648# Then r = r - 4294967296#
    Swap32 = CLng(r)
End Function

Public Function loadGamFile_OldBinary(ByVal filePath As String, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol) As GamInfo
    Dim readStream As BinaryIOStream, aFileType() As Byte, version As Long, gamInfo_ As GamInfo, n As Long
    Dim v As Long, levelName() As Byte, k As Long
    Set readStream = mapIO.openBinaryStream(filePath, OpenMode_READ)
    If readStream Is Nothing Then
        ' Failed to open binary stream
        Set loadGamFile_OldBinary = Nothing
        Exit Function
    End If

    ' get the file header
    ReDim aFileType(0 To 3)
    version = 0

    If Not readStream.readBytes(aFileType, 0, 4, n) Then n = -1
    If n <> 4 Then
        Set loadGamFile_OldBinary = Nothing
        Exit Function
    End If

    If aFileType(0) <> 103 Or aFileType(1) <> 97 Or aFileType(2) <> 109 Or aFileType(3) <> 101 Then   ' "game"
        ' Invalid header
        Set loadGamFile_OldBinary = Nothing
        Exit Function
    End If

    ' All binary save game file versions below version 35 (i.e. _not_ version 35 itself)
    ' have their version numbers stored as little endian. Versions from 35 and
    ' onward use big-endian. This basically means that, because of endian
    ' swapping, numbers from 35 and onward will be ridiculously high if a
    ' little-endian byte-order is assumed.

    ' So first attempt reading it as little endian
    If Not BinaryIOStream_readULE32(readStream, version) Then
        Set loadGamFile_OldBinary = Nothing
        Exit Function
    End If

    If version > 34 Or version < 0 Then                 ' (uint32_t) version > 34
        ' Apparently we get a larger number than expected if using little-endian.
        ' So assume we have a version of 35 and onward

        ' Reverse the little-endian decoding
        version = Swap32(version)
    End If

    Set gamInfo_ = New GamInfo

    ' All binary gam files from version 34 or before are little endian.
    ' All from version 35, are big-endian.

    ReDim levelName(0 To OLD_MAX_LEVEL_SIZE - 1)
    If Not BinaryIOStream_readULE32(readStream, v) Then GoTo failed_
    gamInfo_.gameTime = v
    If Not BinaryIOStream_readULE32(readStream, v) Then GoTo failed_
    gamInfo_.GameType = v
    If Not BinaryIOStream_readSLE32(readStream, v) Then GoTo failed_
    gamInfo_.ScrollMinX = v
    If Not BinaryIOStream_readSLE32(readStream, v) Then GoTo failed_
    gamInfo_.ScrollMinY = v
    If Not BinaryIOStream_readULE32(readStream, v) Then GoTo failed_
    gamInfo_.ScrollMaxX = v
    If Not BinaryIOStream_readULE32(readStream, v) Then GoTo failed_
    gamInfo_.ScrollMaxY = v
    If Not readStream.readBytes(levelName, 0, OLD_MAX_LEVEL_SIZE, n) Then GoTo failed_
    If n <> OLD_MAX_LEVEL_SIZE Then GoTo failed_
    k = 0
    Do While k < OLD_MAX_LEVEL_SIZE
        If levelName(k) = 0 Then Exit Do
        k = k + 1
    Loop
    gamInfo_.levelName = utf8_decode(levelName, 0, k)

    If version > 34 Or version < 0 Then
        ' Reverse the little-endian decoding
        gamInfo_.gameTime = Swap32(gamInfo_.gameTime)
        gamInfo_.GameType = Swap32(gamInfo_.GameType)
        gamInfo_.ScrollMinX = Swap32(gamInfo_.ScrollMinX)
        gamInfo_.ScrollMinY = Swap32(gamInfo_.ScrollMinY)
        gamInfo_.ScrollMaxX = Swap32(gamInfo_.ScrollMaxX)
        gamInfo_.ScrollMaxY = Swap32(gamInfo_.ScrollMaxY)
    End If

    Set loadGamFile_OldBinary = gamInfo_
    Exit Function

failed_:
    ' Failed to read basic data
    Set loadGamFile_OldBinary = Nothing
End Function

Public Function loadGamFile_JSON(ByVal filePath As String, ByVal mapIO As IOProvider, ByVal pCustomLogger As LoggingProtocol) As GamInfo
    Dim mRoot As nlohmann_json, gamInfo_ As GamInfo, version As Long, levelNameStr As String, b() As Byte, n As Long, what As String
    Set mRoot = loadJsonObjectFromFile(filePath, mapIO, pCustomLogger)
    If mRoot Is Nothing Then
        ' Failed to load JSON - rely on loadJsonFromFile to handle output of errors
        Set loadGamFile_JSON = Nothing
        Exit Function
    End If

    printLog_ pCustomLogger, 966, LogLevel_Info, "loadGamFile_JSON", "Loading: %s", filePath

    Set gamInfo_ = New GamInfo
    version = 0

    On Error GoTo catch_
    ' Get required properties
    version = mRoot.at("version").get_uint()
    If version >= 0 And version < 7 Then                ' (uint32_t) version < 7
        printLog_ pCustomLogger, 977, LogLevel_Warning, "loadGamFile_JSON", "Unexpected gam.json version (%u) in: %s", version, filePath
    End If
    gamInfo_.gameTime = mRoot.at("gameTime").get_uint()
    gamInfo_.GameType = mRoot.at("GameType").get_uint()
    gamInfo_.ScrollMinX = mRoot.at("ScrollMinX").get_int()
    gamInfo_.ScrollMinY = mRoot.at("ScrollMinY").get_int()
    gamInfo_.ScrollMaxX = mRoot.at("ScrollMaxX").get_uint()
    gamInfo_.ScrollMaxY = mRoot.at("ScrollMaxY").get_uint()
    levelNameStr = mRoot.at("levelName").get_string()
    On Error GoTo 0
    ' strncpy(gamInfo.levelName, levelNameStr.c_str(), OLD_MAX_LEVEL_SIZE - 1): its first 19 bytes (to the first NUL)
    n = utf8_encode(levelNameStr, b)
    If n > OLD_MAX_LEVEL_SIZE - 1 Then n = OLD_MAX_LEVEL_SIZE - 1
    gamInfo_.levelName = utf8_decode(b, 0, n)

    Set loadGamFile_JSON = gamInfo_
    Exit Function

catch_:
    If Err.Source <> "nlohmann::json" Then Err.Raise Err.Number, Err.Source, Err.Description
    what = Err.Description
    On Error GoTo 0
    printLog_ pCustomLogger, 990, LogLevel_Error, "loadGamFile_JSON", "JSON document from %s is invalid: %s", filePath, what
    Set loadGamFile_JSON = Nothing
End Function

' Load map GAM details from either:
' - a classic <map>.gam file
' - the new (in-folder) gam.json file
' NOTE: This function only handles the basic gam data used by maps
Public Function loadGamFile(ByVal gamFile As String, ByVal mapIO As IOProvider, Optional ByVal pCustomLogger As LoggingProtocol = Nothing) As GamInfo
    If strEndsWith(gamFile, "gam.json") Then
        Set loadGamFile = loadGamFile_JSON(gamFile, mapIO, pCustomLogger)
        Exit Function
    ElseIf strEndsWith(gamFile, ".gam") Then
        Set loadGamFile = loadGamFile_OldBinary(gamFile, mapIO, pCustomLogger)
        Exit Function
    End If
    printLog_ pCustomLogger, 1011, LogLevel_Error, "loadGamFile", "loadGamInfo failure - unsupported file: %s", gamFile
    Set loadGamFile = Nothing
End Function

' Load a map package from a specified path, optionally supplying:
' - a logger
' - a WzMap::IOProvider
'
' The default StdIOProvider will assume pathToMapPackage is a path to an extracted map package (i.e. standard filesystem I/O)
'
' To load from an archive (.zip/.wz), create a custom implementation of WzMap::IOProvider that supports compressed archive files that you initialize with the path to the zip. (The game's: WzMapZipIO, the exe's.)
' (pMapIO Nothing: the default StdIOProvider)
Public Function MapPackage_loadPackage(ByVal pathToMapPackage As String, Optional ByVal logger As LoggingProtocol = Nothing, Optional ByVal pMapIO As IOProvider = Nothing) As MapPackage
    Dim pCustomLogger As LoggingProtocol, mapIO As IOProvider, mapFolderNames() As String, mapFolderCount As Long
    Dim pathSeparator As String, baseMapsPath As String, enumSuccess As Boolean, names() As String, n As Long, i As Long
    Dim loadedLevelDetails As LevelDetails, result As MapPackage, rootLevFiles() As String, rootLevCount As Long
    Dim levFilePath As String, idx As Long, originalGenerator As String, files() As String, nf As Long, k As Long

    Set pCustomLogger = logger
    If pMapIO Is Nothing Then Set pMapIO = New StdIOProvider     ' = std::make_shared<StdIOProvider>()

    Set mapIO = pMapIO

    ' 1.) Look for a map folder
    ' i.e. multiplay/maps/<map name>
    pathSeparator = mapIO.pathSeparator()
    baseMapsPath = IOProvider_pathJoin(mapIO, pathToMapPackage, IOProvider_pathJoin(mapIO, "multiplay", "maps"))
    enumSuccess = mapIO.enumerateFolders(baseMapsPath, names, n)
    For i = 0 To n - 1
        If Len(names(i)) = 0 Then GoTo continue1_
        ReDim Preserve mapFolderNames(0 To mapFolderCount)
        mapFolderNames(mapFolderCount) = baseMapsPath & pathSeparator & names(i)
        mapFolderCount = mapFolderCount + 1
continue1_:
    Next

    If enumSuccess And mapFolderCount > 0 Then
        If mapFolderCount > 1 Then
            ' multi-map packs are not supported (deprecated for a while)
            printLog_ pCustomLogger, 1161, LogLevel_Error, "loadPackage", "Detected multiple (%zu) map folders - multi-map packs are not supported", mapFolderCount
            Set MapPackage_loadPackage = Nothing
            Exit Function
        End If

        ' check the map folder for a level.json file
        Set loadedLevelDetails = loadLevelDetails_JSON(IOProvider_pathJoin(mapIO, mapFolderNames(0), "level.json"), mapIO, pCustomLogger)
        If Not loadedLevelDetails Is Nothing Then
            ' Successfully found a level.json file inside the enumerated map folder
            Set result = New MapPackage
            result.initLoaded loadedLevelDetails, LevelFormat_JSON, pathToMapPackage, pMapIO, logger, False
            Set MapPackage_loadPackage = result
            Exit Function
        End If
    End If

    ' 2.) Look for a level.json file in the root (i.e. a "flattened" plain map package)
    Set loadedLevelDetails = loadLevelDetails_JSON(IOProvider_pathJoin(mapIO, pathToMapPackage, "level.json"), mapIO, pCustomLogger)
    If Not loadedLevelDetails Is Nothing Then
        ' Successfully found a level.json file inside the root folder
        loadedLevelDetails.mapFolderPath = ""
        Set result = New MapPackage
        result.initLoaded loadedLevelDetails, LevelFormat_JSON, pathToMapPackage, pMapIO, logger, True
        Set MapPackage_loadPackage = result
        Exit Function
    End If

    ' 3.) If no map folder with a valid level.json was found, nor a level.json in the root, look for a .lev file in the root.
    enumSuccess = mapIO.enumerateFiles(pathToMapPackage, names, n)
    For i = 0 To n - 1
        If Len(names(i)) = 0 Then GoTo continue2_
        If endsWithNoCase(names(i), ".addon.lev") Or endsWithNoCase(names(i), ".xplayers.lev") Then
            ReDim Preserve rootLevFiles(0 To rootLevCount)
            rootLevFiles(rootLevCount) = names(i)
            rootLevCount = rootLevCount + 1
        End If
continue2_:
    Next

    If Not enumSuccess Then
        printLog_ pCustomLogger, 1212, LogLevel_Error, "loadPackage", "Failed to enumerate files"
        Set MapPackage_loadPackage = Nothing
        Exit Function
    End If

    If rootLevCount = 0 Then
        printLog_ pCustomLogger, 1218, LogLevel_Error, "loadPackage", "Unable to detect any in-folder map level.json files (new format), or root .lev files (older format)."
        Set MapPackage_loadPackage = Nothing
        Exit Function
    End If

    levFilePath = pathToMapPackage
    If Len(levFilePath) > 0 And Not strEndsWith(levFilePath, pathSeparator) Then
        levFilePath = levFilePath & pathSeparator
    End If
    For idx = 0 To rootLevCount - 1
        originalGenerator = ""
        Set loadedLevelDetails = loadLevelDetails_LEV(levFilePath & rootLevFiles(idx), mapIO, pCustomLogger)
        If Not loadedLevelDetails Is Nothing Then
            enumSuccess = mapIO.enumerateFiles(loadedLevelDetails.mapFolderPath, files, nf)
            For k = 0 To nf - 1
                If Len(files(k)) > 0 Then
                    printLog_ pCustomLogger, 1237, LogLevel_Info, "loadPackage", "Map folder file: %s", files(k)
                End If
            Next

            ' Successfully loaded level details from a .lev file in the root
            Set result = New MapPackage
            result.initLoaded loadedLevelDetails, LevelFormat_LEV, pathToMapPackage, pMapIO, logger, False
            If Len(originalGenerator) > 0 Then
                result.setOriginalGenerator originalGenerator
            End If

            If idx < rootLevCount - 1 Then
                ' There are more possible .lev files to load??
                printLog_ pCustomLogger, 1256, LogLevel_Warning, "loadPackage", "Loaded level details from file ""%s"", but there are %zu additional *.lev file(s) in the base path. Skipping the rest...", rootLevFiles(idx), rootLevCount - 1 - idx
            End If
            Set MapPackage_loadPackage = result
            Exit Function
        End If
    Next
    printLog_ pCustomLogger, 1261, LogLevel_Error, "loadPackage", "Found %zu root *.lev files (older format), but unable to load level details from any of them.", rootLevCount
    Set MapPackage_loadPackage = Nothing
End Function

Public Function safeDirectoryPathCompare(ByVal folderPath As String, ByVal desiredPath As String, ByVal desiredPathLen As Long, ByVal pathSeparator As String) As Boolean
    Dim folderPathLen As Long
    If Len(folderPath) = 0 Then
        safeDirectoryPathCompare = False
        Exit Function
    End If
    If Left$(folderPath, desiredPathLen) = Left$(desiredPath, desiredPathLen) And Len(folderPath) >= desiredPathLen Then
        folderPathLen = Len(folderPath)
        If folderPathLen = desiredPathLen Then
            safeDirectoryPathCompare = True ' found match
            Exit Function
        ElseIf folderPathLen = (desiredPathLen + Len(pathSeparator)) Then
            If Mid$(folderPath, desiredPathLen + 1, Len(pathSeparator)) = pathSeparator Then
                safeDirectoryPathCompare = True ' found match
                Exit Function
            End If
        End If
    End If

    safeDirectoryPathCompare = False
End Function
