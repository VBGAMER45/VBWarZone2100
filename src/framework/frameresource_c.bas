Attribute VB_Name = "frameresource_c"
' lib/framework/frameresource.cpp, frameresource.h: the resource system.
' .wrf files list the data files of a level by type; each type has the
' game's load and release functions (src/data.cpp), which this module calls
' through IFrameworkGame by their numbers (frame_c.frameGame); 0 is
' nullptr. void *pData is a Variant (Empty: nullptr).
' std::list<RES_TYPE *> psResTypes and each type's std::list<RES_DATA*>
' psRes grow with emplace_front: here arrays in insertion order, walked
' from the end (the front of the list) to the start. RES_DATA's aID is its
' own string (the C++ allocates it with the struct).
' The names are hashed from their UTF-8 bytes, as char (signed) values.
' resource_parser_c parses the .wrf file (the bison grammar and the flex
' lexer: resource_lexer_c), with aResDir and aCurrResDir from here.
Option Explicit

' Maximum number of characters in a resource type.
Private Const RESTYPE_MAXCHAR As Long = 20
Private Const PATH_MAX As Long = 260

Private Type RES_DATA
    pData As Variant                    ' pointer to the actual data
    blockID As Long                     ' which of the blocks is it in (so we can clear some of them...)
    HashedID As Long                    ' hashed version of the name of the id
    usage As Long                       ' Reference count
    aID As String                       ' ID of the resource - filename from the .wrf - e.g. "TRON.PIE"
End Type

Private Type RES_TYPE
    aType As String                     ' type string (e.g. "PIE"	 - just for debug use only, only applicable when loading from wrf (not wdg)
    buffLoad As Long                    ' routine to process the data for this type
    release As Long                     ' routine to release the data (NULL indicates none)
    psRes() As RES_DATA                 ' Linked list of data items of this type (its front is the last element)
    resCount As Long
    HashedType As Long                  ' hashed version of the name of the id - // a null hashedtype indicates end of list
    fileLoad As Long                    ' This isn't really used any more ?
End Type

' Structure for each file currently in use in the resource  ... probably only going to be one ... but we will handle upto MAXLOADEDRESOURCE
Private Type RESOURCEFILE
    pBuffer() As Byte                   ' a pointer to the data
    size As Long                        ' number of bytes
    typ As Long                         ' what type of resource is it
End Type

Private Const RESFILETYPE_EMPTY As Long = 0     ' empty entry
Private Const RESFILETYPE_LOADED As Long = 2    ' Loaded from a file (!)

Private Const MAXLOADEDRESOURCES As Long = 6

' next four used in HashPJW
Private Const BITS_IN_int As Long = 32
Private Const THREE_QUARTERS As Long = 24      ' ((BITS_IN_int * 3) / 4)
Private Const ONE_EIGHTH As Long = 4           ' (BITS_IN_int / 8)
Private Const HIGH_BITS As Long = &HF0000000   ' ( ~((UDWORD)(~0) >> ONE_EIGHTH ))

' Local prototypes
Private psResTypes() As RES_TYPE        ' its front is the last element
Private resTypeCount As Long

' The initial resource directory and the current resource directory
Public aResDir As String
Public aCurrResDir As String

' the current resource block ID
Private resBlockID As Long

' callback to resload screen.
Private resLoadCallback As Long

Private LastResourceFilename As String

Private LoadedResourceFiles(0 To MAXLOADEDRESOURCES - 1) As RESOURCEFILE

' sstrcpy into a char[PATH_MAX]
Private Function PathMax(ByVal s As String) As String
    If Len(s) > PATH_MAX - 1 Then PathMax = Left$(s, PATH_MAX - 1) Else PathMax = s
End Function

' The char (signed) values of c's UTF-8 bytes
Private Function CharsOf(ByVal c As String, chars() As Long) As Long
    Dim b() As Byte, n As Long, i As Long
    n = utf8_encode(c, b)
    If n > 0 Then ReDim chars(0 To n - 1)
    For i = 0 To n - 1
        If b(i) >= 128 Then chars(i) = CLng(b(i)) - 256 Else chars(i) = b(i)
    Next
    CharsOf = n
End Function

' One step of HashPJW: iHashValue = (iHashValue << ONE_EIGHTH) + c, folding the high bits
Private Function HashStep(ByVal iHashValue As Long, ByVal c As Long) As Long
    Dim i As Long
    iHashValue = FromU32(U32(SHL(iHashValue, ONE_EIGHTH)) + c)
    i = iHashValue And HIGH_BITS
    If i <> 0 Then
        iHashValue = (iHashValue Xor SHRU(i, THREE_QUARTERS)) And Not HIGH_BITS
    End If
    HashStep = iHashValue
End Function

'**************************************************************************
'
' HashString
'
' Adaptation of Peter Weinberger's (PJW) generic hashing algorithm listed
' in Binstock+Rex, "Practical Algorithms" p 69.
'
' Accepts string and returns hashed integer.
'
'**************************************************************************
Private Function HashString(ByVal c As String) As Long
    Dim iHashValue As Long, chars() As Long, n As Long, k As Long
    n = CharsOf(c, chars)
    iHashValue = 0
    For k = 0 To n - 1
        iHashValue = HashStep(iHashValue, chars(k))
    Next
    HashString = iHashValue
End Function

' Converts lower case ASCII characters into upper case characters
' \param c the character to convert
' \return an upper case ASCII character
Private Function upcaseASCII(ByVal c As Long) As Long
    ' If this is _not_ a lower case character simply return
    If c < 97 Or c > 122 Then
        upcaseASCII = c
    ' Otherwise substract 32 to make the lower case character an upper case one
    Else
        upcaseASCII = c - 32
    End If
End Function

Private Function HashStringIgnoreCase(ByVal c As String) As Long
    Dim iHashValue As Long, chars() As Long, n As Long, k As Long
    n = CharsOf(c, chars)
    iHashValue = 0
    For k = 0 To n - 1
        iHashValue = HashStep(iHashValue, upcaseASCII(chars(k)))
    Next
    HashStringIgnoreCase = iHashValue
End Function

' set the callback function for the res loader
Public Sub resSetLoadCallback(ByVal funcToCall As Long)
    resLoadCallback = funcToCall
End Sub

' do the callback for the resload display function
Public Sub resDoResLoadCallback()
    If resLoadCallback <> 0 Then
        frameGame.resLoadCallback resLoadCallback
    End If
End Sub

' Initialise the resource module
Public Function resInitialise() As Boolean
    If Not (resTypeCount = 0) Then ASSERT_FAILURE "psResTypes.empty()", "lib/framework/frameresource.cpp:152", "resInitialise", "resInitialise: resource module hasn't been shut down??"
    Erase psResTypes
    resTypeCount = 0
    resBlockID = 0
    resLoadCallback = 0

    ResetResourceFile

    resInitialise = True
End Function

' Shutdown the resource module
Public Sub resShutDown()
    If resTypeCount <> 0 Then
        debug_ 169, LOG_WZ, "resShutDown", "resShutDown: warning resources still allocated"
        resReleaseAll
    End If
End Sub

' force set the base resource directory
Public Sub resForceBaseDir(ByVal pResDir As String)
    aResDir = PathMax(pResDir)
    aCurrResDir = PathMax(pResDir)
End Sub

' set the base resource directory
Public Sub resSetBaseDir(ByVal pResDir As String)
    aResDir = PathMax(pResDir)
End Sub

' Parse the res file
Public Function resLoad(ByVal pResFile As String, ByVal blockID As Long) As Boolean
    Dim retval As Boolean, physfsfile As Long
    retval = True

    aCurrResDir = aResDir

    ' Note the block id number
    resBlockID = blockID

    debug_ 198, LOG_WZ, "resLoad", "resLoad: loading [directory: %s] %s", PHYSFS_getRealDir(pResFile) & "", pResFile

    ' Load the RES file; allocate memory for a wrf, and load it
    physfsfile = openLoadFile(pResFile, True)
    If physfsfile = 0 Then
        debug_ 205, LOG_FATAL, "resLoad", "Could not open file %s", pResFile
        resLoad = False
        Exit Function
    End If

    ' and parse it
    res_set_extra physfsfile
    If res_parse() <> 0 Then
        debug_ 213, LOG_FATAL, "resLoad", "Failed to parse %s", pResFile
        retval = False
    End If

    res_lex_destroy
    PHYSFS_close physfsfile

    resLoad = retval
End Function

' Allocate a RES_TYPE structure (#ifdef DEBUG's duplicate check: not in a release build)
Private Function resAlloc(ByVal pType As String) As RES_TYPE
    Dim psT As RES_TYPE
    ' setup the structure
    If Len(pType) > RESTYPE_MAXCHAR - 1 Then psT.aType = Left$(pType, RESTYPE_MAXCHAR - 1) Else psT.aType = pType
    psT.HashedType = HashString(psT.aType) ' store a hased version for super speed !
    resAlloc = psT
End Function

' psResTypes.emplace_front(psT)
Private Sub AddResType(psT As RES_TYPE)
    If resTypeCount = 0 Then
        ReDim psResTypes(0 To 31)
    ElseIf resTypeCount > UBound(psResTypes) Then
        ReDim Preserve psResTypes(0 To 2 * resTypeCount - 1)
    End If
    psResTypes(resTypeCount) = psT
    resTypeCount = resTypeCount + 1
End Sub

' Add a buffer load function for a file type
Public Function resAddBufferLoad(ByVal pType As String, ByVal buffLoad As Long, ByVal release As Long) As Boolean
    Dim psT As RES_TYPE
    psT = resAlloc(pType)

    psT.buffLoad = buffLoad
    psT.fileLoad = 0
    psT.release = release

    AddResType psT

    resAddBufferLoad = True
End Function

' Add a file name load function for a file type
Public Function resAddFileLoad(ByVal pType As String, ByVal fileLoad As Long, ByVal release As Long) As Boolean
    Dim psT As RES_TYPE
    psT = resAlloc(pType)

    psT.buffLoad = 0
    psT.fileLoad = fileLoad
    psT.release = release

    AddResType psT

    resAddFileLoad = True
End Function

' Make a string lower case (isupper: the C locale's A-Z), in place
Public Sub resToLower(pStr As String)
    Dim i As Long, c As Long
    For i = 1 To Len(pStr)
        c = AscW(Mid$(pStr, i, 1))
        If c >= 65 And c <= 90 Then Mid$(pStr, i, 1) = ChrW$(c - (65 - 97))
    Next
End Sub

'
' Returns the filename of the last resource file loaded
' The filename is always null terminated
'
Public Function GetLastResourceFilename() As String
    GetLastResourceFilename = LastResourceFilename
End Function

'
' Set the resource name of the last resource file loaded
'
Public Sub SetLastResourceFilename(ByVal pName As String)
    LastResourceFilename = PathMax(pName)
End Sub

' Clear out the resource list ... needs to be called during init.
Private Sub ResetResourceFile()
    Dim i As Long
    For i = 0 To MAXLOADEDRESOURCES - 1
        LoadedResourceFiles(i).typ = RESFILETYPE_EMPTY
    Next
End Sub

' Returns an empty resource entry or -1 if none exsist
Private Function FindEmptyResourceFile() As Long
    Dim i As Long
    For i = 0 To MAXLOADEDRESOURCES - 1
        If LoadedResourceFiles(i).typ = RESFILETYPE_EMPTY Then
            FindEmptyResourceFile = i
            Exit Function
        End If
    Next
    FindEmptyResourceFile = -1          ' ERROR
End Function

' Get a resource data file ... either loads it or just returns a pointer (NewResource: its index)
Private Function RetreiveResourceFile(ByVal ResourceName As String, NewResource As Long) As Boolean
    Dim ResID As Long, size As Long, pBuffer() As Byte

    ResID = FindEmptyResourceFile()
    If ResID = -1 Then
        RetreiveResourceFile = False    ' all resource files are full
        Exit Function
    End If

    NewResource = ResID

    ' This is needed for files that do not fit in the WDG cache ... (VAB file for example)
    If Not loadFile(ResourceName, pBuffer, size) Then
        RetreiveResourceFile = False
        Exit Function
    End If

    With LoadedResourceFiles(ResID)
        .typ = RESFILETYPE_LOADED
        .size = size
        .pBuffer = pBuffer
    End With
    RetreiveResourceFile = True
End Function

' Free up the file depending on what type it is
Private Sub FreeResourceFile(ByVal OldResource As Long)
    With LoadedResourceFiles(OldResource)
        Select Case .typ
        Case RESFILETYPE_LOADED
            Erase .pBuffer
        Case Else
            debug_ 392, LOG_WARNING, "FreeResourceFile", "resource not freed"
        End Select

        ' Remove from the list
        .typ = RESFILETYPE_EMPTY
    End With
End Sub

Private Function resDataInit(ByVal DebugName As String, ByVal DataIDHash As Long, pData As Variant, ByVal BlockID As Long) As RES_DATA
    Dim psRes As RES_DATA
    ' Copy over the identifying string
    psRes.aID = DebugName
    If IsObject(pData) Then Set psRes.pData = pData Else psRes.pData = pData
    psRes.blockID = BlockID
    psRes.HashedID = DataIDHash

    psRes.usage = 0

    resDataInit = psRes
End Function

'
' check if given file exists in a locale dependent subdir
' if so, modify given fileName to hold the locale dep. file,
' else do not change given fileName
'
Private Sub makeLocaleFile(fileName As String)  ' given string must have PATH_MAX size
    Dim language As String, localeFile As String
    language = getLanguage()

    If Len(language) = 0 Or Len(fileName) + Len(language) + 1 >= PATH_MAX Then    ' could not get language
        Exit Sub
    End If

    localeFile = PathMax("locale/" & language & "/" & fileName)

    If PHYSFS_exists(localeFile) <> 0 Then
        fileName = PathMax(localeFile)
        debug_ 454, LOG_WZ, "makeLocaleFile", "Found translated file: %s", fileName
    End If
End Sub

' The index of the type named pType (std::find_if from the front), or -1
Private Function FindResType(ByVal HashedType As Long) As Long
    Dim t As Long
    For t = resTypeCount - 1 To 0 Step -1
        If psResTypes(t).HashedType = HashedType Then
            FindResType = t
            Exit Function
        End If
    Next
    FindResType = -1
End Function

' psT->psRes.emplace_front(psRes)
Private Sub AddResData(ByVal t As Long, psRes As RES_DATA)
    With psResTypes(t)
        If .resCount = 0 Then
            ReDim .psRes(0 To 15)
        ElseIf .resCount > UBound(.psRes) Then
            ReDim Preserve .psRes(0 To 2 * .resCount - 1)
        End If
        .psRes(.resCount) = psRes
        .resCount = .resCount + 1
    End With
End Sub

Private Function SameData(a As Variant, b As Variant) As Boolean
    If IsObject(a) Or IsObject(b) Then
        If IsObject(a) And IsObject(b) Then SameData = (a Is b)
    ElseIf IsEmpty(a) Or IsEmpty(b) Then
        SameData = (IsEmpty(a) And IsEmpty(b))
    Else
        SameData = (a = b)
    End If
End Function

'
' Call the load function (registered in data.c)
' for this filetype
'
Public Function resLoadFile(ByVal pType As String, ByVal pFile As String) As Boolean
    Dim pData As Variant, aFileName As String, HashedName As Long, HashedType As Long
    Dim t As Long, r As Long, Resource As Long, psRes As RES_DATA
    HashedType = HashString(pType)

    ' Find the resource-type
    t = FindResType(HashedType)
    If t >= 0 Then
        If Not (psResTypes(t).aType = pType) Then ASSERT_FAILURE "strcmp(psT->aType, pType) == 0", "lib/framework/frameresource.cpp:477", "resLoadFile", "Hash collision ""%s"" vs ""%s""", psResTypes(t).aType, pType
    End If

    If t < 0 Then
        debug_ 485, LOG_WZ, "resLoadFile", "resLoadFile: Unknown type: %s", pType
        resLoadFile = False
        Exit Function
    End If

    ' Check for duplicates
    HashedName = HashStringIgnoreCase(pFile)
    With psResTypes(t)
        For r = .resCount - 1 To 0 Step -1
            If .psRes(r).HashedID = HashedName Then
                If Not (StrComp(.psRes(r).aID, pFile, vbTextCompare) = 0) Then ASSERT_FAILURE "_stricmp(psRes->aID, pFile) == 0", "lib/framework/frameresource.cpp:496", "resLoadFile", "Hash collision ""%s"" vs ""%s""", .psRes(r).aID, pFile
                debug_ 497, LOG_WZ, "resLoadFile", "Duplicate file name: %s (hash %x) for type %s", pFile, HashedName, .aType
                ' assume that they are actually both the same and silently fail
                ' lovely little hack to allow some files to be loaded from disk (believe it or not!).
                resLoadFile = True
                Exit Function
            End If
        Next
    End With

    ' Create the file name
    If Len(aCurrResDir) + Len(pFile) + 1 >= PATH_MAX Then
        debug_ 508, LOG_ERROR, "resLoadFile", "resLoadFile: Filename too long!! %s%s", aCurrResDir, pFile
        resLoadFile = False
        Exit Function
    End If
    aFileName = aCurrResDir
    aFileName = PathMax(aFileName & pFile)

    makeLocaleFile aFileName    ' check for translated file

    SetLastResourceFilename pFile ' Save the filename in case any routines need it

    ' load the resource
    If psResTypes(t).buffLoad <> 0 Then
        ' Load the file in a buffer
        If Not RetreiveResourceFile(aFileName, Resource) Then
            debug_ 526, LOG_ERROR, "resLoadFile", "resLoadFile: Unable to retrieve resource - %s", aFileName
            resLoadFile = False
            Exit Function
        End If

        ' Now process the buffer data
        If Not frameGame.resBufferLoad(psResTypes(t).buffLoad, LoadedResourceFiles(Resource).pBuffer, LoadedResourceFiles(Resource).size, pData) Then
            ASSERT_FAILURE "false", "lib/framework/frameresource.cpp:533", "resLoadFile", "The load function for resource type ""%s"" failed for file ""%s""", pType, pFile
            FreeResourceFile Resource
            If psResTypes(t).release <> 0 Then
                frameGame.resRelease psResTypes(t).release, pData
            End If
            resLoadFile = False
            Exit Function
        End If

        FreeResourceFile Resource
    ElseIf psResTypes(t).fileLoad <> 0 Then
        ' Process data directly from file
        If Not frameGame.resFileLoad(psResTypes(t).fileLoad, aFileName, pData) Then
            ASSERT_FAILURE "false", "lib/framework/frameresource.cpp:549", "resLoadFile", "The load function for resource type ""%s"" failed for file ""%s""", pType, pFile
            If psResTypes(t).release <> 0 Then
                frameGame.resRelease psResTypes(t).release, pData
            End If
            resLoadFile = False
            Exit Function
        End If
    End If

    resDoResLoadCallback        ' do callback.

    ' Set up the resource structure if there is something to store
    If Not IsEmpty(pData) Then
        ' LastResourceFilename may have been changed (e.g. by TEXPAGE loading)
        psRes = resDataInit(GetLastResourceFilename(), HashStringIgnoreCase(GetLastResourceFilename()), pData, resBlockID)

        ' Add the resource to the list
        AddResData t, psRes
    End If
    resLoadFile = True
End Function

' Return the resource for a type and hashedname
Public Function resGetDataFromHash(ByVal pType As String, ByVal HashedID As Long) As Variant
    Dim HashedType As Long, t As Long, r As Long
    ' Find the correct type
    HashedType = HashString(pType)

    t = FindResType(HashedType)
    If Not (t >= 0) Then ASSERT_FAILURE "resTypeIt != psResTypes.end()", "lib/framework/frameresource.cpp:590", "resGetDataFromHash", "resGetDataFromHash: Unknown type: %s", pType
    If t < 0 Then
        Exit Function
    End If

    With psResTypes(t)
        For r = .resCount - 1 To 0 Step -1
            If .psRes(r).HashedID = HashedID Then Exit For
        Next

        If Not (r >= 0) Then ASSERT_FAILURE "res != psT->psRes.end()", "lib/framework/frameresource.cpp:602", "resGetDataFromHash", "resGetDataFromHash: Unknown ID: %0x Type: %s", HashedID, pType
        If r < 0 Then
            Exit Function
        End If

        .psRes(r).usage = .psRes(r).usage + 1

        If IsObject(.psRes(r).pData) Then Set resGetDataFromHash = .psRes(r).pData Else resGetDataFromHash = .psRes(r).pData
    End With
End Function

' Return the resource for a type and ID
Public Function resGetData(ByVal pType As String, ByVal pID As String) As Variant
    Dim data As Variant
    VarAssign data, resGetDataFromHash(pType, HashStringIgnoreCase(pID))
    If Not (Not IsEmpty(data)) Then ASSERT_FAILURE "data != nullptr", "lib/framework/frameresource.cpp:618", "resGetData", "resGetData: Unable to find data for %s type %s", pID, pType
    If IsObject(data) Then Set resGetData = data Else resGetData = data
End Function

Public Function resGetHashfromData(ByVal pType As String, pData As Variant, pHash As Long) As Boolean
    Dim HashedType As Long, t As Long, r As Long
    ' Find the correct type
    HashedType = HashString(pType)

    t = FindResType(HashedType)
    If Not (t >= 0) Then
        ASSERT_FAILURE "resTypeIt != psResTypes.end()", "lib/framework/frameresource.cpp:632", "resGetHashfromData", "Unknown type: %x", HashedType
        resGetHashfromData = False
        Exit Function
    End If

    ' Find the resource
    With psResTypes(t)
        For r = .resCount - 1 To 0 Step -1
            If SameData(.psRes(r).pData, pData) Then Exit For
        Next

        If r < 0 Then
            ASSERT_FAILURE "false", "lib/framework/frameresource.cpp:643", "resGetHashfromData", "resGetHashfromData:: couldn't find data for type %x\n", HashedType
            resGetHashfromData = False
            Exit Function
        End If

        pHash = .psRes(r).HashedID
    End With

    resGetHashfromData = True
End Function

' type (vbNullString) or data (Empty) nullptr: ""
Public Function resGetNamefromData(ByVal typ As String, data As Variant) As String
    Dim HashedType As Long, t As Long, r As Long

    If StrPtr(typ) = 0 Or IsEmpty(data) Then
        resGetNamefromData = ""
        Exit Function
    End If

    ' Find the correct type
    HashedType = HashString(typ)

    ' Find the resource table for the given type
    t = FindResType(HashedType)
    If t < 0 Then
        ASSERT_FAILURE "false", "lib/framework/frameresource.cpp:671", "resGetNamefromData", "resGetHashfromData: Unknown type: %x", HashedType
        resGetNamefromData = ""
        Exit Function
    End If

    ' Find the resource in the resource table
    With psResTypes(t)
        For r = .resCount - 1 To 0 Step -1
            If SameData(.psRes(r).pData, data) Then Exit For
        Next

        If r < 0 Then
            ASSERT_FAILURE "false", "lib/framework/frameresource.cpp:684", "resGetNamefromData", "resGetHashfromData:: couldn't find data for type %x\n", HashedType
            resGetNamefromData = ""
            Exit Function
        End If

        resGetNamefromData = .psRes(r).aID
    End With
End Function

' Simply returns true if a resource is present
Public Function resPresent(ByVal pType As String, ByVal pID As String) As Boolean
    Dim HashedType As Long, HashedID As Long, t As Long, r As Long
    ' Find the correct type
    HashedType = HashString(pType)

    t = FindResType(HashedType)
    ' Bow out if unrecognised type
    If Not (t >= 0) Then ASSERT_FAILURE "resTypeIt != psResTypes.end()", "lib/framework/frameresource.cpp:702", "resPresent", "resPresent: Unknown type"
    If t < 0 Then
        resPresent = False
        Exit Function
    End If

    HashedID = HashStringIgnoreCase(pID)
    With psResTypes(t)
        For r = .resCount - 1 To 0 Step -1
            If .psRes(r).HashedID = HashedID Then
                resPresent = True
                Exit Function
            End If
        Next
    End With
End Function

' Release all the resources currently loaded and the resource load functions
Public Sub resReleaseAll()
    resReleaseAllData

    Erase psResTypes
    resTypeCount = 0
End Sub

' Release all the resources currently loaded but keep the resource load functions
Public Sub resReleaseAllData()
    Dim t As Long, r As Long
    For t = resTypeCount - 1 To 0 Step -1
        With psResTypes(t)
            For r = .resCount - 1 To 0 Step -1
                If .psRes(r).usage = 0 Then
                    debug_ 741, LOG_NEVER, "resReleaseAllData", "resReleaseAllData: %s resource: %s(%04x) not used", .aType, .psRes(r).aID, .psRes(r).HashedID
                End If
                If .release <> 0 Then
                    frameGame.resRelease .release, .psRes(r).pData
                End If
            Next
            Erase .psRes
            .resCount = 0
        End With
    Next
End Sub

' release the data for a particular block ID
Public Sub resReleaseBlockData(ByVal blockID As Long)
    Dim t As Long, r As Long, k As Long
    For t = resTypeCount - 1 To 0 Step -1
        With psResTypes(t)
            For r = .resCount - 1 To 0 Step -1
                If .psRes(r).blockID <> blockID Then
                    GoTo continue_
                End If
                If .psRes(r).usage = 0 Then
                    debug_ 769, LOG_NEVER, "resReleaseBlockData", "resReleaseBlockData: %s resource: %s(%04x) not used", .aType, .psRes(r).aID, .psRes(r).HashedID
                End If
                If .release <> 0 Then
                    frameGame.resRelease .release, .psRes(r).pData
                End If

                ' psT->psRes.erase(resIt)
                For k = r To .resCount - 2
                    .psRes(k) = .psRes(k + 1)
                Next
                .resCount = .resCount - 1
                .psRes(.resCount) = EmptyResData()
continue_:
            Next
        End With
    Next
End Sub

Private Function EmptyResData() As RES_DATA
End Function

' dst = src, for an object as for a value
Public Sub VarAssign(dst As Variant, src As Variant)
    If IsObject(src) Then Set dst = src Else dst = src
End Sub
