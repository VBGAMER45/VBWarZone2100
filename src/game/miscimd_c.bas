Attribute VB_Name = "miscimd_c"
' src/miscimd.cpp, miscimd.h: the miscellaneous models (explosions, debris,
' blips, the assembly point numbers). A model is WzGfx's handle: an
' iIMDBaseShape * (modelGet), its display model an iIMDShape *; 0 = nullptr.
' miscimd.h's extern iIMDShape *explosionSmallImd ... are declared but never
' defined in 4.7.0 (the MI_ indices replaced them): not ported.
Option Explicit

Public Const MAX_DEBRIS As Long = 5
Public Const MAX_WRECKAGE As Long = 5

Public Const MAX_FACTORY_FLAG_IMDS As Long = 32

Public Type MISC_IMD
    pImd As Long                        ' iIMDBaseShape *
    pName As String                     ' const char *
End Type

Public Enum MI_INDEX                    ' (VB) miscimd.h's anonymous enum
    MI_EXPLOSION_SMALL
    MI_EXPLOSION_MEDIUM
    MI_CONSTRUCTION
    MI_SMALL_SMOKE
    MI_BABA_HEAD
    MI_BABA_LEGS
    MI_BABA_ARM
    MI_BABA_BODY
    MI_CYBORG_HEAD
    MI_CYBORG_LEGS
    MI_CYBORG_ARM
    MI_CYBORG_BODY
    MI_WATER
    MI_DROID_DAMAGE
    MI_SMALL_STEAM
    MI_PLASMA
    MI_BLIP
    MI_SHADOW
    MI_TRANSPORTER_SHADOW
    MI_BLOOD
    MI_TRAIL
    MI_FLAME
    MI_TESLA
    MI_MFLARE
    MI_RAIN
    MI_SNOW
    MI_SPLASH
    MI_KICK
    MI_LANDING
    MI_SHOCK
    MI_BLIP_ENEMY
    MI_BLIP_RESOURCE
    MI_BLIP_ARTEFACT

    MI_WRECK0
    MI_WRECK1
    MI_WRECK2
    MI_WRECK3
    MI_WRECK4

    MI_DEBRIS0
    MI_DEBRIS1
    MI_DEBRIS2
    MI_DEBRIS3
    MI_DEBRIS4

    MI_FIREWORK


    MI_TOO_MANY
End Enum

' Our great big array of imds
Private miscImds(0 To MI_TOO_MANY) As MISC_IMD
Private miscImdsReady As Boolean

Public pAssemblyPointIMDs(0 To NUM_FLAG_TYPES - 1, 0 To MAX_FACTORY_FLAG_IMDS - 1) As Long   ' iIMDBaseShape *

' (VB) static MISC_IMD miscImds[] = { {nullptr, "fxsexp"}, ... }: the table's initialiser
Private Sub initMiscImdsTable()
    Dim n As Variant, i As Long
    If miscImdsReady Then Exit Sub
    miscImdsReady = True
    i = 0
    For Each n In Array( _
        "fxsexp", "fxlexp", "fxdust", "fxsmoke", "parthead", "partlegs", "partarm", "partbody", _
        "cybitrkt", "cybitlg1", "cybitgun", "cybitbod", "fxatexp", "fxssteam", "fxssteam", "fxplasma", _
        "fxblip", "cyshadow", "Mitrnshd", "fxblood", "fxssmoke", "fxft", "fxpower", "fxmflare", _
        "mirain", "misnow", "fxssplsh", "fxexpdrt", "fxlightr", "fxl3dshk", "blipenm", "blipres", _
        "blipart", "miwrek1", "miwrek2", "miwrek3", "miwrek4", "miwrek5", "midebr1", "midebr2", _
        "midebr3", "midebr4", "midebr5", "fxflecht", "END_OF_IMD_LIST")
        ' explosionSmallImd MI_EXPLOSION_SMALL, explosionMediumImd, constructionImd, smallSmokeImd, babaHeadImd,
        ' babaLegsImd, babaArmImd, babaBodyImd, cyborgHeadImd, cyborgLegsImd, cyborgArmImd, cyborgBodyImd,
        ' waterImd, droidDamageImd, smallSteamImd, plasmaImd, blipImd, shadowImd, transporterShadowImd,
        ' bloodImd, trailImd, flameImd, teslaImd, mFlareImd, rainImd, snowImd, splashImd, kickImd,
        ' landingImd, shockImd, proximityImds[0..2] (MI_BLIP_ENEMY ...), wreckageImds[0..4] (MI_WRECK0 ...),
        ' debrisImds[0..4] (MI_DEBRIS0 ...), met hit - for repair centre MI_FIREWORK
        miscImds(i).pImd = 0
        miscImds(i).pName = n
        i = i + 1
    Next
End Sub

' -------------------------------------------------------------------------------
' Load up all the imds into an array
Private Function multiLoadMiscImds() As Boolean
    Dim i As Long, bMoreToProcess As Boolean, name_ As String

    initMiscImdsTable
    i = 0
    bMoreToProcess = True

    ' Go thru' the list
    Do While bMoreToProcess
        name_ = Left$(miscImds(i).pName & ".pie", 14)   ' char name[15]; // hopefully!
        ' see if the resource loader can find it
        miscImds(i).pImd = modelGet(name_)
        ' If it didn't get it then...
        If miscImds(i).pImd = 0 Then
            ' Say which one and return false
            debug_ 110, LOG_ERROR, "multiLoadMiscImds", "Can't find misselaneous PIE file : %s", miscImds(i).pName
            ASSERT_FAILURE "false", "src/miscimd.cpp:111", "multiLoadMiscImds", "NULL PIE"
            multiLoadMiscImds = False
            Exit Function
        End If
        '	If the next one's the end one, then get out now.
        '	This is cos strcmp will return 0 only at end of list
        i = i + 1
        bMoreToProcess = (miscImds(i).pName <> "END_OF_IMD_LIST")
    Loop
    multiLoadMiscImds = True
End Function

' -------------------------------------------------------------------------------
' Returns a pointer to the imd from a #define number passed in - see above
Public Function getImdFromIndex(ByVal index As Long) As Long
    initMiscImdsTable
    If Not (ULT(index, MI_TOO_MANY)) Then ASSERT_FAILURE "index < MI_TOO_MANY", "src/miscimd.cpp:125", "getImdFromIndex", "Out of range index in getImdFromIndex"
    getImdFromIndex = miscImds(index).pImd
End Function

' DISPLAY-ONLY
Public Function getDisplayImdFromIndex(ByVal index As Long) As Long
    Dim pBaseIMD As Long
    pBaseIMD = getImdFromIndex(index)
    If pBaseIMD <> 0 Then getDisplayImdFromIndex = iIMDBaseShape_displayModel(pBaseIMD) Else getDisplayImdFromIndex = 0
End Function

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
Public Function getRandomDebrisImd() As Long
    Dim DebrisIMD As Long
    DebrisIMD = getImdFromIndex(MI_DEBRIS0 + rand() Mod ((MI_DEBRIS4 - MI_DEBRIS0) + 1))
    If Not (DebrisIMD <> 0) Then ASSERT_FAILURE "DebrisIMD != nullptr", "src/miscimd.cpp:142", "getRandomDebrisImd", "getRandomDebrisImd : NULL PIE"
    getRandomDebrisImd = DebrisIMD
End Function

' -------------------------------------------------------------------------------

Private Function initMiscImd(ByVal i As Long, ByVal n As Long, ByVal nameFormat As String, ByVal flagType As Long) As Boolean
    Dim pieName As String
    pieName = va(nameFormat, n)
    If Len(pieName) > 99 Then pieName = Left$(pieName, 99)   ' char pieName[100]
    pAssemblyPointIMDs(flagType, i) = modelGet(pieName)
    If pAssemblyPointIMDs(flagType, i) = 0 Then
        debug_ 157, LOG_ERROR, "initMiscImd", "Can't find assembly point graphic %s for factory", pieName
        initMiscImd = False
        Exit Function
    End If
    initMiscImd = True
End Function

Public Function initMiscImds() As Boolean
    Dim i As Long, n As Long
    ' Do the new loading system
    multiLoadMiscImds

    ' Now load the multi array stuff
    For i = 0 To MAX_FACTORY_FLAG_IMDS - 1
        n = i + 1

        ' STATIC_ASSERT(MAX_FACTORY <= MAX_FACTORY_FLAG_IMDS);
        ' STATIC_ASSERT(MAX_FACTORY_FLAG_IMDS <= 32);  // Need to add more assembly point graphics, if increasing MAX_FACTORY_FLAG_IMDS.
        If Not initMiscImd(i, n, "minum%u.pie", FACTORY_FLAG) Then
            initMiscImds = False
            Exit Function
        End If
        If Not initMiscImd(i, n, "micnum%u.pie", CYBORG_FLAG) Then
            initMiscImds = False
            Exit Function
        End If
        If Not initMiscImd(i, n, "mivnum%u.pie", VTOL_FLAG) Then
            initMiscImds = False
            Exit Function
        End If
        If Not initMiscImd(i, 1, "mirnum%u.pie", REPAIR_FLAG) Then
            initMiscImds = False
            Exit Function
        End If
    Next

    initMiscImds = True
End Function
