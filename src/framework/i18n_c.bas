Attribute VB_Name = "i18n_c"
' lib/framework/i18n.cpp, the Windows branch with ENABLE_NLS (as Warzone's
' Windows build): the language table, getLanguage, getLanguageName,
' setLocaleWindows, setLanguage, getLocales, wzBindTextDomain, initI18n,
' getCompileDate (on build.py's BUILD_DATE; C: __DATE__ in ISO 8601).
' gettext (libintl) is not ported, so the branches without _LIBINTL_H are
' taken, bindtextdomain / textdomain stand in for libintl's, and the text is
' English: gettext(s) is s. The language names are written with \uXXXX
' escapes (VB source is ANSI) and decoded once.
Option Explicit

Private Declare Function GetThreadLocale Lib "kernel32" () As Long
Private Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long

Private Const PACKAGE As String = "warzone2100"
Private Const WZ_LOCALEDIR As String = "./locale"       ' the Windows build's CMAKE_INSTALL_LOCALEDIR, relative

' winnt.h language and sub-language ids
Private Const LANG_NEUTRAL As Long = &H0
Private Const LANG_ARABIC As Long = &H1
Private Const LANG_BULGARIAN As Long = &H2
Private Const LANG_CATALAN As Long = &H3
Private Const LANG_CHINESE As Long = &H4
Private Const LANG_CZECH As Long = &H5
Private Const LANG_DANISH As Long = &H6
Private Const LANG_GERMAN As Long = &H7
Private Const LANG_GREEK As Long = &H8
Private Const LANG_ENGLISH As Long = &H9
Private Const LANG_SPANISH As Long = &HA
Private Const LANG_FINNISH As Long = &HB
Private Const LANG_FRENCH As Long = &HC
Private Const LANG_HUNGARIAN As Long = &HE
Private Const LANG_ITALIAN As Long = &H10
Private Const LANG_KOREAN As Long = &H12
Private Const LANG_DUTCH As Long = &H13
Private Const LANG_NORWEGIAN As Long = &H14
Private Const LANG_POLISH As Long = &H15
Private Const LANG_PORTUGUESE As Long = &H16
Private Const LANG_ROMANIAN As Long = &H18
Private Const LANG_RUSSIAN As Long = &H19
Private Const LANG_CROATIAN As Long = &H1A
Private Const LANG_SERBIAN As Long = &H1A
Private Const LANG_SLOVAK As Long = &H1B
Private Const LANG_TURKISH As Long = &H1F
Private Const LANG_INDONESIAN As Long = &H21
Private Const LANG_UKRAINIAN As Long = &H22
Private Const LANG_SLOVENIAN As Long = &H24
Private Const LANG_ESTONIAN As Long = &H25
Private Const LANG_LITHUANIAN As Long = &H27
Private Const LANG_IRISH As Long = &H3C
Private Const LANG_FRISIAN As Long = &H62
Private Const SUBLANG_DEFAULT As Long = &H1
Private Const SUBLANG_GERMAN As Long = &H1
Private Const SUBLANG_ENGLISH_UK As Long = &H2
Private Const SUBLANG_SPANISH As Long = &H1
Private Const SUBLANG_FRENCH As Long = &H1
Private Const SUBLANG_FRISIAN_NETHERLANDS As Long = &H1
Private Const SUBLANG_IRISH_IRELAND As Long = &H2
Private Const SUBLANG_ITALIAN As Long = &H1
Private Const SUBLANG_DUTCH As Long = &H1
Private Const SUBLANG_PORTUGUESE_BRAZILIAN As Long = &H1
Private Const SUBLANG_SERBIAN_LATIN As Long = &H2
Private Const SUBLANG_CHINESE_SIMPLIFIED As Long = &H2
Private Const SUBLANG_CHINESE_TRADITIONAL As Long = &H1
Private Const SORT_DEFAULT As Long = &H0

Private Type LanguageMapEntry
    language As String
    nm As String                        ' C: name
    localeFilename As String            ' the actual filename on disk, minus the extension (ex. "en_GB")
    usPrimaryLanguage As Long
    usSubLanguage As Long
End Type

Private map() As LanguageMapEntry
Private mapSize As Long

Private selectedLanguage As Long
Private canUseLANGUAGEEnvVar As Boolean

' "\uXXXX" escapes to characters
Private Function Unescape(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW$(CLng("&H" & Mid$(s, p + 2, 4))) & Mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    Unescape = s
End Function

Private Sub AddLanguage(ByVal language As String, ByVal nm As String, ByVal localeFilename As String, ByVal usPrimaryLanguage As Long, ByVal usSubLanguage As Long)
    If mapSize = 0 Then
        ReDim map(0 To 39)
    ElseIf mapSize > UBound(map) Then
        ReDim Preserve map(0 To 2 * mapSize - 1)
    End If
    map(mapSize).language = language
    map(mapSize).nm = Unescape(nm)
    map(mapSize).localeFilename = localeFilename
    map(mapSize).usPrimaryLanguage = usPrimaryLanguage
    map(mapSize).usSubLanguage = usSubLanguage
    mapSize = mapSize + 1
End Sub

'  See msdn.microsoft.com for this stuff, esp.
'  http://msdn.microsoft.com/en-us/library/ms693062%28VS.85,printer%29.aspx
'  http://msdn.microsoft.com/en-us/library/dd318693%28VS.85,printer%29.aspx
Private Sub InitMap()
    If mapSize > 0 Then Exit Sub
    ' Language names (http://en.wikipedia.org/wiki/List_of_ISO_639-2_codes)
    AddLanguage "", "System locale", "", LANG_NEUTRAL, SUBLANG_DEFAULT
    AddLanguage "ar_SA", "العربية", "ar_SA", LANG_ARABIC, SUBLANG_DEFAULT
    AddLanguage "bg", "български език", "bg_BG", LANG_BULGARIAN, SUBLANG_DEFAULT
    AddLanguage "ca_ES", "Català", "ca_ES", LANG_CATALAN, SUBLANG_DEFAULT
    AddLanguage "cs_CZ", "Česky", "cs", LANG_CZECH, SUBLANG_DEFAULT
    AddLanguage "da", "Dansk", "da", LANG_DANISH, SUBLANG_DEFAULT
    AddLanguage "de", "Deutsch", "de", LANG_GERMAN, SUBLANG_GERMAN
    AddLanguage "el_GR", "Ελληνικά", "el", LANG_GREEK, SUBLANG_DEFAULT
    AddLanguage "en", "English", "en", LANG_ENGLISH, SUBLANG_DEFAULT
    AddLanguage "en_GB", "British English", "en_GB", LANG_ENGLISH, SUBLANG_ENGLISH_UK
    AddLanguage "es", "Español", "es", LANG_SPANISH, SUBLANG_SPANISH
    AddLanguage "et_EE", "Eesti Keel", "et_EE", LANG_ESTONIAN, SUBLANG_DEFAULT
    AddLanguage "fi", "Suomi", "fi", LANG_FINNISH, SUBLANG_DEFAULT
    AddLanguage "fr", "Français", "fr", LANG_FRENCH, SUBLANG_FRENCH
    ' Our Frisian translation is the "West Frisian" variation of it. This
    ' variation is mostly spoken in the Dutch province Friesland (Fryslan
    ' in Frisian) and has ISO 639-3 code "fry".
    AddLanguage "fy_NL", "Frysk", "fy", LANG_FRISIAN, SUBLANG_FRISIAN_NETHERLANDS
    AddLanguage "ga_IE", "Imruagadh", "ga", LANG_IRISH, SUBLANG_IRISH_IRELAND
    AddLanguage "hr", "Hrvatski", "hr", LANG_CROATIAN, SUBLANG_DEFAULT
    AddLanguage "hu", "Magyar", "hu", LANG_HUNGARIAN, SUBLANG_DEFAULT
    AddLanguage "id", "Bahasa Indonesia", "id", LANG_INDONESIAN, SUBLANG_DEFAULT
    AddLanguage "it", "Italiano", "it", LANG_ITALIAN, SUBLANG_ITALIAN
    AddLanguage "ko_KR", "한국어", "ko", LANG_KOREAN, SUBLANG_DEFAULT
    AddLanguage "lt", "Lietuvių kalba", "lt", LANG_LITHUANIAN, SUBLANG_DEFAULT
    ' MSDN uses "no"...
    AddLanguage "nb_NO", "Norsk", "nb", LANG_NORWEGIAN, SUBLANG_DEFAULT
    AddLanguage "nl", "Nederlands", "nl", LANG_DUTCH, SUBLANG_DUTCH
    AddLanguage "pl", "Polski", "pl", LANG_POLISH, SUBLANG_DEFAULT
    AddLanguage "pt_BR", "Português Brasileiro", "pt_BR", LANG_PORTUGUESE, SUBLANG_PORTUGUESE_BRAZILIAN
    AddLanguage "pt", "Português", "pt", LANG_PORTUGUESE, SUBLANG_DEFAULT
    AddLanguage "ro", "Română", "ro", LANG_ROMANIAN, SUBLANG_DEFAULT
    AddLanguage "ru", "Русский", "ru", LANG_RUSSIAN, SUBLANG_DEFAULT
    AddLanguage "sk", "Slovensky", "sk", LANG_SLOVAK, SUBLANG_DEFAULT
    AddLanguage "sl_SI", "Slovenski", "sl", LANG_SLOVENIAN, SUBLANG_DEFAULT
    AddLanguage "sr_RS", "Srpski", "sr_CS", LANG_SERBIAN, SUBLANG_SERBIAN_LATIN ' Due to how gettext-libintl maps Win32 LangIDs to locales, use sr_CS for Windows
    AddLanguage "tr", "Türkçe", "tr", LANG_TURKISH, SUBLANG_DEFAULT
    AddLanguage "uk_UA", "Українська", "uk_UA", LANG_UKRAINIAN, SUBLANG_DEFAULT
    AddLanguage "zh_CN", "汉语", "zh_CN", LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED
    AddLanguage "zh_TW", "漢語", "zh_TW", LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL
End Sub

'!
' Return the language part of the selected locale
Public Function getLanguage() As String
    Dim usPrimaryLanguage As Long, i As Long

    InitMap
    usPrimaryLanguage = GetThreadLocale() And &H3FF&   ' PRIMARYLANGID(LANGIDFROMLCID(...))

    If selectedLanguage = 0 Then
        getLanguage = "" ' Return empty string for system default
        Exit Function
    End If

    For i = 0 To mapSize - 1
        If usPrimaryLanguage = map(i).usPrimaryLanguage Then
            getLanguage = map(i).language
            Exit Function
        End If
    Next

    getLanguage = ""
End Function

Public Function getLanguageName() As String
    Dim language As String, i As Long
    language = getLanguage()

    For i = 0 To mapSize - 1
        If language = map(i).language Then
            getLanguageName = map(i).nm         ' gettext(map[i].name)
            Exit Function
        End If
    Next

    getLanguageName = language
End Function

Private Function setLocaleWindows(ByVal usPrimaryLanguage As Long, ByVal usSubLanguage As Long) As Boolean
    Dim success As Boolean, langid As Long
    langid = (usSubLanguage * &H400&) Or usPrimaryLanguage                         ' MAKELANGID
    success = (SetThreadLocale((SORT_DEFAULT * &H10000) Or langid) >= 0)            ' SUCCEEDED(SetThreadLocale(MAKELCID(...))): a BOOL is never negative

    If Not success Then
        wz_info 327, "setLocaleWindows", "Failed to set locale to ""%d""", usPrimaryLanguage
    Else
        debug_ 331, LOG_WZ, "setLocaleWindows", "Requested locale ""%d""", usPrimaryLanguage
    End If

    ' (C: setlocale(LC_NUMERIC, "C"): the port's number formatting does not use the locale)

    setLocaleWindows = success
End Function

Public Function setLanguage(ByVal language As String) As Boolean
    Dim i As Long

    InitMap
    For i = 0 To mapSize - 1
        If language = map(i).language Then
            selectedLanguage = i
            debug_ 439, LOG_WZ, "setLanguage", "Setting language to ""%s"" (%s)", map(i).nm, map(i).language

            setLanguage = setLocaleWindows(map(i).usPrimaryLanguage, map(i).usSubLanguage)
            Exit Function
        End If
    Next

    debug_ 451, LOG_ERROR, "setLanguage", "Requested language ""%s"" not supported.", language

    setLanguage = False
End Function

' C: std::vector<locale_info> getLocales(): the codes and names in code(0 ..) and nm(0 ..), and their count
Public Function getLocales(code() As String, nm() As String) As Long
    Dim i As Long
    InitMap
    ReDim code(0 To mapSize - 1)
    ReDim nm(0 To mapSize - 1)
    For i = 0 To mapSize - 1
        code(i) = map(i).language
        nm(i) = map(i).nm
    Next
    getLocales = mapSize
End Function

' libintl's bindtextdomain, which this port does not have: the directory, as bound
Private Function bindtextdomain(ByVal domainname As String, ByVal dirname As String) As String
    bindtextdomain = dirname
End Function

Public Function wzBindTextDomain(ByVal domainname As String, ByVal dirname As String) As String
    Dim pResult As String, dn As String, dr As String
    If StrPtr(domainname) <> 0 Then dn = domainname Else dn = "NULL"
    If StrPtr(dirname) <> 0 Then dr = dirname Else dr = "NULL"
    debug_ 470, LOG_WZ, "wzBindTextDomain", "bindTextDomain(%s, %s)", dn, dr

    ' call the normal bindtextdomain function
    pResult = bindtextdomain(domainname, dirname)
    If StrPtr(pResult) = 0 Then
        wzBindTextDomain = ""
        Exit Function
    End If
    wzBindTextDomain = pResult
End Function

Private Function checkSupportsLANGUAGEenvVarOverride() As Boolean
    debug_ 574, LOG_WZ, "checkSupportsLANGUAGEenvVarOverride", "No support for LANGUAGE env var override (not GNU gettext?)"
    checkSupportsLANGUAGEenvVarOverride = False
End Function

Private Function wzBindAllWZTextDomains(ByVal resourcePath As String) As Boolean
    Dim textdomainDirectory As String

    ' Bind the "main" message catalog ("warzone2100")
    textdomainDirectory = wzBindTextDomain(PACKAGE, resourcePath)
    ' (bind_textdomain_codeset(PACKAGE, "UTF-8"))
    If Len(textdomainDirectory) = 0 Then
        ' Failed to bind main message catalog - treat as critical failure
        Exit Function
    End If

    ' Bind the guide topics message catalog
    textdomainDirectory = wzBindTextDomain(PACKAGE & "_guide", resourcePath)
    ' (bind_textdomain_codeset(PACKAGE "_guide", "UTF-8"))
    If Len(textdomainDirectory) = 0 Then
        ' Non-fatal failure
        debug_ 635, LOG_INFO, "wzBindAllWZTextDomains", "initI18n: Failed to bind guide message catalog"
    End If

    wzBindAllWZTextDomains = True
End Function

Public Sub initI18n()
    Dim boundAtLeastMainTextDomain As Boolean, defaultLanguage As String
    Dim prefixDir As String, dirSeparator As String, localeDir As String

    defaultLanguage = "" ' "" is system default (in most cases)

    ' New locale-dir setup (CMake)
    ' Treat WZ_LOCALEDIR as a relative path - append to the install PREFIX
    prefixDir = getWZInstallPrefix()
    debug_ 684, LOG_WZ, "initI18n", "Install prefix: %s", prefixDir
    dirSeparator = PHYSFS_getDirSeparator()
    localeDir = prefixDir & dirSeparator & WZ_LOCALEDIR
    boundAtLeastMainTextDomain = wzBindAllWZTextDomains(localeDir)
    If Not boundAtLeastMainTextDomain Then
        debug_ 699, LOG_ERROR, "initI18n", "initI18n: bindtextdomain failed!"
    End If

    ' Set default text domain
    ' (textdomain(PACKAGE))
    debug_ 704, LOG_WZ, "initI18n", "textdomain: %s", PACKAGE

    ' Should come *after* bindTextDomain
    canUseLANGUAGEEnvVar = checkSupportsLANGUAGEenvVarOverride()

    If Not setLanguage(defaultLanguage) Then ' set to system default
        ' no system default?
        debug_ 712, LOG_ERROR, "initI18n", "initI18n: No system language found"
    End If
End Sub

' convert macro __DATE__ to ISO 8601 format
Public Function getCompileDate() As String
    getCompileDate = BUILD_DATE
End Function
