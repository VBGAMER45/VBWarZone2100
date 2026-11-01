VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzWeaponGradesColumnManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzWeaponGradesColumnManager
' src/spectatorwidgets.cpp: the fields of class WzWeaponGradesColumnManager (spectatorwidgets_c); the initialisers are its _construct's.
Option Explicit

' std::vector<WeaponSubclassInfo> weapSubclassInfo (WSC_NUM_WEAPON_SUBCLASSES): maxNumWeaponGrade and the
' std::unordered_set<uint32_t> playersWithMaxGrade (a Boolean per player)
Private maxNumWeaponGrade_(0 To WSC_NUM_WEAPON_SUBCLASSES - 1) As Long
Private playersWithMaxGrade_(0 To WSC_NUM_WEAPON_SUBCLASSES - 1, 0 To MAX_PLAYERS - 1) As Boolean
' std::vector<WEAPON_SUBCLASS> visibleColumnOrder: visibleColumn(0 .. visibleColumnCount - 1)
Private visibleColumn_(0 To WSC_NUM_WEAPON_SUBCLASSES - 1) As Long
Public visibleColumnCount As Long
Public cachedNumberTexts As Collection  ' std::unordered_map<uint32_t, std::shared_ptr<WzCachedText>>: keys "n<number>"
Public columnIcons As Collection        ' ColumnIconsMap (WEAPON_SUBCLASS -> WzWeapColumnIcon): keys "s<subClass>"
Public onIdealWidthChangeFuncs As Collection ' std::vector<OnChangeFunc>: WzFuncs
Public updateInterval As Long           ' uint32_t = 0
Public lastGameTimeUpdated As Long      ' uint32_t = 0
Public idealHeightForEachRow As Long    ' int = 0
' (optional<uint32_t> lastMouseOverColumn: never read)

Friend Property Get maxNumWeaponGrade(ByVal subClass As Long) As Long
    maxNumWeaponGrade = maxNumWeaponGrade_(subClass)
End Property

Friend Property Let maxNumWeaponGrade(ByVal subClass As Long, ByVal value As Long)
    maxNumWeaponGrade_(subClass) = value
End Property

Friend Property Get hasMaxGrade(ByVal subClass As Long, ByVal player As Long) As Boolean
    hasMaxGrade = playersWithMaxGrade_(subClass, player)
End Property

Friend Property Let hasMaxGrade(ByVal subClass As Long, ByVal player As Long, ByVal value As Boolean)
    playersWithMaxGrade_(subClass, player) = value
End Property

' playersWithMaxGrade.clear()
Friend Sub clearPlayersWithMaxGrade(ByVal subClass As Long)
    Dim p As Long
    For p = 0 To MAX_PLAYERS - 1
        playersWithMaxGrade_(subClass, p) = False
    Next
End Sub

Friend Property Get visibleColumn(ByVal i As Long) As Long
    visibleColumn = visibleColumn_(i)
End Property

Friend Property Let visibleColumn(ByVal i As Long, ByVal value As Long)
    visibleColumn_(i) = value
End Property
