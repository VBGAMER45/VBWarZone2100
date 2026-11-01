VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzTeamStratImagesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/hci/teamstrategy.cpp: class WzTeamStrategyColumnImagesManager (teamstrategy_c). Its textures go with it
' (~WzTeamStrategyColumnImagesManager's deletes). getWeaponSubclassImageRef / getUnitTypeImageRef are teamstrategy_c's
' WzTeamStrategyColumnImageRef_weapon / _unitType.
Option Explicit

Private loadedWeaponSubclassImages(0 To WSC_NUM_WEAPON_SUBCLASSES - 1) As gfx_api_texture   ' std::vector<gfx_api::texture*>(WSC_NUM_WEAPON_SUBCLASSES, nullptr)
Private loadedUnitTypeImages(0 To WzStrategyPlanningUnitTypes_NUM_TYPES - 1) As gfx_api_texture ' std::vector<gfx_api::texture*>(WzStrategyPlanningUnitTypes_NUM_TYPES, nullptr)
Private neverSymbol As WzText

Private Sub Class_Initialize()
    Set neverSymbol = New WzText
End Sub

Public Function getWeaponSubclassImageTexture(ByVal subClass As WEAPON_SUBCLASS) As gfx_api_texture
    Dim pLoadedImage As gfx_api_texture
    If loadedWeaponSubclassImages(subClass) Is Nothing Then
        ' load image for the first time
        Set pLoadedImage = loadImageForWeapSubclass(subClass)
        Set loadedWeaponSubclassImages(subClass) = pLoadedImage
    End If
    Set getWeaponSubclassImageTexture = loadedWeaponSubclassImages(subClass)
End Function

Public Function getUnitTypeImageTexture(ByVal unitType As WzStrategyPlanningUnitTypes) As gfx_api_texture
    Dim pLoadedImage As gfx_api_texture
    If loadedUnitTypeImages(unitType) Is Nothing Then
        ' load image for the first time
        Set pLoadedImage = loadImageForUnitType(unitType)
        Set loadedUnitTypeImages(unitType) = pLoadedImage
    End If
    Set getUnitTypeImageTexture = loadedUnitTypeImages(unitType)
End Function

Public Sub renderNeverSymbol(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    Dim x0 As Long, y0 As Long
    neverSymbol.setText "x", font_regular_bold
    x0 = x + ((width - neverSymbol.width()) \ 2)
    y0 = y + ((height - neverSymbol.lineSize()) \ 2)
    neverSymbol.render x0, y0 - CSng(neverSymbol.aboveBase()), psPalette(WZCOL_RED)
End Sub
