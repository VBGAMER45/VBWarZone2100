VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeaponGradesCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WeaponGradesCol
' src/spectatorwidgets.cpp: the fields of class WeaponGradesCol (spectatorwidgets_c); the initialisers are its _construct's.
Option Explicit

Public manager As WIDGET                ' std::shared_ptr<WzWeaponGradesColumnManager>
Public playerIdx As Long                ' uint32_t
