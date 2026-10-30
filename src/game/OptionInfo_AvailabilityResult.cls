VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionInfo_AvailabilityResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' src/titleui/widgets/optionsform.h: struct OptionInfo::AvailabilityResult (a class, so that an
' OptionAvailabilityCondition WzFunc can return it: OptionInfo_AvailabilityResult_new, optionsform_c)
Option Explicit

Public available_ As Boolean          ' (VB) available: stats_c's AVAILABLE has the name
Public localizedUnavailabilityReason As String
