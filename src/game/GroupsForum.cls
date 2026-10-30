VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroupsForum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class GroupsForum
' src/hci/groups.h, groups.cpp: GroupsForum's fields (groups_c)
Option Explicit

Public groupsList As WIDGET                     ' std::shared_ptr<IntListTabWidget>
Public groupsUIController As GroupsUIController ' std::shared_ptr<GroupsUIController>
Public help As WidgetHelp                       ' optional<WidgetHelp> (Nothing = nullopt)
