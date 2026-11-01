VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HelpDisplayForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class HelpDisplayForm
' src/screens/helpscreen.cpp: the fields of class HelpDisplayForm (helpscreen_c); the initialisers are its _construct's.
Option Explicit

Public title_ As WIDGET                 ' std::shared_ptr<W_LABEL> title
Public description As WIDGET            ' std::shared_ptr<Paragraph>
Public interactions As WIDGET           ' std::shared_ptr<ScrollableListWidget>
Public relatedKeybindings As WIDGET     ' std::shared_ptr<ScrollableListWidget>
