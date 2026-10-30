VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TitleFormWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class TitleFormWrapper
' src/titleui/widgets/titleformwrapper.cpp: TitleFormWrapper's fields (titleformwrapper_c); the
' initialisers are TitleFormWrapper_construct's.
Option Explicit

Public getTitleStringFunc As WzFunc     ' GetTitleStringFunc: std::function<WzString(void)>
Public sideText As WIDGET               ' std::shared_ptr<TitleFormSideText>
Public formWrapper As WIDGET            ' std::shared_ptr<IntFormAnimatedWrapper>
Public outerPaddingX As Long
Public titleFormSpacing As Long
Public outerPaddingY As Long
