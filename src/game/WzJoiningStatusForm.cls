VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WzJoiningStatusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class WzJoiningStatusForm
' src/screens/joiningscreen.cpp: the fields of class WzJoiningStatusForm (joiningscreen_c)
Option Explicit

Public startTime As Double              ' std::chrono::steady_clock::time_point (widgSteadyNowMs)
Public titleLabel As WIDGET             ' std::shared_ptr<W_LABEL>
Public progressIndicator As WIDGET      ' std::shared_ptr<WzJoiningIndeterminateIndicatorWidget>
Public statusTitle As WIDGET            ' std::shared_ptr<W_LABEL>
Public statusDetails As WIDGET          ' std::shared_ptr<W_LABEL>
Public scrollableParagraphContainer As WIDGET ' std::shared_ptr<ScrollableListWidget>
Public detailsParagraph As WIDGET       ' std::shared_ptr<Paragraph>
Public resolvableJoinErrorWidget As WIDGET ' std::shared_ptr<WIDGET>
Public actionButton As WIDGET           ' std::shared_ptr<WzJoiningActionButton>

Public weakSerial As Long               ' (VB) its std::weak_ptr<WzJoiningStatusForm>s: the widgweak_vb serial
