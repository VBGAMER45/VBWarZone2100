VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlignmentWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class AlignmentWidget
' lib/widget/alignment.h: AlignmentWidget's field (alignment_c). Alignment alignment is alignment_.
Option Explicit

Private m_alignment_ As Alignment

Friend Property Get alignment_() As Alignment
    alignment_ = m_alignment_
End Property

Friend Property Let alignment_(v As Alignment)
    m_alignment_ = v
End Property
