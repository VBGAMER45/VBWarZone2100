VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gfx_api_pipeline_state_object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' gfx_api.h: gfx_api::pipeline_state_object, what a context's build_pipeline
' returns. The backends' subclasses hold no more than an index, so this one
' class stands for all of them: gl_pipeline_id (gfx_api_gl.h: psoID, the index
' of the gl_pipeline_state_object in gl_context's list) and null_context's
' (psoID unused). C++'s delete is the last reference going.
Option Explicit

Public broken As Boolean                ' = false
Public psoID As Long                    ' gl_pipeline_id::psoID (= 0)
