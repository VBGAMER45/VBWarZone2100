VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetQueuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' lib/netplay/netqueue.h: class NetQueuePair, used for talking to a socket.
' We insert NetMessages in the send NetQueue, which converts the messages
' into a stream of bytes for the socket. We take incoming bytes from the
' socket and insert them in the receive NetQueue, which converts the bytes
' back into NetMessages.
Option Explicit

Public send As NetQueue
Public receive As NetQueue

Private Sub Class_Initialize()
    Set send = New NetQueue
    Set receive = New NetQueue
    send.setWillNeverGetMessages
    receive.setWillNeverGetMessagesForNet
End Sub
