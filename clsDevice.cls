VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "clsDevice"

'VGC 03/31/2011: CHANGES!

' ** From Access 2000 Developer's Handbook, Volume I
' ** by Getz, Litwin, and Gilbert (Sybex)
' ** Copyright 1999.  All rights reserved.

' ** Device class.
' ** Represents a single printer device.

' ** Requires:
' **   No other modules

Public DriverName As String
Public DeviceName As String
Public Port As String
Public Default As Boolean
' **

Public Property Get PrintInfo() As String

100   On Error GoTo ERR

        Const THIS_PROC As String = "PrintInfo"

110     PrintInfo = DeviceName & " on " & Port

EX:
120     Exit Property

ERR:
130     Select Case ERR.Number
        Case Else
140       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
150     End Select
160     Resume EX

End Property
