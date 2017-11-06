VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Extra1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'VGC 10/06/2015: CHANGES!

Private Const THIS_NAME As String = "frmMenu_Extra1"
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me

120     End With

EXITP:
130     Exit Sub

ERRH:
140     Select Case ERR.Number
        Case Else
150       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
160     End Select
170     Resume EXITP

End Sub
