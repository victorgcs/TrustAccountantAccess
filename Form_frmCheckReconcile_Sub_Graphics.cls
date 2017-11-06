VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckReconcile_Sub_Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckReconcile_Sub_Graphics"

'VGC 12/19/2012: CHANGES!

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmCheckReconcile" Then

160       Else
170         Cancel = -1
180       End If
190     End With

200     If Cancel = -1 Then
210       Beep
220       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
230     End If

EXITP:
240     Exit Sub

ERRH:
250     DoCmd.Hourglass False
260     Select Case ERR.Number
        Case Else
270       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
280     End Select
290     Resume EXITP

End Sub

Private Sub Form_Load()

300   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim frm As Access.Form, ctl As Access.Control

310     With Me
320       Set frm = .Parent
330       For Each ctl In .Controls
340         With ctl
350           If .ControlType = acBoundObjectFrame Then
360             frm.Controls(.Name) = .Value
370           End If
380         End With
390       Next
400     End With

EXITP:
410     Set ctl = Nothing
420     Set frm = Nothing
430     Exit Sub

ERRH:
440     DoCmd.Hourglass False
450     Select Case ERR.Number
        Case Else
460       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
470     End Select
480     Resume EXITP

End Sub
