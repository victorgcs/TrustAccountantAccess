VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmPleaseWait"

'VGC 10/27/2017: CHANGES!

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strTmp01 = Trim(.OpenArgs)
150           intPos01 = InStr(strTmp01, "~")
160           If intPos01 > 0 Then
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strTmp01 = Mid(strTmp01, (intPos01 + 1))
190           Else
200             strCallingForm = strTmp01
210             strTmp01 = vbNullString
220           End If

230           .PictureData = .Form_PleaseWait_img

240           .WaitMsg_lbl.Caption = strTmp01
250           DoEvents

260           If gblnDev_Debug = True Then
270             .TimerInterval = 5000&
280           End If

290         Else
300           Cancel = -1
310         End If
320       Else
330         Cancel = -1
340       End If
350     End With

360     If Cancel = -1 Then
370       Beep
380       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
390     End If

EXITP:
400     Exit Sub

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Form_Timer()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

510     With Me
520       .TimerInterval = 0&
530       If gblnGoToReport = False Then
            'Version_Convert_RunAllStart glngInstance  ' ** Module Function: modVersionDocFuncs.
540       End If
550     End With

EXITP:
560     Exit Sub

ERRH:
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub
