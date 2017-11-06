VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Columns_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Columns_Add"

'VGC 12/06/2011: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   New Entry:        E {chkNewEntry}
' **   No:               N {cmdNo}
' **   Yes:              Y {cmdYes}

' ** gblnCrtRpt_Zero = Commit  Yes/No
' ** gblnMessage     = New Rec Yes/No

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)
150           Beep

160           .PictureData = .Form_Journal_Columns_Add_img

170           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

180           chkNewEntry_AfterUpdate  ' ** Procedure: Below.

190         Else
200           Cancel = -1
210         End If
220       Else
230         Cancel = -1
240       End If
250     End With

260     If Cancel = -1 Then
270       Beep
280       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
290     End If

EXITP:
300     Exit Sub

ERRH:
310     Select Case ERR.Number
        Case Else
320       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
330     End Select
340     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        Const THIS_PROC As String = "Form_KeyDown"

        ' ** Use bit masks to determine which key was pressed.
410     intShiftDown = (Shift And acShiftMask) > 0
420     intAltDown = (Shift And acAltMask) > 0
430     intCtrlDown = (Shift And acCtrlMask) > 0

440     intRetVal = KeyCode

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyEscape
470         With Me
480           intRetVal = 0
490           cmdNo_Click  ' ** Procedure: Below.
500         End With
510       End Select
520     End If

EXITP:
530     KeyCode = intRetVal
540     Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdYes_Click"

610     Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
620     Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub cmdYes_Click()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdYes_Click"

710     With Me
720       gblnCrtRpt_Zero = True  ' ** Borrowing this variable from Court Reports.
730       gblnMessage = .chkNewEntry
740       DoCmd.Close acForm, THIS_NAME
750     End With

EXITP:
760     Exit Sub

ERRH:
770     Select Case ERR.Number
        Case Else
780       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
790     End Select
800     Resume EXITP

End Sub

Private Sub cmdNo_Click()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdNo_Click"

910     With Me
920       gblnCrtRpt_Zero = False  ' ** Borrowing this variable from Court Reports.
930       gblnMessage = .chkNewEntry
940       DoCmd.Close acForm, THIS_NAME
950     End With

EXITP:
960     Exit Sub

ERRH:
970     Select Case ERR.Number
        Case Else
980       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
990     End Select
1000    Resume EXITP

End Sub

Private Sub chkNewEntry_AfterUpdate()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkNewEntry_AfterUpdate"

1110    With Me
1120      gblnMessage = .chkNewEntry
1130      Select Case .chkNewEntry
          Case True
1140        .chkNewEntry_lbl.FontBold = True
1150      Case False
1160        .chkNewEntry_lbl.FontBold = False
1170      End Select
1180    End With

EXITP:
1190    Exit Sub

ERRH:
1200    Select Case ERR.Number
        Case Else
1210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1220    End Select
1230    Resume EXITP

End Sub
