VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Post_Clear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Post_Clear"

'VGC 12/31/2011: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Both:             B {cmdDeleteBoth}
' **   Journal:          J {cmdDeleteJournal}
' **   No:               N {cmdNo}
' **   Unommitted:       U {cmdDeleteUncommitted}

Private blnAdmin As Boolean, intMode As Integer
Private strCallingForm As String, strThisJournalUser As String
Private lngCnt1 As Long, lngCnt2 As Long, lngUsers1 As Long, lngUsers2 As Long
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
150           If CharCnt(strTmp01, "~") = 7& Then  ' ** Module Function: modStringFuncs.

                ' ** THIS_NAME & "~" & CStr(intMode) & "~" & IIf(gblnAdmin = True, "True", "False") & "~" & gstrJournalUser & "~" & _
                ' **   CStr(lngCnt1) & "~" & CStr(lngCnt2) & "~" & CStr(lngUsers1) & "~" & CStr(lngUsers2)
160             intPos01 = InStr(strTmp01, "~")
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strTmp01 = Mid(strTmp01, (intPos01 + 1))
190             intPos01 = InStr(strTmp01, "~")
200             intMode = CInt(Left(strTmp01, (intPos01 - 1)))
210             strTmp01 = Mid(strTmp01, (intPos01 + 1))
220             intPos01 = InStr(strTmp01, "~")
230             blnAdmin = CBool(Left(strTmp01, (intPos01 - 1)))
240             strTmp01 = Mid(strTmp01, (intPos01 + 1))
250             intPos01 = InStr(strTmp01, "~")
260             strThisJournalUser = Left(strTmp01, (intPos01 - 1))
270             strTmp01 = Mid(strTmp01, (intPos01 + 1))
280             intPos01 = InStr(strTmp01, "~")
290             lngCnt1 = CLng(Left(strTmp01, (intPos01 - 1)))
300             strTmp01 = Mid(strTmp01, (intPos01 + 1))
310             intPos01 = InStr(strTmp01, "~")
320             lngCnt2 = CLng(Left(strTmp01, (intPos01 - 1)))
330             strTmp01 = Mid(strTmp01, (intPos01 + 1))
340             intPos01 = InStr(strTmp01, "~")
350             lngUsers1 = CLng(Left(strTmp01, (intPos01 - 1)))
360             lngUsers2 = CLng(Mid(strTmp01, (intPos01 + 1)))

370             .PictureData = .Form_Menu_Post_Clear_img

380             .Trans_lbl.Visible = True
390             .TransCnt_lbl.Visible = True

400             .Uncommit_lbl.Visible = False
410             .UncommitCnt_lbl.Visible = False

420             .cmdNo.Visible = True
430             .cmdDeleteJournal.Visible = False
440             .cmdDeleteBoth.Visible = False
450             .cmdDeleteUncommitted.Visible = False

                ' ** gblnAdmin:
                ' **   True
                ' **     intMode = 1  : (lngCnt1 > 0& And lngUsers1 = 1&) And (lngCnt2 = 0& And lngUsers2 = 0&) J only   1 form
                ' **     intMode = 2  : (lngCnt1 > 0& And lngUsers1 > 1&) And (lngCnt2 = 0& And lngUsers2 = 0&) mJ only  1 form
                ' **     intMode = 3  : (lngCnt1 = 0& And lngUsers1 = 0&) And (lngCnt2 > 0& And lngUsers2 = 1&) U only   1 form
                ' **     intMode = 4  : (lngCnt1 = 0& And lngUsers1 = 0&) And (lngCnt2 > 0& And lngUsers2 > 1&) mU only  1 form
                ' **     intMode = 5  : (lngCnt1 > 0& And lngUsers1 = 1&) And (lngCnt2 > 0& And lngUsers2 = 1&) J & U    1 form
                ' **     intMode = 6  : (lngCnt1 > 0& And lngUsers1 > 1&) And (lngCnt2 > 0& And lngUsers2 = 1&) mJ & U   2 forms
                ' **     intMode = 7  : (lngCnt1 > 0& And lngUsers1 = 1&) And (lngCnt2 > 0& And lngUsers2 > 1&) J & mU   2 forms
                ' **     intMode = 8  : (lngCnt1 > 0& And lngUsers1 > 1&) And (lngCnt2 > 0& And lngUsers2 > 1&) mJ & mU  2 forms
                ' **   False
                ' **     intMode = 9  : (lngCnt1 > 0& And lngCnt2 = 0&) J only
                ' **     intMode = 10 : (lngCnt1 = 0& And lngCnt2 > 0&) U only
                ' **     intMode = 11 : (lngCnt1 > 0& And lngCnt2 > 0&) J & U

460             Select Case intMode
                Case 1, 9
                  ' ** intMode = 1  : (lngCnt1 > 0& And lngUsers1 = 1&) And (lngCnt2 = 0& And lngUsers2 = 0&) J only   1 form
                  ' **   blnAdmin = True
                  ' ** intMode = 9  : (lngCnt1 > 0& And lngCnt2 = 0&) J only
                  ' **   blnAdmin = False
470               .Caption = "Clear Posting Journal"
480               .Trans_lbl.Caption = "Are you sure you want to clear the posting journal?"
490               .TransCnt_lbl.Caption = "Entries for User " & strThisJournalUser & ":  " & CStr(lngCnt1)
500               .cmdDeleteJournal.Caption = "&Yes"
510               .cmdDeleteJournal.Visible = True
520             Case 3, 10
                  ' ** intMode = 3  : (lngCnt1 = 0& And lngUsers1 = 0&) And (lngCnt2 > 0& And lngUsers2 = 1&) U only   1 form
                  ' **   blnAdmin = True
                  ' ** intMode = 10 : (lngCnt1 = 0& And lngCnt2 > 0&) U only
                  ' **   blnAdmin = False
530               .Caption = "Clear Uncommitted Entries"
540               .Trans_lbl.Caption = "There are no committed Journal entries."
550               .TransCnt_lbl.Visible = False
560               .Uncommit_lbl.Caption = "Are you sure you want to clear the uncommitted entries?"
570               .Uncommit_lbl.Visible = True
580               .UncommitCnt_lbl.Caption = "Uncommitted entries for User " & strThisJournalUser & ":  " & CStr(lngCnt2)
590               .UncommitCnt_lbl.Visible = True
600               .cmdDeleteUncommitted.Caption = "&Yes"
610               .cmdDeleteUncommitted.Width = .cmdDeleteJournal.Width
620               .cmdDeleteUncommitted.Left = .cmdDeleteJournal.Left
630               .cmdDeleteUncommitted.Visible = True
640             Case 5, 11
                  ' ** intMode = 5  : (lngCnt1 > 0& And lngUsers1 = 1&) And (lngCnt2 > 0& And lngUsers2 = 1&) J & U    1 form
                  ' **   blnAdmin = True
                  ' ** intMode = 11 : (lngCnt1 > 0& And lngCnt2 > 0&) J & U
                  ' **   blnAdmine = False
650               .Caption = "Clear Posting Journal And Uncommitted Entries"
660               .Trans_lbl.Caption = "Are you sure you want to clear the posting journal?"
670               .TransCnt_lbl.Caption = "Entries for User " & strThisJournalUser & ":  " & CStr(lngCnt1)
680               .Uncommit_lbl.Caption = "Do you also want to clear the uncommitted entries?"
690               .Uncommit_lbl.Visible = True
700               .UncommitCnt_lbl.Caption = "Uncommitted entries for User " & strThisJournalUser & ":  " & CStr(lngCnt2)
710               .UncommitCnt_lbl.Visible = True
720               .cmdDeleteJournal.Visible = True
730               .cmdDeleteBoth.Visible = True
740               .cmdDeleteUncommitted.Visible = True
750             Case 6
                  ' ** intMode = 6  : (lngCnt1 > 0& And lngUsers1 > 1&) And (lngCnt2 > 0& And lngUsers2 = 1&) mJ & U   2 forms
                  ' **   blnAdmin = True
760               .Caption = "Clear Uncommitted Entries"
770               .Trans_lbl.Visible = False
780               .TransCnt_lbl.Visible = False
790               .Uncommit_lbl.Caption = "Do you also want to clear the uncommitted entries?"
800               .Uncommit_lbl.Visible = True
810               .UncommitCnt_lbl.Caption = "Uncommitted entries for User " & strThisJournalUser & ":  " & CStr(lngCnt2)
820               .UncommitCnt_lbl.Visible = True
830               .cmdDeleteUncommitted.Caption = "&Yes"
840               .cmdDeleteUncommitted.Width = .cmdDeleteJournal.Width
850               .cmdDeleteUncommitted.Left = .cmdDeleteJournal.Left
860               .cmdDeleteUncommitted.Visible = True
870             Case 7
                  ' ** intMode = 7  : (lngCnt1 > 0& And lngUsers1 = 1&) And (lngCnt2 > 0& And lngUsers2 > 1&) J & mU   2 forms
                  ' **   blnAdmin = True
880               .Caption = "Clear Posting Journal"
890               .Trans_lbl.Caption = "Are you sure you want to clear the posting journal?"
900               .TransCnt_lbl.Caption = "Entries for User " & strThisJournalUser & ":  " & CStr(lngCnt1)
910               .cmdDeleteJournal.Caption = "&Yes"
920               .cmdDeleteJournal.Visible = True
930             End Select

940             DoCmd.Hourglass False

950           Else
960             Cancel = -1
970           End If
980         Else
990           Cancel = -1
1000        End If
1010      Else
1020        Cancel = -1
1030      End If
1040    End With

1050    If Cancel = -1 Then
1060      Beep
1070      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1080    End If

EXITP:
1090    Exit Sub

ERRH:
1100    Select Case ERR.Number
        Case Else
1110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
1120    End Select
1130    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyEscape
1270        intRetVal = 0
1280        cmdNo_Click  ' ** Procedure: Below.
1290      End Select
1300    End If

EXITP:
1310    KeyCode = intRetVal
1320    Exit Sub

ERRH:
1330    intRetVal = 0
1340    Select Case ERR.Number
        Case Else
1350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1360    End Select
1370    Resume EXITP

End Sub

Private Sub cmdDeleteBoth_Click()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDeleteBoth_Click"

1410    gdblCrtRpt_CostTot = 3#  ' ** Borrowing this variable from the Court Reports.
1420    DoCmd.Close acForm, THIS_NAME

EXITP:
1430    Exit Sub

ERRH:
1440    Select Case ERR.Number
        Case Else
1450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1460    End Select
1470    Resume EXITP

End Sub

Private Sub cmdDeleteJournal_Click()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDeleteJournal_Click"

1510    gdblCrtRpt_CostTot = 1#  ' ** Borrowing this variable from the Court Reports.
1520    DoCmd.Close acForm, THIS_NAME

EXITP:
1530    Exit Sub

ERRH:
1540    Select Case ERR.Number
        Case Else
1550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1560    End Select
1570    Resume EXITP

End Sub

Private Sub cmdDeleteUncommitted_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDeleteUncommitted_Click"

1610    gdblCrtRpt_CostTot = 2#  ' ** Borrowing this variable from the Court Reports.
1620    DoCmd.Close acForm, THIS_NAME

EXITP:
1630    Exit Sub

ERRH:
1640    Select Case ERR.Number
        Case Else
1650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1660    End Select
1670    Resume EXITP

End Sub

Private Sub cmdNo_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdNo_Click"

1710    gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
1720    DoCmd.Close acForm, THIS_NAME

EXITP:
1730    Exit Sub

ERRH:
1740    Select Case ERR.Number
        Case Else
1750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1760    End Select
1770    Resume EXITP

End Sub
