VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Account_Sub_Nav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Account_Sub_Nav"

'VGC 10/11/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add Account:      A {cmdAdd on frmMenu_Account}
' **   Close:            C {cmdClose on frmMenu_Account}
' **   Delete Account:   D {cmdDelete on frmMenu_Account}
' **   Search:           E {cmdSearch on frmMenu_Account}
' **   Hide:             H {cmdHideTransactions on frmMenu_Account}
' **   Inc/Exp Codes:    I {cmdEditIECodes on frmMenu_Account}
' **   Display List:     L {opgSubForm_optList on frmMenu_Account}
' **   Comments:         M {cmdEditComments on frmMenu_Account}
' **   Short Name:       N {shortname on frmMenu_Account_Sub_One, frmMenu_Account_Sub_List}
' **   Display One:      O {opgSubForm_optOne on frmMenu_Account}
' **   Profile:          P {cmdAccountProfile on frmMenu_Account}
' **   Transactions:     R {cmdTransactions on frmMenu_Account}
' **   Assets:           S {cmdAssets on frmMenu_Account}
' **   Tax Codes:        T {cmdEditTaxCodes on frmMenu_Account}
' **   Include Closed:   U {chkIncludeClosed on frmMenu_Account}
' **   Exit:             X {cmdClose on frmMenu_Account}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next Record:      N {MoveRec}
' **   Previous Record:  P {MoveRec}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Unlock Var:       C {blnFormOpening: Form_KeyDown on frmMenu_Account}
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm on frmMenu_Account}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmMenu_Account" Then

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
250     Select Case ERR.Number
        Case Else
260       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
270     End Select
280     Resume EXITP

End Sub

Private Sub Form_Current()

300   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim strTmp01 As String

310     With Me
320   On Error Resume Next
330       strTmp01 = .Parent.frmMenu_Account_Sub_One.Form.Name
340   On Error GoTo ERRH
350       If strTmp01 = "frmMenu_Account_Sub_One" Then
360         .Parent.frmMenu_Account_Sub_One.Form.MoveRec 0, .accountno  ' ** Form Procedure: frmMenu_Account_Sub_One.
370       End If
380   On Error Resume Next
390       strTmp01 = .Parent.frmMenu_Account_Sub_List.Form.Name
400   On Error GoTo ERRH
410       If strTmp01 = "frmMenu_Account_Sub_List" Then
420         .Parent.frmMenu_Account_Sub_List.Form.MoveRec 0, .accountno  ' ** Form Procedure: frmMenu_Account_Sub_List.
430       End If
440     End With

EXITP:
450     Exit Sub

ERRH:
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim Cancel As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Short Name:       N {shortname on frmMenu_Account_Sub_One, frmMenu_Account_Sub_List}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next Record:      N {MoveRec}
        ' **   Previous Record:  P {MoveRec}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Unlock Var:       C {blnFormOpening: Form_KeyDown on frmMenu_Account}
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm on frmMenu_Account}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         With Me
580           intRetVal = 0
590           .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
600         End With
610       End Select
620     End If

        ' ** Alt keys.
630     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
640       Select Case intRetVal
          Case vbKeyC
650         With Me
660           intRetVal = 0
670           .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
680         End With
690       Case vbKeyN
700         With Me
710           intRetVal = 0
720           DoCmd.SelectObject acForm, .Parent.Name, False
730           Select Case .Parent.opgSubForm
              Case .Parent.opgSubForm_optOne.OptionValue
740             .Parent.frmMenu_Account_Sub_One.SetFocus
750             .Parent.frmMenu_Account_Sub_One.Form.shortname.SetFocus
760           Case .Parent.opgSubForm_optList.OptionValue
770             .Parent.frmMenu_Account_Sub_List.SetFocus
780             .Parent.frmMenu_Account_Sub_List.Form.shortname.SetFocus
790           End Select
800         End With
810       Case vbKeyX
820         With Me
830           intRetVal = 0
840           .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
850         End With
860       End Select
870     End If

        ' ** Ctrl keys.
880     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
890       Select Case intRetVal
          Case vbKeyN
900         With Me
910           intRetVal = 0
920           .MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
930         End With
940       Case vbKeyP
950         With Me
960           intRetVal = 0
970           .MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
980         End With
990       End Select
1000    End If

        ' ** Ctrl-Shift keys.
1010    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1020      Select Case intRetVal
          Case vbKeyC
1030        With Me
              ' ** In case this gets stuck on.
1040          .Parent.FormCloseVar False  ' ** blnFormOpening = False
1050          DoBeeps 2, 300  ' ** Module Function: modWindowFunctions.
1060        End With
1070      Case vbKeyF
1080        With Me
1090          intRetVal = 0
1100          DoCmd.SelectObject acForm, .Parent.Name, False
1110          .Parent.FocusHolder.SetFocus
1120        End With
1130      Case vbKeyT
1140        With Me
1150          intRetVal = 0  ' ** User doesn't matter.
1160          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
1170            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1180          Else
1190            If CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Module Function: modFileUtilities.
1200              With .Parent
1210                Select Case .chkLoadTimer_vis
                    Case True
1220                  .chkLoadTimer.Visible = False
1230                  .chkLoadTimer_lbl1.Visible = False
1240                  .cmdDelete.Caption = "&Delete Account"
1250                  .cmdDelete.StatusBarText = "Delete the current account - Alt+D"
1260                  .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Alt+D"
1270                Case False
1280                  .chkLoadTimer.Visible = True
1290                  .chkLoadTimer_lbl1.Visible = True
1300                  .cmdDelete.Caption = "Delete Account"
1310                  .cmdDelete.StatusBarText = "Delete the current account - Ctrl+D"
1320                  .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Ctrl+D"
1330                End Select
1340              End With
1350            End If
1360          End If
1370          .Parent.chkLoadTimer_vis = .Parent.chkLoadTimer.Visible
1380        End With
1390      Case vbKeyX
1400        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1410          With Me
1420            intRetVal = 0
1430            .Parent.JustClose_Set True  ' ** Form Procedure: frmMenu_Account.
1440            SetOption_Dev  ' ** Module Function: modUtilities.
1450            DoCmd.Close acForm, strCallingForm
1460          End With
1470        End If
1480      End Select
1490    End If

EXITP:
1500    KeyCode = intRetVal
1510    Exit Sub

ERRH:
1520    intRetVal = 0
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

1610    Select Case intWhere
        Case acCmdRecordsGoToFirst
1620      DoCmd.RunCommand acCmdRecordsGoToFirst
1630    Case acCmdRecordsGoToLast
1640      DoCmd.RunCommand acCmdRecordsGoToLast
1650    Case acCmdRecordsGoToPrevious
1660      DoCmd.RunCommand acCmdRecordsGoToPrevious
1670    Case acCmdRecordsGoToNext
1680      DoCmd.RunCommand acCmdRecordsGoToNext
1690    Case Else
1700      If IsMissing(varID) = False Then
1710        If IsEmpty(varID) = False Then
1720          If IsNull(varID) = False Then
1730            If Trim(varID) <> vbNullString Then
1740              With Me
1750                Set rst = .RecordsetClone
1760                With rst
1770                  If .BOF = True And .EOF = True Then
                        ' ** Skip it!
1780                  Else
1790                    .MoveFirst
1800                    .FindFirst "[accountno] = '" & CStr(varID) & "'"
1810                    If .NoMatch = False Then
1820                      Me.Bookmark = .Bookmark
1830                    End If
1840                  End If
1850                  .Close
1860                End With
1870              End With
1880            End If
1890          End If
1900        End If
1910      End If
1920    End Select

EXITP:
1930    Set rst = Nothing
1940    Exit Sub

ERRH:
1950    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
1960    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
1970    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
1980    Case Else
1990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000    End Select
2010    Resume EXITP

End Sub
