VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAdminOfficer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAdminOfficer"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   Close:            C {cmdClose}
' **   Delete:           D {cmdDelete}
' **   Officer:          F {officer on frmAdminOfficer_Sub}
' **   Show Accounts:    H {chkShowAccounts}
' **   Print Report:     P {cmdPrintReport}
' **   Unassigned:       U {chkShowUnassigned}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAdminOfficer_Sub}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, lngPrintReportLbl_Left As Long, blnWasGTR As Boolean
Private lngUnassigned As Long, lngUnassignedLbl_Left As Long, blnFromEnter As Boolean
Private strSortNow1 As String, strSortNow2 As String, strSortNow3 As String
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean
        Dim varTmp00 As Variant

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnWasGTR = gblnGoToReport: blnFromEnter = False
160           strSortNow1 = vbNullString: strSortNow3 = vbNullString: strSortNow3 = vbNullString

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
180           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

190           CLR_DISABLED_FG = CLR_DKGRY
200           CLR_DISABLED_BG = CLR_LTTEAL

210           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left
220           lngUnassignedLbl_Left = .Unassigned_lbl.Left

230           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

240           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

250           varTmp00 = DLookup("[cnt]", "qryAdminOfficer_05", "[adminno] = 1")
260           If IsNull(varTmp00) = False Then
270             lngUnassigned = CLng(varTmp00)
280             If lngUnassigned > 0& Then
290               .Unassigned_lbl.Caption = CStr(lngUnassigned) & " Account" & IIf(lngUnassigned = 1, vbNullString, "s") & " Unassigned"
300               .Unassigned_lbl.Visible = True
310               .chkShowUnassigned.Enabled = True
320             Else
330               .Unassigned_lbl.Visible = False
340               .chkShowUnassigned = False
350               .chkShowUnassigned.Enabled = False
360             End If
370           End If

380           chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
390           chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.
400           lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
410           If lngRecsCur > 0& Then
420             .frmAdminOfficer_Sub.SetFocus
430             .frmAdminOfficer_Sub.Form.FocusHolder.SetFocus
440           End If

450         Else
460           Cancel = -1
470         End If
480       Else
490         Cancel = -1
500       End If
510     End With

520     If Cancel = -1 Then
530       Beep
540       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
550     End If

EXITP:
560     Exit Sub

ERRH:
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Officer:          F {officer on frmAdminOfficer_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAdminOfficer_Sub}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdClose_Click  ' ** Procedure: Below.
790       End Select
800     End If

        ' ** Alt keys.
810     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
820       Select Case intRetVal
          Case vbKeyF
830         With Me
840           intRetVal = 0
850           .frmAdminOfficer_Sub.SetFocus
860           .frmAdminOfficer_Sub.Form.officer.SetFocus
870         End With
880       Case vbKeyX
890         intRetVal = 0
900         cmdClose_Click  ' ** Procedure: Below.
910       End Select
920     End If

        ' ** Ctrl keys.
930     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
940       Select Case intRetVal
          Case vbKeyS
950         With Me
960           intRetVal = 0
970           lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
980           If lngRecsCur > 0& Then
990             .frmAdminOfficer_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAdminOfficer_Sub.
1000          Else
1010            Beep
1020          End If
1030        End With
1040      End Select
1050    End If

        ' ** Ctrl-Shift keys.
1060    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1070      Select Case intRetVal
          Case vbKeyF
1080        With Me
1090          intRetVal = 0
1100          .FocusHolder.SetFocus
1110        End With
1120      End Select
1130    End If

EXITP:
1140    KeyCode = intRetVal
1150    Exit Sub

ERRH:
1160    intRetVal = 0
1170    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
1180    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
1190    Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strDocName As String

1310    With Me
1320      .TimerInterval = 0&
1330      If gblnSetFocus = True Then
1340        gblnSetFocus = False
1350        strDocName = vbNullString
1360  On Error Resume Next
1370        Do Until strDocName = THIS_NAME
1380          DoCmd.SelectObject acForm, THIS_NAME, False
1390          strDocName = Screen.ActiveForm.Name
1400        Loop
1410  On Error GoTo ERRH
1420      End If
1430      lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
1440      If lngRecsCur = 0& Then
1450        .cmdDelete.Enabled = False
1460      End If
1470      If gblnGoToReport = True Then
1480        Select Case blnGoingToReport
            Case True
1490          Select Case blnGoingToReport2
              Case True
1500            DoCmd.Hourglass False  ' ** It should already be off.
1510            gblnGoToReport = False
1520            blnGoingToReport = False
1530            blnGoingToReport2 = False
1540            .GoToReport_arw_pr_img.Visible = False
1550            .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1560            DoEvents
1570          Case False
1580            Select Case garr_varGoToReport(GTR_RNAM)
                Case "rptListOfAdminOfficers_Detailed", "rptListOfAdminOfficers_Unassigned"
1590              blnGoingToReport2 = True
1600              .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1610              .GoToReport_arw_pr_img.Visible = True
1620              .GoToReport_arw_det_img.Visible = False
1630              .Unassigned_lbl.Left = lngUnassignedLbl_Left
1640              .cmdPrintReport.SetFocus
1650              Beep
1660              DoCmd.Hourglass False
1670              DoEvents
1680              .TimerInterval = (GTR_WAIT * 2&)
1690            Case Else
1700              DoCmd.Hourglass False  ' ** It should already be off.
1710              gblnGoToReport = False
1720              blnGoingToReport = False
1730              .GoToReport_arw_pr_img.Visible = False
1740              .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1750              DoEvents
1760            End Select
1770          End Select
1780        Case False
1790          DoCmd.Hourglass True  ' ** Make sure it's still running.
1800          .FocusHolder.SetFocus
1810          DoEvents
1820          blnGoingToReport = True
1830          Select Case garr_varGoToReport(GTR_RNAM)
              Case "rptListOfAdminOfficers"
1840            .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1850            .GoToReport_arw_pr_img.Top = (.cmdPrintReport.Top + (3& * lngTpp))
1860            .GoToReport_arw_pr_img.Left = ((.cmdPrintReport.Left + .cmdPrintReport.Width) - lngTpp)
1870            .GoToReport_arw_pr_img.Visible = True
1880            .cmdPrintReport.SetFocus
1890            If .chkShowAccounts = True Then
1900              .chkShowAccounts = False
1910              chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
1920              DoEvents
1930            End If
1940            If .chkShowUnassigned = True Then
1950              .chkShowUnassigned = False
1960              chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.
1970              DoEvents
1980            End If
1990            Beep
2000            DoCmd.Hourglass False
2010            DoEvents
2020            .TimerInterval = (GTR_WAIT * 2&)
2030          Case "rptListOfAdminOfficers_Detailed"
2040            .GoToReport_arw_det_img.Top = (.chkShowAccounts_lbl.Top - lngTpp)
2050            .GoToReport_arw_det_img.Visible = True
2060            DoEvents
2070            If .chkShowUnassigned = True Then
2080              .chkShowUnassigned = False
2090              chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.
2100            End If
2110            If .chkShowAccounts = False Then
2120              .chkShowAccounts = True
2130              chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
2140            End If
2150            DoEvents
2160            .TimerInterval = GTR_WAIT
2170          Case "rptListOfAdminOfficers_Unassigned"
2180            .Unassigned_lbl.Left = (.Unassigned_lbl.Left + .GoToReport_arw_det_img.Width)
2190            .GoToReport_arw_det_img.Top = (.chkShowUnassigned_lbl.Top - lngTpp)
2200            .GoToReport_arw_det_img.Visible = True
2210            DoEvents
2220            If .chkShowAccounts = True Then
2230              .chkShowAccounts = False
2240              chkShowAccounts_AfterUpdate  ' ** Procedure: Below.
2250            End If
2260            If .chkShowUnassigned = False Then
2270              .chkShowUnassigned = True
2280              chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.
2290            End If
2300            DoEvents
2310            .TimerInterval = GTR_WAIT
2320          End Select
2330        End Select
2340      Else
2350        DoCmd.Hourglass False
2360      End If
2370    End With

EXITP:
2380    Exit Sub

ERRH:
2390    Select Case ERR.Number
        Case Else
2400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2410    End Select
2420    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim blnRetVal As Boolean

2510    Select Case blnWasGTR
        Case True
2520      blnWasGTR = False
2530    Case False
          ' ** Don't save settings imposed by GoToReport!
2540      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
2550    End Select

2560    blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel

2570    gblnSetFocus = True
2580    gstrReturningForm = THIS_NAME
2590    Select Case strCallingForm
        Case "frmMenu_Utility"
2600      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2610    End Select

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2710    With Me
2720      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2730        Select Case blnPrintReport_Focus
            Case True
2740          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2750          .cmdPrintReport_raised_img.Visible = False
2760        Case False
2770          .cmdPrintReport_raised_img.Visible = True
2780          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2790        End Select
2800        .cmdPrintReport_raised_focus_img.Visible = False
2810        .cmdPrintReport_raised_focus_dots_img.Visible = False
2820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2830        .cmdPrintReport_raised_img_dis.Visible = False
2840      End If
2850    End With

EXITP:
2860    Exit Sub

ERRH:
2870    Select Case ERR.Number
        Case Else
2880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2890    End Select
2900    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Public Sub cmdClose_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3110    DoCmd.Close acForm, THIS_NAME

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab
3270        With Me
3280          intRetVal = 0
3290          lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
3300          If lngRecsCur > 0& Then
3310            .frmAdminOfficer_Sub.SetFocus
3320            .frmAdminOfficer_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAdminOfficer_Sub.
3330          Else
3340            If .cmdPrintReport.Enabled = True Then
3350              .cmdPrintReport.SetFocus
3360            ElseIf .cmdAdd.Enabled = True Then
3370              .cmdAdd.SetFocus
3380            ElseIf .cmdDelete.Enabled = True Then
3390              .cmdDelete.SetFocus
3400            Else
                  ' ** Nowhere else to go!
3410            End If
3420          End If
3430        End With
3440      End Select
3450    End If

        ' ** Shift keys.
3460    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3470      Select Case intRetVal
          Case vbKeyTab
3480        With Me
3490          intRetVal = 0
3500          If .cmdDelete.Enabled = True Then
3510            .cmdDelete.SetFocus
3520          ElseIf .cmdAdd.Enabled = True Then
3530            .cmdAdd.SetFocus
3540          Else
3550            lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
3560            If lngRecsCur > 0& Then
3570              .frmAdminOfficer_Sub.SetFocus
3580              .frmAdminOfficer_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAdminOfficer_Sub.
3590            Else
                  ' ** Nowhere else to go!
3600            End If
3610          End If
3620        End With
3630      End Select
3640    End If

EXITP:
3650    KeyCode = intRetVal
3660    Exit Sub

ERRH:
3670    intRetVal = 0
3680    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
3690    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3700    Case Else
3710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3720    End Select
3730    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3810    With Me
3820      blnPrintReport_Focus = True
3830      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3840      .cmdPrintReport_raised_img.Visible = False
3850      .cmdPrintReport_raised_focus_img.Visible = False
3860      .cmdPrintReport_raised_focus_dots_img.Visible = False
3870      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3880      .cmdPrintReport_raised_img_dis.Visible = False
3890    End With

EXITP:
3900    Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

4010    With Me
4020      blnPrintReport_MouseDown = True
4030      .cmdPrintReport_sunken_focus_dots_img.Visible = True
4040      .cmdPrintReport_raised_img.Visible = False
4050      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4060      .cmdPrintReport_raised_focus_img.Visible = False
4070      .cmdPrintReport_raised_focus_dots_img.Visible = False
4080      .cmdPrintReport_raised_img_dis.Visible = False
4090    End With

EXITP:
4100    Exit Sub

ERRH:
4110    Select Case ERR.Number
        Case Else
4120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4130    End Select
4140    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()
' ** This is a bit messy!

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strDocName As String, strLastOfficer As String, strSortNow As String
        Dim lngRecs As Long, lngCnt As Long
        Dim blnContinue As Boolean
        Dim lngX As Long

4210    With Me

4220      blnContinue = True

4230      DoCmd.Hourglass True
4240      DoEvents

4250      Select Case .chkShowUnassigned
          Case True
4260        strDocName = "rptListOfAdminOfficers_Unassigned"
4270      Case False

4280        Select Case .chkShowAccounts
            Case True

4290          strSortNow = .frmAdminOfficer_Sub.Form.SortNow_Get  ' ** Form Function: frmAdminOfficer_Sub.
4300          If strSortNow = "[officer], [accountno]" Then
4310            strDocName = "rptListOfAdminOfficers_Detailed"
4320          Else
4330            strDocName = "rptListOfAdminOfficers_Detailed_Sort"
4340          End If

4350          Set dbs = CurrentDb
4360          With dbs

                ' ** Empty tmpAccount.
4370            Set qdf = .QueryDefs("qryRpt_ListOfAdminOfficers_02")  ' ** Using tmpAccount for convenience.
4380            qdf.Execute

4390            Set rst2 = .OpenRecordset("tmpAccount", dbOpenDynaset, dbAppendOnly)

                ' ** AdminOfficer, linked to qryRpt_ListOfAdminOfficers_01_01
                ' ** (Account, just dateclosed = Null), sorted.
4400            Set qdf = .QueryDefs("qryRpt_ListOfAdminOfficers_01_03")
4410            Set rst1 = qdf.OpenRecordset
4420            With rst1
4430              If .BOF = True And .EOF = True Then
4440                blnContinue = False
4450                DoCmd.Hourglass False
4460                MsgBox "There are no Administrative Officers currently listed.", vbInformation + vbOKOnly, "Nothing To Do"
4470              Else
4480                .MoveLast
4490                lngRecs = .RecordCount
4500                .MoveFirst
4510                lngCnt = 0&
4520                For lngX = 1& To lngRecs
4530                  With rst2
4540                    .AddNew
4550                    ![adminno] = rst1![adminno]
4560                    ![legalname] = rst1![officer]
4570                    ![alphasort] = "0."
4580                    If IsNull(rst1![accountno]) = True Then
4590                      lngCnt = lngCnt + 1&
4600                      ![accountno] = "{no_entry_" & CStr(lngCnt) & "}"
4610                    Else
4620                      ![accountno] = rst1![accountno]
4630                      ![shortname] = rst1![shortname]
4640                    End If
4650                    ![accounttype] = rst1![accounttype]
4660                    .Update
4670                  End With
4680                  If lngX < lngRecs Then .MoveNext
4690                Next
4700              End If
4710              .Close
4720            End With

4730            rst2.Close

4740            If blnContinue = True Then
                  ' ** tmpAccount, just needed fields, sorted.
4750              Set qdf = .QueryDefs("qryRpt_ListOfAdminOfficers_04")
4760              Set rst1 = qdf.OpenRecordset
4770              With rst1
4780                If .BOF = True And .EOF = True Then
4790                  blnContinue = False
4800                  DoCmd.Hourglass False
4810                  MsgBox "There are no Administrative Officers currently listed.", vbInformation + vbOKOnly, "Nothing To Do"
4820                Else
4830                  .MoveLast
4840                  lngRecs = .RecordCount
4850                  .MoveFirst
4860                  lngCnt = 0&
4870                  strLastOfficer = vbNullString
                      ' ** Number an officer's accounts sequentially.
4880                  For lngX = 1& To lngRecs
4890                    .Edit
4900                    If ![admin_officer] <> strLastOfficer Then
4910                      lngCnt = 1&
4920                      strLastOfficer = ![admin_officer]
4930                      ![admin_count] = CStr(lngCnt) & "."
4940                    Else
4950                      lngCnt = lngCnt + 1&
4960                      ![admin_count] = CStr(lngCnt) & "."
4970                    End If
4980                    .Update
4990                    If lngX < lngRecs Then .MoveNext
5000                  Next
5010                End If
5020                .Close
5030              End With  ' ** rst1.
5040            End If  ' ** blnContinue.

5050            .Close
5060          End With  ' ** dbs.

5070        Case False
5080          strDocName = "rptListOfAdminOfficers"
5090        End Select

5100      End Select

5110    End With

5120    If blnContinue = True Then
5130      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
5140        DoCmd.OpenReport strDocName, acViewPreview
5150        DoCmd.Maximize
5160        DoCmd.RunCommand acCmdFitToWindow
5170        Me.Visible = False
5180      Else
            '##GTR_Ref: rptListOfAdminOfficers
            '##GTR_Ref: rptListOfAdminOfficers_Detailed
            '##GTR_Ref: rptListOfAdminOfficers_Unassigned
5190        DoCmd.OpenReport strDocName, acViewNormal
5200      End If
5210    End If

5220    DoCmd.Hourglass False

EXITP:
5230    Set rst1 = Nothing
5240    Set rst2 = Nothing
5250    Set qdf = Nothing
5260    Set dbs = Nothing
5270    Exit Sub

ERRH:
5280    DoCmd.Hourglass False
5290    Me.Visible = True
5300    Select Case ERR.Number
        Case Else
5310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5320    End Select
5330    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5420    intShiftDown = (Shift And acShiftMask) > 0
5430    intAltDown = (Shift And acAltMask) > 0
5440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5460      Select Case intRetVal
          Case vbKeyTab
5470        With Me
5480          intRetVal = 0
5490          .chkShowAccounts.SetFocus
5500        End With
5510      End Select
5520    End If

        ' ** Shift keys.
5530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5540      Select Case intRetVal
          Case vbKeyTab
5550        With Me
5560          intRetVal = 0
5570          lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
5580          If lngRecsCur > 0& Then
5590            .frmAdminOfficer_Sub.SetFocus
5600            .frmAdminOfficer_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAdminOfficer_Sub.
5610          Else
5620            .cmdClose.SetFocus
5630          End If
5640        End With
5650      End Select
5660    End If

EXITP:
5670    KeyCode = intRetVal
5680    Exit Sub

ERRH:
5690    intRetVal = 0
5700    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
5710    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5720    Case Else
5730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5740    End Select
5750    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

5810    With Me
5820      If blnPrintReport_MouseDown = False Then
5830        Select Case blnPrintReport_Focus
            Case True
5840          .cmdPrintReport_raised_focus_dots_img.Visible = True
5850          .cmdPrintReport_raised_focus_img.Visible = False
5860        Case False
5870          .cmdPrintReport_raised_focus_img.Visible = True
5880          .cmdPrintReport_raised_focus_dots_img.Visible = False
5890        End Select
5900        .cmdPrintReport_raised_img.Visible = False
5910        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
5930        .cmdPrintReport_raised_img_dis.Visible = False
5940      End If
5950    End With

EXITP:
5960    Exit Sub

ERRH:
5970    Select Case ERR.Number
        Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

6110    With Me
6120      .cmdPrintReport_raised_focus_dots_img.Visible = True
6130      .cmdPrintReport_raised_img.Visible = False
6140      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6150      .cmdPrintReport_raised_focus_img.Visible = False
6160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6170      .cmdPrintReport_raised_img_dis.Visible = False
6180      blnPrintReport_MouseDown = False
6190    End With

EXITP:
6200    Exit Sub

ERRH:
6210    Select Case ERR.Number
        Case Else
6220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6230    End Select
6240    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

6310    With Me
6320      .cmdPrintReport_raised_img.Visible = True
6330      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6340      .cmdPrintReport_raised_focus_img.Visible = False
6350      .cmdPrintReport_raised_focus_dots_img.Visible = False
6360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6370      .cmdPrintReport_raised_img_dis.Visible = False
6380      blnPrintReport_Focus = False
6390    End With

EXITP:
6400    Exit Sub

ERRH:
6410    Select Case ERR.Number
        Case Else
6420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6430    End Select
6440    Resume EXITP

End Sub

Private Sub Unassigned_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Unassigned_lbl_MouseMove"

6510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6520    Exit Sub

ERRH:
6530    Select Case ERR.Number
        Case Else
6540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6550    End Select
6560    Resume EXITP

End Sub

Public Sub chkShowAccounts_AfterUpdate()
' ** Public for GoToReport.

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAccounts_AfterUpdate"

        Dim frm As Access.Form
        Dim strSortNow As String, blnDesc As Boolean

6610    With Me
6620      Set frm = .frmAdminOfficer_Sub.Form
6630      lngRecsCur = frm.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
6640      Select Case .chkShowAccounts
          Case True
6650        .chkShowAccounts_lbl.FontBold = True
6660        frm.ShowAccounts True  ' ** Form Procedure: frmAdminOfficer_Sub.
6670        DoEvents
6680        .cmdDelete.Enabled = False
6690        strSortNow = frm.SortNow_Get  ' ** Form Function: frmAdminOfficer_Sub.
6700        If strSortNow2 <> vbNullString Then
              ' ** Use the last sort this option used.
6710          If InStr(strSortNow2, " DESC") > 0 Then
6720            blnDesc = True
6730          Else
6740            blnDesc = False
6750          End If
6760          If Left(strSortNow2, 9) = "[adminno]" Then
6770            frm.SortNow "adminno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
6780            If blnDesc = True Then
6790              frm.SortNow "adminno_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
6800            End If
6810          ElseIf Left(strSortNow2, 9) = "[officer]" Then
6820            frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
6830            If blnDesc = True Then
6840              frm.SortNow "officer_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
6850            End If
6860          ElseIf Left(strSortNow2, 11) = "[alphasort]" Then
6870            frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
6880            If blnDesc = True Then
6890              frm.SortNow "accountno_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
6900            End If
6910          ElseIf Left(strSortNow2, 11) = "[shortname]" Then
6920            frm.SortNow "shortname_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
6930            If blnDesc = True Then
6940              frm.SortNow "shortname_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
6950            End If
6960          Else
6970            frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
6980          End If
6990        Else
7000          frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
7010        End If
7020        DoEvents
7030        strSortNow2 = frm.SortNow_Get  ' ** Form Procedure: frmAdminOfficer_Sub.
7040      Case False
7050        .chkShowAccounts_lbl.FontBold = False
7060        frm.ShowAccounts False  ' ** Form Procedure: frmAdminOfficer_Sub.
7070        DoEvents
7080        If lngRecsCur > 0& Then
7090          .cmdDelete.Enabled = True
7100        End If
7110        strSortNow = frm.SortNow_Get  ' ** Form Function: frmAdminOfficer_Sub.
7120        If strSortNow1 <> vbNullString Then
              ' ** Use the last sort this option used.
7130          If InStr(strSortNow1, " DESC") > 0 Then
7140            blnDesc = True
7150          Else
7160            blnDesc = False
7170          End If
7180          If Left(strSortNow1, 9) = "[adminno]" Then
7190            frm.SortNow "adminno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
7200            If blnDesc = True Then
7210              frm.SortNow "adminno_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
7220            End If
7230          ElseIf Left(strSortNow1, 9) = "[officer]" Then
7240            frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
7250            If blnDesc = True Then
7260              frm.SortNow "officer_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
7270            End If
7280          ElseIf Left(strSortNow1, 10) = "[assigned]" Then
7290            frm.SortNow "assigned_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
7300            If blnDesc = True Then
7310              frm.SortNow "assigned_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
7320            End If
7330          Else
7340            frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
7350          End If
7360        Else
7370          frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
7380        End If
7390        DoEvents
7400        strSortNow1 = frm.SortNow_Get  ' ** Form Procedure: frmAdminOfficer_Sub.
7410      End Select
7420      If lngRecsCur > 0& Then
7430        .frmAdminOfficer_Sub.SetFocus
7440      End If
7450    End With

EXITP:
7460    Set frm = Nothing
7470    Exit Sub

ERRH:
7480    Select Case ERR.Number
        Case Else
7490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7500    End Select
7510    Resume EXITP

End Sub

Private Sub chkShowAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7620    intShiftDown = (Shift And acShiftMask) > 0
7630    intAltDown = (Shift And acAltMask) > 0
7640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7670        With Me
7680          intRetVal = 0
7690          If .chkShowUnassigned.Enabled = True Then
7700            .chkShowUnassigned.SetFocus
7710          ElseIf .cmdPrintReport.Enabled = True Then
7720            .cmdPrintReport.SetFocus
7730          ElseIf .cmdAdd.Enabled = True Then
7740            .cmdAdd.SetFocus
7750          ElseIf .cmdDelete.Enabled = True Then
7760            .cmdDelete.SetFocus
7770          Else
7780            .cmdClose.SetFocus
7790          End If
7800        End With
7810      End Select
7820    End If

        ' ** Shift keys.
7830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7850        With Me
7860          intRetVal = 0
7870          lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
7880          If lngRecsCur > 0& Then
7890            .frmAdminOfficer_Sub.SetFocus
7900            .frmAdminOfficer_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAdminOfficer_Sub.
7910          Else
7920            .cmdClose.SetFocus
7930          End If
7940        End With
7950      End Select
7960    End If

EXITP:
7970    KeyCode = intRetVal
7980    Exit Sub

ERRH:
7990    intRetVal = 0
8000    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
8010    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8020    Case Else
8030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8040    End Select
8050    Resume EXITP

End Sub

Private Sub chkShowUnassigned_Enter()
' ** If there are no unassigned accounts,
' ** it shouldn't even be able to get here.

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowUnassigned_Enter"

        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

8110    With Me
8120      blnRetVal = True
8130      blnFromEnter = False
8140      varTmp00 = DCount("*", "qryAdminOfficer_12")
8150      If IsNull(varTmp00) Then
8160        blnRetVal = False
8170      Else
8180        If varTmp00 = 0 Then
8190          blnRetVal = False
8200        Else
              ' ** Yes, there are unassigned accounts.
8210        End If
8220      End If
8230      If blnRetVal = False Then
            ' ** Move on and disable.
8240        If .chkShowUnassigned = True Then
8250          .chkShowUnassigned = False
8260          blnFromEnter = True
8270          chkShowUnassigned_AfterUpdate  ' ** Procedure: Below.
8280          DoEvents
8290        End If
8300        If .chkShowAccounts.Enabled = True Then
8310          .chkShowAccounts.SetFocus
8320        ElseIf .cmdPrintReport.Enabled = True Then
8330          .cmdPrintReport.SetFocus
8340        ElseIf .cmdAdd.Enabled = True Then
8350          .cmdAdd.SetFocus
8360        ElseIf .cmdDelete.Enabled = True Then
8370          .cmdDelete.SetFocus
8380        Else
8390          .cmdClose.SetFocus
8400        End If
8410        .chkShowUnassigned.Enabled = False
8420        .Unassigned_lbl.Visible = False
8430      End If
8440    End With

EXITP:
8450    Exit Sub

ERRH:
8460    Select Case ERR.Number
        Case Else
8470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8480    End Select
8490    Resume EXITP

End Sub

Private Sub chkShowUnassigned_AfterUpdate()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowUnassigned_AfterUpdate"

        Dim frm As Access.Form
        Dim strSortNow As String, blnDesc As Boolean

8510    With Me
8520      Set frm = .frmAdminOfficer_Sub.Form
8530      Select Case .chkShowUnassigned
          Case True
8540        .chkShowUnassigned_lbl.FontBold = True
8550        .chkShowAccounts.Enabled = False
8560        .cmdDelete.Enabled = False
8570        .cmdAdd.Enabled = False
8580        frm.ShowUnassigned True  ' ** Form Procedure: frmAdminOfficer_Sub.
8590        DoEvents
8600        strSortNow = frm.SortNow_Get  ' ** Form Function: frmAdminOfficer_Sub.
8610        If strSortNow3 <> vbNullString Then
              ' ** Use the last sort this option used.
8620          If InStr(strSortNow3, " DESC") > 0 Then
8630            blnDesc = True
8640          Else
8650            blnDesc = False
8660          End If
8670          If Left(strSortNow3, 9) = "[officer]" Then
8680            frm.SortNow "cmbAdminNos_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
8690            If blnDesc = True Then
8700              frm.SortNow "cmbAdminNos_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
8710            End If
8720          ElseIf Left(strSortNow3, 11) = "[alphasort]" Then
8730            frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
8740            If blnDesc = True Then
8750              frm.SortNow "accountno_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
8760            End If
8770          ElseIf Left(strSortNow3, 11) = "[shortname]" Then
8780            frm.SortNow "shortname_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
8790            If blnDesc = True Then
8800              frm.SortNow "shortname_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
8810            End If
8820          Else
8830            frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
8840          End If
8850        Else
8860          frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
8870        End If
8880        DoEvents
8890        strSortNow3 = frm.SortNow_Get  ' ** Form Procedure: frmAdminOfficer_Sub.
8900      Case False
8910        .chkShowUnassigned_lbl.FontBold = False
8920        .chkShowAccounts.Enabled = True
8930        frm.ShowUnassigned False  ' ** Form Procedure: frmAdminOfficer_Sub.
8940        DoEvents
8950        strSortNow = frm.SortNow_Get  ' ** Form Function: frmAdminOfficer_Sub.
8960        Select Case .chkShowAccounts
            Case True
8970          If strSortNow2 <> vbNullString Then
                ' ** Use the last sort this option used.
8980            If InStr(strSortNow2, " DESC") > 0 Then
8990              blnDesc = True
9000            Else
9010              blnDesc = False
9020            End If
9030            If Left(strSortNow2, 9) = "[adminno]" Then
9040              frm.SortNow "adminno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9050              If blnDesc = True Then
9060                frm.SortNow "adminno_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
9070              End If
9080            ElseIf Left(strSortNow2, 9) = "[officer]" Then
9090              frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9100              If blnDesc = True Then
9110                frm.SortNow "officer_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
9120              End If
9130            ElseIf Left(strSortNow2, 11) = "[alphasort]" Then
9140              frm.SortNow "accountno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9150              If blnDesc = True Then
9160                frm.SortNow "accountno_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
9170              End If
9180            ElseIf Left(strSortNow2, 11) = "[shortname]" Then
9190              frm.SortNow "shortname_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9200              If blnDesc = True Then
9210                frm.SortNow "shortname_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
9220              End If
9230            Else
9240              frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9250            End If
9260          Else
9270            frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9280          End If
9290          DoEvents
9300          strSortNow2 = frm.SortNow_Get  ' ** Form Procedure: frmAdminOfficer_Sub.
9310        Case False
9320          If strSortNow1 <> vbNullString Then
                ' ** Use the last sort this option used.
9330            If InStr(strSortNow1, " DESC") > 0 Then
9340              blnDesc = True
9350            Else
9360              blnDesc = False
9370            End If
9380            If Left(strSortNow1, 9) = "[adminno]" Then
9390              frm.SortNow "adminno_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9400              If blnDesc = True Then
9410                frm.SortNow "adminno_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
9420              End If
9430            ElseIf Left(strSortNow1, 9) = "[officer]" Then
9440              frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9450              If blnDesc = True Then
9460                frm.SortNow "officer_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
9470              End If
9480            ElseIf Left(strSortNow1, 10) = "[assigned]" Then
9490              frm.SortNow "assigned_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9500              If blnDesc = True Then
9510                frm.SortNow "assigned_lbl_DblClick", False  ' ** Form Procedure: frmAdminOfficer_Sub.
9520              End If
9530            Else
9540              frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9550            End If
9560          Else
9570            frm.SortNow "officer_lbl_DblClick", True  ' ** Form Procedure: frmAdminOfficer_Sub.
9580          End If
9590          DoEvents
9600          strSortNow1 = frm.SortNow_Get  ' ** Form Procedure: frmAdminOfficer_Sub.
9610        End Select
9620      End Select
9630      If blnFromEnter = False Then
9640        .frmAdminOfficer_Sub.SetFocus
9650      Else
9660        blnFromEnter = False
9670      End If
9680    End With

EXITP:
9690    Set frm = Nothing
9700    Exit Sub

ERRH:
9710    Select Case ERR.Number
        Case Else
9720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9730    End Select
9740    Resume EXITP

End Sub

Private Sub chkShowUnassigned_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowUnassigned_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9820    intShiftDown = (Shift And acShiftMask) > 0
9830    intAltDown = (Shift And acAltMask) > 0
9840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9870        With Me
9880          intRetVal = 0
9890          If .cmdPrintReport.Enabled = True Then
9900            .cmdPrintReport.SetFocus
9910          ElseIf .cmdAdd.Enabled = True Then
9920            .cmdAdd.SetFocus
9930          ElseIf .cmdDelete.Enabled = True Then
9940            .cmdDelete.SetFocus
9950          Else
9960            .cmdClose.SetFocus
9970          End If
9980        End With
9990      End Select
10000   End If

        ' ** Shift keys.
10010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10030       With Me
10040         intRetVal = 0
10050         If .chkShowAccounts.Enabled = True Then
10060           .chkShowAccounts.SetFocus
10070         Else
10080           lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
10090           If lngRecsCur > 0& Then
10100             .frmAdminOfficer_Sub.SetFocus
10110             .frmAdminOfficer_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAdminOfficer_Sub.
10120           Else
10130             .cmdClose.SetFocus
10140           End If
10150         End If
10160       End With
10170     End Select
10180   End If

EXITP:
10190   KeyCode = intRetVal
10200   Exit Sub

ERRH:
10210   intRetVal = 0
10220   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
10230   Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
10240   Case Else
10250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10260   End Select
10270   Resume EXITP

End Sub

Private Sub cmdAdd_Click()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim strDocName As String

10310   With Me
10320     strDocName = "frmAdminOfficer_Add"
10330     DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME
10340     .frmAdminOfficer_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAdminOfficer_Sub.
10350   End With

EXITP:
10360   Exit Sub

ERRH:
10370   Select Case ERR.Number
        Case Else
10380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10390   End Select
10400   Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10520   intShiftDown = (Shift And acShiftMask) > 0
10530   intAltDown = (Shift And acAltMask) > 0
10540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10560     Select Case intRetVal
          Case vbKeyTab
10570       With Me
10580         intRetVal = 0
10590         If .cmdDelete.Enabled = True Then
10600           .cmdDelete.SetFocus
10610         Else
10620           .cmdClose.SetFocus
10630         End If
10640       End With
10650     End Select
10660   End If

        ' ** Shift keys.
10670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10680     Select Case intRetVal
          Case vbKeyTab
10690       With Me
10700         intRetVal = 0
10710         .chkShowAccounts.SetFocus
10720       End With
10730     End Select
10740   End If

EXITP:
10750   KeyCode = intRetVal
10760   Exit Sub

ERRH:
10770   intRetVal = 0
10780   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
10790   Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
10800   Case Else
10810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10820   End Select
10830   Resume EXITP

End Sub

Private Sub cmdDelete_Click()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccounts As String
        Dim lngAdminNo As Long
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnContinue As Boolean

10910   blnContinue = True
10920   strAccounts = vbNullString

10930   With Me
10940     lngRecsCur = .frmAdminOfficer_Sub.Form.RecCnt  ' ** Form Function: frmAdminOfficer_Sub.
10950     If lngRecsCur > 0& Then
10960       If IsNull(.frmAdminOfficer_Sub.Form.officer) = True Then
10970         blnContinue = False
10980         Beep
10990         MsgBox "There is no administrator to delete.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
11000         .frmAdminOfficer_Sub.SetFocus
11010         .frmAdminOfficer_Sub.Form.officer.SetFocus
11020       Else
11030         If Trim(.frmAdminOfficer_Sub.Form.officer) = vbNullString Then
11040           blnContinue = False
11050           Beep
11060           MsgBox "There is no administrator to delete.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
11070           .frmAdminOfficer_Sub.SetFocus
11080           .frmAdminOfficer_Sub.Form.officer.SetFocus
11090         Else

11100           lngAdminNo = .frmAdminOfficer_Sub.Form.adminno

11110           Set dbs = CurrentDb
11120           With dbs
                  ' ** Account table, by specified [adno].
11130             Set qdf = .QueryDefs("qryAdminOfficer_03")
11140             With qdf.Parameters
11150               ![adno] = lngAdminNo
11160             End With
11170             Set rst = qdf.OpenRecordset
11180             If rst.BOF = True And rst.EOF = True Then
                    ' ** None found, so OK to delete.
11190               rst.Close
11200             Else
11210               With rst
11220                 .MoveLast
11230                 lngRecs = .RecordCount
11240                 .MoveFirst
11250                 For lngX = 1& To lngRecs
11260                   strAccounts = strAccounts & ![accountno]
11270                   If lngX < lngRecs Then
                          ' ** There are more to go, so append a comma.
11280                     strAccounts = strAccounts & ", "
11290                     .MoveNext
11300                   End If
11310                 Next
11320                 .Close
11330               End With
11340             End If
11350             .Close
11360           End With

11370           If strAccounts <> vbNullString Then
                  ' ** This administrator owns accounts.
11380             blnContinue = False
11390             Beep
11400             MsgBox "This administrator cannot be deleted." & vbCrLf & vbCrLf & _
                    "Currently assigned to account" & IIf(lngRecs > 1, "s", vbNullString) & ": " & strAccounts & ".", _
                    vbInformation + vbOKOnly, ("Delete Administrator" & Space(40))
11410             .frmAdminOfficer_Sub.SetFocus
11420             .frmAdminOfficer_Sub.Form.officer.SetFocus
11430           End If

11440           If blnContinue = True Then
11450             Beep
11460             If MsgBox("Are you sure you want to delete '" & .frmAdminOfficer_Sub.Form.officer & "' from the list of administrators?", _
                      vbQuestion + vbYesNo, ("Delete Administrator" & Space(40))) = vbYes Then
11470               Set dbs = CurrentDb
11480               With dbs
                      ' ** Delete AdminOfficer, by specified [adno].
11490                 Set qdf = .QueryDefs("qryAdminOfficer_04")
11500                 With qdf.Parameters
11510                   ![adno] = lngAdminNo
11520                 End With
11530                 qdf.Execute
11540                 .Close
11550               End With
11560               .frmAdminOfficer_Sub.Form.Requery
11570             End If
11580           End If

11590         End If
11600       End If
11610     Else
11620       Beep
11630       MsgBox "There are no Administrative Officers to delete.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
11640     End If
11650   End With

EXITP:
11660   Set rst = Nothing
11670   Set qdf = Nothing
11680   Set dbs = Nothing
11690   Exit Sub

ERRH:
11700   Select Case ERR.Number
        Case Else
11710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11720   End Select
11730   Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11820   intShiftDown = (Shift And acShiftMask) > 0
11830   intAltDown = (Shift And acAltMask) > 0
11840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11860     Select Case intRetVal
          Case vbKeyTab
11870       With Me
11880         intRetVal = 0
11890         .cmdClose.SetFocus
11900       End With
11910     End Select
11920   End If

        ' ** Shift keys.
11930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11940     Select Case intRetVal
          Case vbKeyTab
11950       With Me
11960         intRetVal = 0
11970         If .cmdAdd.Enabled = True Then
11980           .cmdAdd.SetFocus
11990         Else
12000           .chkShowAccounts.SetFocus
12010         End If
12020       End With
12030     End Select
12040   End If

EXITP:
12050   KeyCode = intRetVal
12060   Exit Sub

ERRH:
12070   intRetVal = 0
12080   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
12090   Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
12100   Case Else
12110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12120   End Select
12130   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

12210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12220   Exit Sub

ERRH:
12230   Select Case ERR.Number
        Case Else
12240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12250   End Select
12260   Resume EXITP

End Sub

Public Sub SortNow_Set(strSort As String, intMode As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Set"

12310   Select Case intMode
        Case 1
12320     strSortNow1 = strSort
12330   Case 2
12340     strSortNow2 = strSort
12350   Case 3
12360     strSortNow3 = strSort
12370   End Select

EXITP:
12380   Exit Sub

ERRH:
12390   Select Case ERR.Number
        Case Else
12400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12410   End Select
12420   Resume EXITP

End Sub
