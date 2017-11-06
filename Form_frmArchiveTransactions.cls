VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmArchiveTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmArchiveTransactions"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Archive Report:   R {cmdArchiveReports}
' **   Archive Through:  T {DateEnd}
' **   Archive:          V {cmdArchive}
' **   Exit:             X {cmdClose}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnJustClose As Boolean, blnNoMsg As Boolean
Private lngArchiveNum As Long, datArchiveDate As Date, datMaxArchDate As Date
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           lngArchiveNum = 0&: datArchiveDate = 0
160           blnJustClose = False: blnNoMsg = False
170           blnGTR_Emblem = False: lngGTR_Stat = 0&

180           CLR_DISABLED_FG = CLR_DKGRY
190           CLR_DISABLED_BG = CLR_LTTEAL

200           datMaxArchDate = CDate("01/01/" + CStr(year(Date) - 1)) - 1

              ' ** Disable the archive button until they put in a date.
210           .cmdArchive.Enabled = False
220           .cmdArchiveReports.Enabled = False

230           If gblnGoToReport = True And .GoToReport_arw_rpt_img.Visible = False Then
240             .GoToReport_arw_rpt_img.Visible = True
250             DoEvents
260           End If

270           RecsTot  ' ** Procedure: Below.

280         Else
290           Cancel = -1
300         End If
310       Else
320         Cancel = -1
330       End If
340     End With

350     If Cancel = -1 Then
360       Beep
370       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
380     End If

EXITP:
390     Exit Sub

ERRH:
400     Select Case ERR.Number
        Case Else
410       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
420     End Select
430     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim strMon As String, strDay As String, strYear As String
        Dim datEndDate As Date
        Dim Cancel As Integer
        Dim varTmp00 As Variant

510     With Me

520       If gblnGoToReport = True And .GoToReport_arw_rpt_img.Visible = False Then
530         .GoToReport_arw_rpt_img.Visible = True
540         DoEvents
550       End If

560       .ArchiveNum.ForeColor = CLR_DISABLED_FG
570       .ArchiveNum.BackColor = CLR_DISABLED_BG
580       .ArchiveDate.ForeColor = CLR_DISABLED_FG
590       .ArchiveDate.BackColor = CLR_DISABLED_BG

600       If lngArchiveNum > 0& Then
            ' ** Default to 1 year later.
610         strMon = Right("00" & CStr(month(CDate(.ArchiveDate))), 2)
620         strDay = Right("00" & CStr(day(CDate(.ArchiveDate))), 2)
630         strYear = CStr(year(CDate(.ArchiveDate)) + 1)
640         If strMon = "02" Then
650           If strDay = "29" Then
660             strDay = "28"
670           ElseIf strDay = "28" Then
680             If IsDate(strMon & "/29/" & strYear) = True Then
690               strDay = "29"
700             End If
710           End If
720         End If
730         datEndDate = CDate(strMon & "/" & strDay & "/" & strYear)
740         varTmp00 = DLookup("[Statement_Date]", "Statement Date")
750         Select Case IsNull(varTmp00)
            Case True
760           .DateEnd = Null
770         Case False
780           If CDate(varTmp00) < datEndDate Then
790             .DateEnd = Null
800           Else
810             .DateEnd = datEndDate
820             blnNoMsg = True
830             DateEnd_Exit Cancel  ' ** Procedure: Below.
840             blnNoMsg = False
850           End If
860         End Select
870       End If

880     End With

EXITP:
890     Exit Sub

ERRH:
900     Select Case ERR.Number
        Case Else
910       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
920     End Select
930     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyEscape
1070        intRetVal = 0
1080        DoCmd.Close acForm, THIS_NAME
1090        DoCmd.OpenForm "frmMenu_Utility"
1100      End Select
1110    End If

        ' ** Alt keys.
1120    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1130      Select Case intRetVal
          Case vbKeyX
1140        intRetVal = 0
1150        cmdClose_Click  ' ** Procedure: Below.
1160      End Select
1170    End If

EXITP:
1180    KeyCode = intRetVal
1190    Exit Sub

ERRH:
1200    intRetVal = 0
1210    THAT_PROC = THIS_PROC
1220    That_Erl = Erl
1230    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1240    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

1310    With Me
1320      .TimerInterval = 0&
1330      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1340        lngGTR_Stat = lngGTR_Stat + 1&
1350        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1360        lngTmp01 = lngTmp01 + 1&
1370        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1380        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1390        If lngTmp01 > 1& Then
1400          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1410        End If
1420        If lngTmp01 = 24& Then
1430          lngTmp01 = 0&
1440          lngTmp02 = lngTmp02 + 1&
1450          If lngTmp02 < 5& Then
1460            .TimerInterval = 50&
1470          Else
1480            blnGTR_Emblem = False
1490          End If
1500        Else
1510          .TimerInterval = 50&
1520        End If
1530      End If  ' ** blnGTR_Emblem.
1540      If gblnSetFocus = True Then
1550        gblnSetFocus = False
1560        strFormName = vbNullString
1570  On Error Resume Next
1580        Do Until Screen.ActiveForm.Name = strFormName
1590          DoCmd.SelectObject acForm, THIS_NAME, False
1600          strFormName = Screen.ActiveForm.Name
1610        Loop
1620  On Error GoTo ERRH
1630      End If
1640      If blnGoingToReport = False Then
1650        Select Case gstrReturningForm
            Case "frmRpt_ArchivedTransactions"
1660          .cmdClose.SetFocus
1670        Case Else
1680          If lngArchiveNum > 0& Then
1690            .cmdArchiveReports.SetFocus
1700          ElseIf .cmdArchive.Enabled = True Then
1710            .cmdArchive.SetFocus
1720          Else
1730            .DateEnd.SetFocus
1740          End If
1750        End Select
1760        gstrReturningForm = vbNullString
1770      End If
1780      If gblnGoToReport = True Then
1790        DoCmd.Hourglass True  ' ** Make sure it's still running.
1800        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
1810        Else
1820          If blnGTR_Emblem = False Then
1830            blnGTR_Emblem = True
1840            lngTmp01 = 0&: lngTmp02 = 0&
1850          End If
1860          DoEvents
1870          Select Case blnGoingToReport
              Case True
1880            .TimerInterval = 50&
1890            cmdArchiveReports_Click  ' ** Procedure: Below.
1900          Case False
1910            DoCmd.Hourglass True  ' ** Make sure it's still running.
1920            DoEvents
1930            blnGoingToReport = True
1940            .GoToReport_arw_rpt_img.Visible = True
1950            DoEvents
1960            If .cmdArchiveReports.Enabled = True Then
1970              .cmdArchiveReports.SetFocus
1980              DoEvents
1990              .TimerInterval = GTR_WAIT
2000            Else
                  ' ** Just sit here, we're done.
2010              blnGoingToReport = False
2020              DoCmd.Hourglass False
2030              gblnGoToReport = False
2040              blnGTR_Emblem = False
2050              GTREmblem_Off  ' ** Procedure: Below.
2060              Beep
2070            End If
2080          End Select
2090        End If
2100      Else
2110        DoCmd.Hourglass False
2120      End If
2130    End With

EXITP:
2140    Exit Sub

ERRH:
2150    DoCmd.Hourglass False
2160    blnGoingToReport = False
2170    gblnGoToReport = False
2180    Select Case ERR.Number
        Case Else
2190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2200    End Select
2210    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2310    Select Case DataErr
        Case 2090  ' ** Microsoft Access can't move the focus to the control |.
          ' ** Ignore.
2320    Case 2113  ' ** The value you entered isn't valid for this field.
2330      Select Case Me.ActiveControl.Name
          Case "DateEnd"
2340        If Me.DateEnd.text <> "__/__/____" Then
2350          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
2360        Else
2370          Me.DateEnd.text = vbNullString
2380        End If
2390        Response = acDataErrContinue
2400      End Select
2410    Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
2420      Select Case Me.ActiveControl.Name
          Case "DateEnd"
2430        If Me.DateEnd.text <> "__/__/____" Then
2440          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
2450        Else
2460          Me.DateEnd.text = vbNullString
2470        End If
2480        Response = acDataErrContinue
2490      End Select
2500    Case Else
2510      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2520      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2530    End Select

EXITP:
2540    Exit Sub

ERRH:
2550    Select Case ERR.Number
        Case Else
2560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2570    End Select
2580    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String

2610    If blnJustClose = False Then
2620      gblnSetFocus = True
2630      gstrReturningForm = THIS_NAME
2640      Select Case strCallingForm
          Case "frmMenu_Utility"
2650        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2660      Case vbNullString
2670        strDocName = "frmMenu_Utility"
2680        DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
2690      Case Else
2700        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
2710      End Select
2720    End If

EXITP:
2730    Exit Sub

ERRH:
2740    Select Case ERR.Number
        Case Else
2750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2760    End Select
2770    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2810    DoCmd.Close acForm, THIS_NAME

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2920    intShiftDown = (Shift And acShiftMask) > 0
2930    intAltDown = (Shift And acAltMask) > 0
2940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case intRetVal
          Case vbKeyTab
2970        With Me
2980          intRetVal = 0
2990          If .cmdArchiveReports.Enabled = True Then
3000            .cmdArchiveReports.SetFocus
3010          Else
3020            .DateEnd.SetFocus
3030          End If
3040        End With
3050      End Select
3060    End If

        ' ** Shift keys.
3070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3080      Select Case intRetVal
          Case vbKeyTab
3090        With Me
3100          intRetVal = 0
3110          If .cmdArchive.Enabled = True Then
3120            .cmdArchive.SetFocus
3130          Else
3140            .DateEnd.SetFocus
3150          End If
3160        End With
3170      End Select
3180    End If

EXITP:
3190    KeyCode = intRetVal
3200    Exit Sub

ERRH:
3210    intRetVal = 0
3220    THAT_PROC = THIS_PROC
3230    That_Erl = Erl
3240    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3250    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3320    intShiftDown = (Shift And acShiftMask) > 0
3330    intAltDown = (Shift And acAltMask) > 0
3340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3370        With Me
3380          intRetVal = 0
3390          If .cmdArchive.Enabled = True Then
3400            .cmdArchive.SetFocus
3410          Else
3420            .cmdClose.SetFocus
3430          End If
3440        End With
3450      End Select
3460    End If

        ' ** Shift keys.
3470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3480      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3490        With Me
3500          intRetVal = 0
3510          If .cmdArchiveReports.Enabled = True Then
3520            .cmdArchiveReports.SetFocus
3530          Else
3540            .cmdClose.SetFocus
3550          End If
3560        End With
3570      End Select
3580    End If

EXITP:
3590    KeyCode = intRetVal
3600    Exit Sub

ERRH:
3610    intRetVal = 0
3620    THAT_PROC = THIS_PROC
3630    That_Erl = Erl
3640    Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control |.
          ' ** Ignore.
3650    Case Else
3660      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3670    End Select
3680    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

        Dim varTmp00 As Variant

3710    With Me
3720      .cmdArchive.Enabled = False
3730      If IsNull(.DateEnd) = True Then
3740        If IsNull(.DateEnd.text) = True Then
              '.DateEnd = datMaxArchDate
3750        Else
3760          If Trim(.DateEnd.text) = vbNullString Then
                '.DateEnd = datMaxArchDate
3770          Else
3780            If .DateEnd.text = "__/__/____" Then
                  '.DateEnd = datMaxArchDate
3790            Else
3800              If IsDate(.DateEnd.text) = False Then
3810                If blnNoMsg = False Then
3820                  MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3830                  Cancel = -1
3840                End If
3850              Else
                    ' ** Let it proceed.
3860                .cmdArchive.Enabled = True
3870                If blnNoMsg = False Then
3880                  .TimerInterval = 50&
3890                End If
3900              End If
3910            End If
3920          End If
3930        End If
3940      Else
3950        If IsDate(.DateEnd) = True Then
3960          If .DateEnd > datMaxArchDate Then
3970            If blnNoMsg = False Then
3980              MsgBox "The archive 'Through' date must be no later than " & Format(datMaxArchDate, "mm/dd/yyy") & ".", _
                    vbExclamation + vbOKOnly, "Invalid Date"
3990              Cancel = -1
4000            End If
4010          Else
4020            varTmp00 = DLookup("[Statement_Date]", "Statement Date")
4030            If IsNull(varTmp00) = False Then
4040              If CDate(varTmp00) < .DateEnd Then
4050                If blnNoMsg = False Then
4060                  MsgBox "You have only run statements through " & Format(varTmp00, "mm/dd/yyyy") & "." & vbCrLf & vbCrLf & _
                        "Please select an archive date earlier than this, or run statements" & vbCrLf & _
                        "up to and including the date you wish to archive through.", vbExclamation + vbOKOnly, "Invalid Date"
4070                  .Undo
4080                  Cancel = -1
4090                End If
4100              Else
                    ' ** Date looks good, so enable the button.
4110                .cmdArchive.Enabled = True
4120                If blnNoMsg = False Then
4130                  .TimerInterval = 50&
4140                End If
4150              End If
4160            Else
4170              If blnNoMsg = False Then
4180                MsgBox "You have run no statements." & vbCrLf & vbCrLf & _
                      "Run statements up to and including the date" & vbCrLf & _
                      "you wish to archive through.", vbExclamation + vbOKOnly, "Invalid Date"
4190                .Undo
4200                Cancel = -1
4210              End If
4220            End If
4230          End If
4240        Else
4250          If blnNoMsg = False Then
4260            MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4270            Cancel = -1
4280          End If
4290        End If
4300      End If
4310    End With

EXITP:
4320    Exit Sub

ERRH:
4330    THAT_PROC = THIS_PROC
4340    That_Erl = Erl
4350    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4360    Resume EXITP

End Sub

Private Sub cmdArchiveReports_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchiveReports_Click"

        Dim strDocName As String

4410    DoCmd.Hourglass True
4420    DoEvents

4430    gblnSetFocus = True
4440    blnJustClose = True
4450    strDocName = "frmRpt_ArchivedTransactions"
4460    DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & CStr(CLng(datArchiveDate))
4470    If gblnGoToReport = True Then
4480      Forms(strDocName).TimerInterval = 50&
4490      blnGoingToReport = False
4500      blnGTR_Emblem = False
4510    End If
4520    DoCmd.Close acForm, THIS_NAME

EXITP:
4530    Exit Sub

ERRH:
4540    Select Case ERR.Number
        Case Else
4550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4560    End Select
4570    Resume EXITP

End Sub

Private Sub cmdArchiveReports_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchiveReports_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab
4670        With Me
4680          intRetVal = 0
4690          .DateEnd.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab
4750        With Me
4760          intRetVal = 0
4770          .cmdClose.SetFocus
4780        End With
4790      End Select
4800    End If

EXITP:
4810    KeyCode = intRetVal
4820    Exit Sub

ERRH:
4830    intRetVal = 0
4840    THAT_PROC = THIS_PROC
4850    That_Erl = Erl
4860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4870    Resume EXITP

End Sub

Private Sub cmdArchive_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchive_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstFrom As DAO.Recordset, rstTo As DAO.Recordset
        Dim dbsArch As DAO.Database, dbsData As DAO.Database
        Dim datDateEnd As Date
        Dim intRecords As Long, intArchivedCount As Long
        Dim strSQL As String, strStrucCompare As String, strFieldName As String
        Dim strArchiveDatabase As String, strDatabase As String
        Dim blnArchFound As Boolean, blnArchLinkFound As Boolean
        Dim intLoop As Integer
        Dim blnContinue As Boolean

4910    intArchivedCount = 0
4920    blnContinue = True  ' ** Unless proven otherwise.

4930    With Me

4940      datDateEnd = CDate(.DateEnd)

4950      Set dbs = CurrentDb
4960      With dbs
            ' ** Ledger, grouped, with RecCount, by specified [datend].
4970        Set qdf = .QueryDefs("qryArchiveTransactions_08")
4980        With qdf.Parameters
4990          ![datEnd] = datDateEnd
5000        End With
5010        Set rstFrom = qdf.OpenRecordset
5020        With rstFrom
5030          If .BOF = True And .EOF = True Then
5040            blnContinue = False
5050            MsgBox "No records for transactions before " + Format(datDateEnd, "mm/dd/yyyy") + ".", _
                  vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
5060          Else
5070            .MoveFirst
5080            If ![RecCount] = 0 Then
5090              blnContinue = False
5100              MsgBox "No records for transactions before " + Format(datDateEnd, "mm/dd/yyyy") + ".", _
                    vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
5110            End If
5120          End If
5130          .Close
5140        End With
5150        .Close
5160      End With

5170      If blnContinue = True Then
5180        If MsgBox("WARNING: The archived transactions will no longer be available " & _
                "for any system functions including statements." & vbCrLf & _
                "Please be sure you have run statements which include ALL account activity through " & _
                Format(datDateEnd, "mm/dd/yyyy") & "." & vbCrLf & _
                "See the manual for other reports affected." & vbCrLf & vbCrLf & _
                "Are you sure you want to archive transactions though " & Format(datDateEnd, "mm/dd/yyyy") & " at this time?", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Archive") <> vbYes Then
5190          blnContinue = False
5200        End If
5210        If blnContinue = True Then
5220          If MsgBox("Please click 'Yes' again to confirm that you wish" & vbCrLf & "to archive transactions at this time.", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Archive") <> vbYes Then
5230            blnContinue = False
5240          End If
5250        End If  ' ** blnContinue.
5260      End If  ' ** blnContinue.

5270      If blnContinue = True Then

5280        DoCmd.Hourglass True

5290        If gstrTrustDataLocation = vbNullString Then
5300          blnContinue = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
5310        End If

5320        strArchiveDatabase = gstrTrustDataLocation & gstrFile_ArchDataName
5330        Set dbsArch = DBEngine.Workspaces(0).OpenDatabase(strArchiveDatabase)

5340        strDatabase = gstrTrustDataLocation & gstrFile_DataName
5350        strArchiveDatabase = gstrTrustDataLocation & gstrFile_ArchDataName

5360        Set dbsData = DBEngine.Workspaces(0).OpenDatabase(strDatabase)

            ' ** Make sure the archive ledger table exists.
5370        blnArchFound = False  ' ** Unless proven otherwise.
5380        intLoop = 0
5390        Do While (intLoop < dbsArch.TableDefs.Count) And (blnArchFound = False)
5400          If dbsArch.TableDefs(intLoop).Name = "Ledger" Then blnArchFound = True
5410          intLoop = intLoop + 1
5420        Loop

            ' ** Make sure the archive database Ledger table is linked.
5430        blnArchLinkFound = False  ' ** Unless proven otherwise.
5440        intLoop = 0
5450        Do While (intLoop < CurrentDb.TableDefs.Count) And (blnArchLinkFound = False)
5460          If CurrentDb.TableDefs(intLoop).Name = "LedgerArchive" Then blnArchLinkFound = True
5470          intLoop = intLoop + 1
5480        Loop

5490        If (blnArchLinkFound = False) And (blnArchFound = True) Then   ' ** Try to create the link.
              ' ** If they've never archived before, then set up the new table.
5500          DoCmd.TransferDatabase acLink, "Microsoft Access", strArchiveDatabase, acTable, "ledger", "LedgerArchive"
5510          blnArchLinkFound = True  ' ** It exists now!
5520        End If
5530        DoCmd.Hourglass False

5540      Else
5550        .DateEnd.SetFocus
5560      End If  ' ** blnContinue.

5570      If blnContinue = True Then
5580        If blnArchFound = True Then   ' ** There should be a link, so check the record count.

              ' ** Try to get the record count from LedgerArchive.
5590          Set dbs = CurrentDb
5600          With dbs
                ' ** LedgerArchive, grouped, with RecCount.
5610            Set qdf = .QueryDefs("qryArchiveTransactions_09")
5620            Set rstTo = qdf.OpenRecordset
5630            With rstTo
5640              If .BOF = True And .EOF = True Then
5650                intRecords = 0
5660              Else
5670                .MoveFirst
5680                intRecords = ![RecCount]
5690              End If
5700              .Close
5710            End With
5720            .Close
5730          End With

5740          strStrucCompare = CompareTableStructure("Ledger", "LedgerArchive")  ' ** Module Function: modFileUtilities.
              ' ** If LedgerArchive has the field revcode_KB, chances are
              ' ** they haven't used it yet, so it will be recopied below.
              ' ** strStrucCompare will return "DIFFERENT: revcode_KD"
5750          If strStrucCompare = RET_ERR Then
5760            MsgBox "An error occurred comparing table structures.", vbExclamation + vbOKOnly, "Unable To Complete Operation"
5770            blnContinue = False
5780          End If
5790        Else
5800          intRecords = 0                     '  ** Table not found, so can be no records.
5810          strStrucCompare = "DIFFERENT"      '  ** Well, the archive ledger doesn't exist!
5820        End If
5830      End If  ' ** blnContinue.

5840      If blnContinue = True Then
5850        If (intRecords > 0) And (strStrucCompare <> "SAME") Then
5860          MsgBox "Table structures differ, but the archive table has records already." & vbCrLf & vbCrLf & _
                "Please contact Delta Data, Inc." & vbCrLf & vbCrLf & strStrucCompare, vbExclamation + vbOKOnly, "Unable To Complete Operation"
5870          blnContinue = False
5880        End If
5890      End If

5900      If blnContinue = True Then
5910        If strStrucCompare <> "SAME" Then
              ' ** Zero archived records, but different structures, so we need to make the archive
              ' ** structure the same as the main structure.

5920          If blnArchLinkFound = True Then
5930            TableDelete ("LedgerArchive")  ' ** Module Function: modFileUtilities.
5940          End If

              ' ** Drop the table.
5950          If blnArchFound = True Then
5960            strSQL = "DROP TABLE Ledger;"
5970            dbsArch.Execute strSQL
5980          End If

5990          DoCmd.TransferDatabase acImport, "Microsoft Access", strDatabase, acTable, "Ledger", "LedgerTemp", True
6000          DoCmd.TransferDatabase acExport, "Microsoft Access", strArchiveDatabase, acTable, "LedgerTemp", "Ledger", True
6010          CurrentDb.TableDefs.Delete "LedgerTemp"

6020          If blnContinue = True Then
                ' ** Recreate the link from this database to the new archive table.
6030            DoCmd.TransferDatabase acLink, "Microsoft Access", strArchiveDatabase, acTable, "ledger", "LedgerArchive"
6040          End If

6050        End If
6060      End If

6070      If blnContinue = True Then

6080        Set rstTo = CurrentDb.OpenRecordset("SELECT * FROM [LedgerArchive]", dbOpenDynaset)
6090        Set rstFrom = CurrentDb.OpenRecordset("SELECT * FROM ledger WHERE ledger.transdate <= #" & Format(datDateEnd, "mm/dd/yyyy") & "#" & _
              "ORDER BY ledger.transdate, ledger.accountno, ledger.assetno", dbOpenDynaset)

6100        DoCmd.Hourglass True
6110        DBEngine.Workspaces(0).BeginTrans

6120  On Error GoTo ERRH_TRANS

6130        rstFrom.MoveLast
6140        rstFrom.MoveFirst
6150        Do While Not rstFrom.EOF
6160          rstTo.AddNew
6170          intLoop = 0
6180          Do While (intLoop < rstFrom.Fields.Count) And blnContinue = True
6190            strFieldName = rstFrom.Fields(intLoop).Name
6200            rstTo.Fields(strFieldName) = rstFrom.Fields(strFieldName)
6210            intLoop = intLoop + 1
6220          Loop
6230          rstTo.Update  ' ** Add the archive ledger record.
6240          rstFrom.Delete ' ** Delete the main ledger record.
6250          intArchivedCount = intArchivedCount + 1
6260          rstFrom.MoveNext  ' ** Next record to archive.
6270        Loop
6280        DBEngine.Workspaces(0).CommitTrans

6290  On Error GoTo ERRH

6300        zErrorLogWriter ("End Archive Process " & Format(Now(), "mm/dd/yy hh:mm:ss:dd"))

6310        DoEvents
6320        RecsTot  ' ** Procedure: Below.
6330        DoCmd.Hourglass False
6340        DoEvents

6350      End If

6360      If blnContinue = True Then
6370        .cmdArchiveReports.Enabled = True
6380        MsgBox CStr(intArchivedCount) & " records archived through " & Format(datDateEnd, "mm/dd/yyyy") & "." & vbCrLf & vbCrLf & _
              "It is now recommended that you repair and compact your Trust Data, and that you do a fresh backup." & vbCrLf & vbCrLf & _
              "Backups made prior to this archiving should not be used until contacting Delta Data, Inc.", vbExclamation + vbOKOnly, ("Finished" & Space(40))
6390      End If

6400    End With

EXITP:
6410    DoCmd.Hourglass False
6420  On Error Resume Next
6430    rstTo.Close
6440    rstFrom.Close
6450    dbsArch.Close
6460    dbsData.Close
6470    Set rstTo = Nothing
6480    Set rstFrom = Nothing
6490    Set dbsArch = Nothing
6500    Set dbsData = Nothing
6510    Set qdf = Nothing
6520    Set dbs = Nothing
6530    Exit Sub

ERRH_TRANS:
6540    DBEngine.Workspaces(0).Rollback

ERRH:
6550    DoCmd.Hourglass False
6560    blnContinue = False
6570    Select Case ERR.Number
        Case Else
6580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6590    End Select
6600    Resume EXITP

End Sub

Private Sub cmdArchive_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6720    intShiftDown = (Shift And acShiftMask) > 0
6730    intAltDown = (Shift And acAltMask) > 0
6740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6760      Select Case intRetVal
          Case vbKeyTab
6770        With Me
6780          intRetVal = 0
6790          .cmdClose.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab
6850        With Me
6860          intRetVal = 0
6870          .DateEnd.SetFocus
6880        End With
6890      End Select
6900    End If

EXITP:
6910    KeyCode = intRetVal
6920    Exit Sub

ERRH:
6930    intRetVal = 0
6940    THAT_PROC = THIS_PROC
6950    That_Erl = Erl
6960    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6970    Resume EXITP

End Sub

Private Sub RecsTot()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "RecsTot"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

7010    Set dbs = CurrentDb
7020    With dbs
          ' ** LedgerArchive, grouped, with max transdate, cnt.
7030      Set qdf = .QueryDefs("qryArchiveTransactions_02")
7040      Set rst = qdf.OpenRecordset
7050      With rst
7060        .MoveFirst
7070        lngArchiveNum = ![cnt]
7080        If ![cnt] > 0 Then
7090          datArchiveDate = ![transdate]
7100        Else
7110          datArchiveDate = 0
7120        End If
7130        .Close
7140      End With
7150      .Close
7160    End With
7170    With Me
7180      If lngArchiveNum > 0& Then
7190        .cmdArchiveReports_lbl.ForeColor = CLR_VDKGRY
7200        .cmdArchiveReports_lbl_dim_hi.Visible = False
7210        .ArchiveNum = lngArchiveNum
7220        .ArchiveNum.Locked = True
7230        .ArchiveNum.BorderColor = CLR_LTBLU2
7240        .ArchiveDate = datArchiveDate
7250        .ArchiveDate.Locked = True
7260        .ArchiveDate.BorderColor = CLR_LTBLU2
7270        .cmdArchiveReports.Enabled = True
7280      Else
7290        .cmdArchiveReports_lbl.ForeColor = WIN_CLR_DISF
7300        .cmdArchiveReports_lbl_dim_hi.Visible = True
7310        .ArchiveNum = 0&
7320        .ArchiveNum.Locked = False
7330        .ArchiveNum.BorderColor = WIN_CLR_DISR
7340        .ArchiveDate = Null
7350        .ArchiveDate.Locked = False
7360        .ArchiveDate.BorderColor = WIN_CLR_DISR
7370        .cmdArchiveReports.Enabled = False
7380      End If
7390    End With

EXITP:
7400    Set rst = Nothing
7410    Set qdf = Nothing
7420    Set dbs = Nothing
7430    Exit Sub

ERRH:
7440    Select Case ERR.Number
        Case Else
7450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7460    End Select
7470    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

7510    With Me
7520      blnGTR_Emblem = False
7530      For lngX = 1& To 24&
7540        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
7550      Next
7560    End With

EXITP:
7570    Exit Sub

ERRH:
7580    Select Case ERR.Number
        Case Else
7590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7600    End Select
7610    Resume EXITP

End Sub
