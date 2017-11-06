VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_StatementOfCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_StatementOfCondition"

'VGC 07/16/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   All:              A {opgAccountType_optAll}
' **   Close:            C {cmdClose}
' **   Discretionary:    D {opgAccountType_optDisc}
' **   Excel:            E {cmdExcel}
' **   As Of:            F {DateAsOf}
' **   NonDiscretionary: N {opgAccountType_optNonDisc}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Include Closed:   U {chkIncludeClosed}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** For GTR, it just opens and stops.

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private blnDateChangeMsg As Boolean
Private lngDiscYes As Long, lngDiscNo As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           gstrFormQuerySpec = THIS_NAME

160           .DateAsOf = Null
170           blnDateChangeMsg = False
180           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
190           DoEvents

200           .opgAccountType = .opgAccountType_optAll.OptionValue

210           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
220           .cmdExcel.Enabled = False
      #Else
230           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberDates!
240           If IsNull(.DateAsOf) = False Then
250             DateAsOf_Exit Cancel  ' ** Procedure: Below.
260           End If

270         Else
280           Cancel = -1
290         End If
300       Else
310         Cancel = -1
320       End If
330     End With

340     If Cancel = -1 Then
350       Beep
360       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
370     End If

EXITP:
380     Exit Sub

ERRH:
390     DoCmd.Hourglass False
400     Select Case ERR.Number
        Case Else
410       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
420     End Select
430     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

510     With Me

          ' ** Create an instance of the Calendar class.
520       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
530       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
540       clsMonthClass.PositionAtCursor = True

550       Set dbs = CurrentDb

          ' ** Default Null Discretion to 'No'.
560       Set qdf = dbs.QueryDefs("qryStatementOfCondition_21")
570       qdf.Execute
580       Set qdf = Nothing

          ' ** Retrieve the Discretionary/Non-Discretionary counts.
590       Select Case .chkIncludeClosed
          Case True
            ' ** qryStatementOfCondition_23c (Union of qryStatementOfCondition_23a
            ' ** (Account, for all accounts, cnt_yes), qryStatementOfCondition_23b
            ' ** (Account, for all accounts, cnt_no), for all accounts),
            ' ** grouped to give 1 record with both counts, for dateclosed = Null.
600         Set qdf = dbs.QueryDefs("qryStatementOfCondition_23d")
610       Case False
            ' ** qryStatementOfCondition_22c (Union of qryStatementOfCondition_22a
            ' ** (Account, for dateclosed = Null, cnt_yes), qryStatementOfCondition_22b
            ' ** (Account, for dateclosed = Null, cnt_no), for dateclosed = Null),
            ' **  grouped to give 1 record with both counts, for all accounts.
620         Set qdf = dbs.QueryDefs("qryStatementOfCondition_22d")
630       End Select
640       Set rst = qdf.OpenRecordset
650       With rst
660         .MoveFirst
670         lngDiscYes = ![cnt_yes]
680         lngDiscNo = ![cnt_no]
690         .Close
700       End With
710       Set rst = Nothing
720       Set qdf = Nothing

730       dbs.Close
740       Set dbs = Nothing

750       .cnt_yes = lngDiscYes
760       .cnt_no = lngDiscNo

770       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

780       opgAccountType_AfterUpdate  ' ** Procedure: Below.
790       chkRememberDates_AfterUpdate  ' ** Procedure: Below.
800       chkPageOf_AfterUpdate  ' ** Procedure: Below.
810       chkIncludeClosed_AfterUpdate  ' ** Procedure: Below.

820       If IsNull(.DateAsOf) = True Then
830         .DateAsOf = TodaysDate  ' ** Module Function: modUtilities.
840       End If

850       DoCmd.Hourglass False

860     End With

EXITP:
870     Set rst = Nothing
880     Set qdf = Nothing
890     Set dbs = Nothing
900     Exit Sub

ERRH:
910     DoCmd.Hourglass False
920     Select Case ERR.Number
        Case Else
930       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
940     End Select
950     Resume EXITP

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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyEscape
1070        intRetVal = 0
1080        cmdClose_Click  ' ** Procedure: Below.
1090      Case vbKeyF7
1100        With Me
1110          intRetVal = 0
1120          If .cmdCalendar.Enabled = True Then
1130            .cmdCalendar.SetFocus
1140            cmdCalendar_Click  ' ** Procedure: Below.
1150          End If
1160        End With
1170      End Select
1180    End If

        ' ** Alt keys.
1190    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1200      Select Case intRetVal
          Case vbKeyX
1210        intRetVal = 0
1220        cmdClose_Click  ' ** Procedure: Below.
1230      End Select
1240    End If

        ' ** Ctrl keys.
1250    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyG
1270        With Me
1280          intRetVal = 0
1290          If .chkPageOf.Enabled = True Then
1300            .chkPageOf.SetFocus
1310            .chkPageOf = (Not .chkPageOf)
1320            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1330          End If
1340        End With
1350      End Select
1360    End If

        ' ** Ctrl-Shift keys.
1370    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1380      Select Case intRetVal
          Case vbKeyF
1390        With Me
1400          intRetVal = 0
1410          .FocusHolder.SetFocus
1420        End With
1430      End Select
1440    End If

EXITP:
1450    KeyCode = intRetVal
1460    Exit Sub

ERRH:
1470    intRetVal = 0
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub Form_Timer()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1610    With Me
1620      .TimerInterval = 0&
1630      If gblnSetFocus = True Then
1640        gblnSetFocus = False
1650        strFormName = vbNullString
1660  On Error Resume Next
1670        Do Until strFormName = THIS_NAME
1680          DoCmd.SelectObject acForm, THIS_NAME, False
1690          strFormName = Screen.ActiveForm.Name
1700        Loop
1710  On Error GoTo ERRH
1720      End If
1730      If gblnGoToReport = True Then
1740        If .cmdPrint.Enabled = True Then
1750          .cmdPrint.SetFocus
1760        Else
1770          .cmdClose.SetFocus
1780        End If
1790        gblnGoToReport = False
1800        Beep
1810        DoCmd.Hourglass False
1820        DoEvents
1830      Else
1840        DoCmd.Hourglass False
1850      End If
1860    End With

EXITP:
1870    Exit Sub

ERRH:
1880    DoCmd.Hourglass False
1890    Select Case ERR.Number
        Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2010    DoCmd.Hourglass True
2020    DoEvents

2030    With Me
2040      If .chkRememberDates = False Then
2050        .DateAsOf = Null  ' ** So it won't be there when Pref_Save() runs.
2060        DoEvents
2070      End If
2080    End With

2090    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2100    If Not clsMonthClass Is Nothing Then
2110      If clsMonthClass.IsCalendar Then
2120        Cancel = -1
2130      Else
2140        Set clsMonthClass = Nothing
2150      End If
2160    End If

        ' ** Make sure all reports are closed.
2170    Do While Reports.Count > 0
2180      DoCmd.Close acReport, Reports(0).Name
2190    Loop

2200    gstrFormQuerySpec = vbNullString
2210    gblnSetFocus = True
2220    gstrReturningForm = THIS_NAME
2230    Select Case strCallingForm
        Case "frmMenu_Report"
2240      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2250    End Select

EXITP:
2260    Exit Sub

ERRH:
2270    Select Case ERR.Number
        Case Else
2280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2290    End Select
2300    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2410    With Me
2420      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
2430        Select Case blnCalendar1_Focus
            Case True
2440          .cmdCalendar_raised_semifocus_dots_img.Visible = True
2450          .cmdCalendar_raised_img.Visible = False
2460        Case False
2470          .cmdCalendar_raised_img.Visible = True
2480          .cmdCalendar_raised_semifocus_dots_img.Visible = False
2490        End Select
2500        .cmdCalendar_raised_focus_img.Visible = False
2510        .cmdCalendar_raised_focus_dots_img.Visible = False
2520        .cmdCalendar_sunken_focus_dots_img.Visible = False
2530        .cmdCalendar_raised_img_dis.Visible = False
2540      End If
2550    End With

EXITP:
2560    Exit Sub

ERRH:
2570    Select Case ERR.Number
        Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2710    DoCmd.Hourglass True
2720    DoEvents

2730    DoCmd.Close acForm, THIS_NAME

EXITP:
2740    Exit Sub

ERRH:
2750    Select Case ERR.Number
        Case Else
2760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2770    End Select
2780    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab
2870        With Me
2880          intRetVal = 0
2890          .DateAsOf.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2940      Select Case intRetVal
          Case vbKeyTab
2950        With Me
2960          intRetVal = 0
2970          If .cmdExcel.Enabled = True Then
2980            .cmdExcel.SetFocus
2990          Else
3000            .cmdWord.SetFocus
3010          End If
3020        End With
3030      End Select
3040    End If

EXITP:
3050    KeyCode = intRetVal
3060    Exit Sub

ERRH:
3070    intRetVal = 0
3080    Select Case ERR.Number
        Case Else
3090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3100    End Select
3110    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

3210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub DateAsOf_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_KeyDown"

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
3390          .opgAccountType.SetFocus
3400        End With
3410      End Select
3420    End If

        ' ** Shift keys.
3430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3450        With Me
3460          intRetVal = 0
3470          .cmdClose.SetFocus
3480        End With
3490      End Select
3500    End If

        ' ** Ctrl keys.
3510    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3520      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3530        With Me
3540          intRetVal = 0
3550          .cmdCalendar.SetFocus
3560        End With
3570      End Select
3580    End If

EXITP:
3590    KeyCode = intRetVal
3600    Exit Sub

ERRH:
3610    intRetVal = 0
3620    Select Case ERR.Number
        Case Else
3630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3640    End Select
3650    Resume EXITP

End Sub

Private Sub DateAsOf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_MouseMove"

3710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub DateAsOf_Exit(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_Exit"

3810    With Me
3820      If IsNull(.DateAsOf) = True Then
            ' ** Populate it with today's date.
3830        .DateAsOf = Date
3840      Else
3850        If Trim(.DateAsOf) = vbNullString Then
              ' ** Populate it with today's date.
3860          .DateAsOf = Date
3870        Else
3880          blnDateChangeMsg = True
3890        End If
3900      End If
3910    End With

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

4010    With Me
4020      blnCalendar1_Focus = True
4030      .cmdCalendar_raised_semifocus_dots_img.Visible = True
4040      .cmdCalendar_raised_img.Visible = False
4050      .cmdCalendar_raised_focus_img.Visible = False
4060      .cmdCalendar_raised_focus_dots_img.Visible = False
4070      .cmdCalendar_sunken_focus_dots_img.Visible = False
4080      .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

4210    With Me
4220      blnCalendar1_MouseDown = True
4230      .cmdCalendar_sunken_focus_dots_img.Visible = True
4240      .cmdCalendar_raised_img.Visible = False
4250      .cmdCalendar_raised_semifocus_dots_img.Visible = False
4260      .cmdCalendar_raised_focus_img.Visible = False
4270      .cmdCalendar_raised_focus_dots_img.Visible = False
4280      .cmdCalendar_raised_img_dis.Visible = False
4290    End With

EXITP:
4300    Exit Sub

ERRH:
4310    Select Case ERR.Number
        Case Else
4320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4330    End Select
4340    Resume EXITP

End Sub

Private Sub cmdCalendar_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

4410    With Me
4420      datStartDate = Date
4430      datEndDate = 0
4440      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
4450      If blnRetVal = True Then
4460        .DateAsOf = datStartDate
4470      Else
4480        .DateAsOf = CDate(Format(Date, "mm/dd/yyyy"))
4490      End If
4500      .DateAsOf.SetFocus
4510    End With

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

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
4690          .chkRememberDates.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab
4750        With Me
4760          intRetVal = 0
4770          .DateAsOf.SetFocus
4780        End With
4790      End Select
4800    End If

EXITP:
4810    KeyCode = intRetVal
4820    Exit Sub

ERRH:
4830    intRetVal = 0
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

4910    With Me
4920      If blnCalendar1_MouseDown = False Then
4930        Select Case blnCalendar1_Focus
            Case True
4940          .cmdCalendar_raised_focus_dots_img.Visible = True
4950          .cmdCalendar_raised_focus_img.Visible = False
4960        Case False
4970          .cmdCalendar_raised_focus_img.Visible = True
4980          .cmdCalendar_raised_focus_dots_img.Visible = False
4990        End Select
5000        .cmdCalendar_raised_img.Visible = False
5010        .cmdCalendar_raised_semifocus_dots_img.Visible = False
5020        .cmdCalendar_sunken_focus_dots_img.Visible = False
5030        .cmdCalendar_raised_img_dis.Visible = False
5040      End If
5050    End With

EXITP:
5060    Exit Sub

ERRH:
5070    Select Case ERR.Number
        Case Else
5080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5090    End Select
5100    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

5210    With Me
5220      .cmdCalendar_raised_focus_dots_img.Visible = True
5230      .cmdCalendar_raised_img.Visible = False
5240      .cmdCalendar_raised_semifocus_dots_img.Visible = False
5250      .cmdCalendar_raised_focus_img.Visible = False
5260      .cmdCalendar_sunken_focus_dots_img.Visible = False
5270      .cmdCalendar_raised_img_dis.Visible = False
5280      blnCalendar1_MouseDown = False
5290    End With

EXITP:
5300    Exit Sub

ERRH:
5310    Select Case ERR.Number
        Case Else
5320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5330    End Select
5340    Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

5410    With Me
5420      .cmdCalendar_raised_img.Visible = True
5430      .cmdCalendar_raised_semifocus_dots_img.Visible = False
5440      .cmdCalendar_raised_focus_img.Visible = False
5450      .cmdCalendar_raised_focus_dots_img.Visible = False
5460      .cmdCalendar_sunken_focus_dots_img.Visible = False
5470      .cmdCalendar_raised_img_dis.Visible = False
5480      blnCalendar1_Focus = False
5490    End With

EXITP:
5500    Exit Sub

ERRH:
5510    Select Case ERR.Number
        Case Else
5520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5530    End Select
5540    Resume EXITP

End Sub

Private Sub chkRememberDates_Enter()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_Enter"

5610    If blnDateChangeMsg = True Then
5620      blnDateChangeMsg = False
5630      DateChangeMsg  ' ** Procedure: Below.
5640    End If

EXITP:
5650    Exit Sub

ERRH:
5660    Select Case ERR.Number
        Case Else
5670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5680    End Select
5690    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

5710    With Me
5720      Select Case .chkRememberDates
          Case True
5730        .chkRememberDates_lbl.FontBold = True
5740      Case False
5750        .chkRememberDates_lbl.FontBold = False
5760      End Select
5770    End With

EXITP:
5780    Exit Sub

ERRH:
5790    Select Case ERR.Number
        Case Else
5800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5810    End Select
5820    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5920    intShiftDown = (Shift And acShiftMask) > 0
5930    intAltDown = (Shift And acAltMask) > 0
5940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5970        With Me
5980          intRetVal = 0
5990          .opgAccountType.SetFocus
6000        End With
6010      End Select
6020    End If

        ' ** Shift keys.
6030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6050        With Me
6060          intRetVal = 0
6070          .cmdCalendar.SetFocus
6080        End With
6090      End Select
6100    End If

EXITP:
6110    KeyCode = intRetVal
6120    Exit Sub

ERRH:
6130    intRetVal = 0
6140    Select Case ERR.Number
        Case Else
6150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6160    End Select
6170    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

6210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6220    Exit Sub

ERRH:
6230    Select Case ERR.Number
        Case Else
6240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6250    End Select
6260    Resume EXITP

End Sub

Private Sub opgAccountType_Enter()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountType_Enter"

6310    If blnDateChangeMsg = True Then
6320      blnDateChangeMsg = False
6330      DateChangeMsg  ' ** Procedure: Below.
6340    End If

EXITP:
6350    Exit Sub

ERRH:
6360    Select Case ERR.Number
        Case Else
6370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6380    End Select
6390    Resume EXITP

End Sub

Private Sub opgAccountType_AfterUpdate()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountType_AfterUpdate"

6410    With Me
6420      Select Case .opgAccountType
          Case .opgAccountType_optAll.OptionValue
6430        .opgAccountType_optAll_lbl.FontBold = True
6440        .opgAccountType_optDisc_lbl.FontBold = False
6450        .opgAccountType_optNonDisc_lbl.FontBold = False
6460        .opgAccountType_optAll_lbl_box.Visible = True
6470        .opgAccountType_optDisc_lbl_box.Visible = False
6480        .opgAccountType_optNonDisc_lbl_box.Visible = False
6490      Case .opgAccountType_optDisc.OptionValue
6500        .opgAccountType_optAll_lbl.FontBold = False
6510        .opgAccountType_optDisc_lbl.FontBold = True
6520        .opgAccountType_optNonDisc_lbl.FontBold = False
6530        .opgAccountType_optAll_lbl_box.Visible = False
6540        .opgAccountType_optDisc_lbl_box.Visible = True
6550        .opgAccountType_optNonDisc_lbl_box.Visible = False
6560      Case .opgAccountType_optNonDisc.OptionValue
6570        .opgAccountType_optAll_lbl.FontBold = False
6580        .opgAccountType_optDisc_lbl.FontBold = False
6590        .opgAccountType_optNonDisc_lbl.FontBold = True
6600        .opgAccountType_optAll_lbl_box.Visible = False
6610        .opgAccountType_optDisc_lbl_box.Visible = False
6620        .opgAccountType_optNonDisc_lbl_box.Visible = True
6630      End Select
6640    End With

EXITP:
6650    Exit Sub

ERRH:
6660    Select Case ERR.Number
        Case Else
6670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6680    End Select
6690    Resume EXITP

End Sub

Private Sub opgAccountType_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountType_optAll_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
6770        With Me
6780          intRetVal = 0
6790          .cmdPreview.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6850        With Me
6860          intRetVal = 0
6870          .DateAsOf.SetFocus
6880        End With
6890      End Select
6900    End If

        ' ** Ctrl keys.
6910    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6920      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6930        With Me
6940          intRetVal = 0
6950          .chkIncludeClosed.SetFocus
6960        End With
6970      End Select
6980    End If

        ' ** Ctrl-Shift keys.
6990    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7000      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7010        With Me
7020          intRetVal = 0
7030          .chkRememberDates.SetFocus
7040        End With
7050      End Select
7060    End If

EXITP:
7070    KeyCode = intRetVal
7080    Exit Sub

ERRH:
7090    intRetVal = 0
7100    Select Case ERR.Number
        Case Else
7110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7120    End Select
7130    Resume EXITP

End Sub

Private Sub opgAccountType_optDisc_KeyDown(KeyCode As Integer, Shift As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountType_optDisc_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7220    intShiftDown = (Shift And acShiftMask) > 0
7230    intAltDown = (Shift And acAltMask) > 0
7240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7270        With Me
7280          intRetVal = 0
7290          .cmdPreview.SetFocus
7300        End With
7310      End Select
7320    End If

        ' ** Shift keys.
7330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7350        With Me
7360          intRetVal = 0
7370          .DateAsOf.SetFocus
7380        End With
7390      End Select
7400    End If

        ' ** Ctrl keys.
7410    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7420      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7430        With Me
7440          intRetVal = 0
7450          .chkIncludeClosed.SetFocus
7460        End With
7470      End Select
7480    End If

        ' ** Ctrl-Shift keys.
7490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7500      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7510        With Me
7520          intRetVal = 0
7530          .chkRememberDates.SetFocus
7540        End With
7550      End Select
7560    End If

EXITP:
7570    KeyCode = intRetVal
7580    Exit Sub

ERRH:
7590    intRetVal = 0
7600    Select Case ERR.Number
        Case Else
7610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7620    End Select
7630    Resume EXITP

End Sub

Private Sub opgAccountType_optNonDisc_KeyDown(KeyCode As Integer, Shift As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountType_optNonDisc_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7720    intShiftDown = (Shift And acShiftMask) > 0
7730    intAltDown = (Shift And acAltMask) > 0
7740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7770        With Me
7780          intRetVal = 0
7790          .cmdPreview.SetFocus
7800        End With
7810      End Select
7820    End If

        ' ** Shift keys.
7830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7850        With Me
7860          intRetVal = 0
7870          .DateAsOf.SetFocus
7880        End With
7890      End Select
7900    End If

        ' ** Ctrl keys.
7910    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7920      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7930        With Me
7940          intRetVal = 0
7950          .chkIncludeClosed.SetFocus
7960        End With
7970      End Select
7980    End If

        ' ** Ctrl-Shift keys.
7990    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8000      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8010        With Me
8020          intRetVal = 0
8030          .chkRememberDates.SetFocus
8040        End With
8050      End Select
8060    End If

EXITP:
8070    KeyCode = intRetVal
8080    Exit Sub

ERRH:
8090    intRetVal = 0
8100    Select Case ERR.Number
        Case Else
8110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8120    End Select
8130    Resume EXITP

End Sub

Private Sub chkIncludeClosed_AfterUpdate()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_AfterUpdate"

8210    With Me
8220      Select Case .chkIncludeClosed
          Case True
8230        .chkIncludeClosed_lbl.FontBold = True
8240      Case False
8250        .chkIncludeClosed_lbl.FontBold = False
8260      End Select
8270    End With

EXITP:
8280    Exit Sub

ERRH:
8290    Select Case ERR.Number
        Case Else
8300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8310    End Select
8320    Resume EXITP

End Sub

Private Sub chkIncludeClosed_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8420    intShiftDown = (Shift And acShiftMask) > 0
8430    intAltDown = (Shift And acAltMask) > 0
8440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8470        With Me
8480          intRetVal = 0
8490          .chkPageOf.SetFocus
8500        End With
8510      End Select
8520    End If

        ' ** Shift keys.
8530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8550        With Me
8560          intRetVal = 0
8570          .opgAccountType.SetFocus
8580        End With
8590      End Select
8600    End If

EXITP:
8610    KeyCode = intRetVal
8620    Exit Sub

ERRH:
8630    intRetVal = 0
8640    Select Case ERR.Number
        Case Else
8650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8660    End Select
8670    Resume EXITP

End Sub

Private Sub chkPageOf_Enter()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_Enter"

8710    If blnDateChangeMsg = True Then
8720      blnDateChangeMsg = False
8730      DateChangeMsg  ' ** Procedure: Below.
8740    End If

EXITP:
8750    Exit Sub

ERRH:
8760    Select Case ERR.Number
        Case Else
8770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8780    End Select
8790    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

8810    With Me
8820      Select Case .chkPageOf
          Case True
8830        .chkPageOf_lbl.FontBold = True
8840      Case False
8850        .chkPageOf_lbl.FontBold = False
8860      End Select
8870    End With

EXITP:
8880    Exit Sub

ERRH:
8890    DoCmd.Hourglass False
8900    Select Case ERR.Number
        Case Else
8910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8920    End Select
8930    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9020    intShiftDown = (Shift And acShiftMask) > 0
9030    intAltDown = (Shift And acAltMask) > 0
9040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9070        With Me
9080          intRetVal = 0
9090          .cmdPreview.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9150        With Me
9160          intRetVal = 0
9170          .chkIncludeClosed.SetFocus
9180        End With
9190      End Select
9200    End If

EXITP:
9210    KeyCode = intRetVal
9220    Exit Sub

ERRH:
9230    intRetVal = 0
9240    Select Case ERR.Number
        Case Else
9250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9260    End Select
9270    Resume EXITP

End Sub

Private Sub cmdPreview_Enter()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Enter"

9310    If blnDateChangeMsg = True Then
9320      blnDateChangeMsg = False
9330      DateChangeMsg  ' ** Procedure: Below.
9340    End If

EXITP:
9350    Exit Sub

ERRH:
9360    Select Case ERR.Number
        Case Else
9370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9380    End Select
9390    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

9410    If DoReport = True Then  ' ** Function: Below.
9420      strDocName = "rptStatementOfCondition"
9430      DoCmd.OpenReport strDocName, acViewPreview
9440      DoCmd.Maximize
9450      DoCmd.RunCommand acCmdFitToWindow
9460    End If

EXITP:
9470    Exit Sub

ERRH:
9480    Select Case ERR.Number
        Case Else
9490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9500    End Select
9510    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab
9670        With Me
9680          intRetVal = 0
9690          .cmdPrint.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab
9750        With Me
9760          intRetVal = 0
9770          .opgAccountType.SetFocus
9780        End With
9790      End Select
9800    End If

        ' ** Ctrl-Shift keys.
9810    If intCtrlDown And (Not intAltDown) And intShiftDown Then
9820      Select Case intRetVal
          Case vbKeyTab
9830        With Me
9840          intRetVal = 0
9850          .chkPageOf.SetFocus
9860        End With
9870      End Select
9880    End If

EXITP:
9890    KeyCode = intRetVal
9900    Exit Sub

ERRH:
9910    intRetVal = 0
9920    Select Case ERR.Number
        Case Else
9930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9940    End Select
9950    Resume EXITP

End Sub

Private Sub cmdPrint_Enter()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Enter"

10010   If blnDateChangeMsg = True Then
10020     blnDateChangeMsg = False
10030     DateChangeMsg  ' ** Procedure: Below.
10040   End If

EXITP:
10050   Exit Sub

ERRH:
10060   Select Case ERR.Number
        Case Else
10070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10080   End Select
10090   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

10110   If DoReport = True Then  ' ** Function: Below.
10120     strDocName = "rptStatementOfCondition"
10130     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
10140       DoCmd.OpenReport strDocName, acViewPreview
10150       DoCmd.Maximize
10160       DoCmd.RunCommand acCmdFitToWindow
10170     Else
            '##GTR_Ref: rptStatementOfCondition
10180       DoCmd.OpenReport strDocName, acViewNormal
10190     End If
10200   End If

EXITP:
10210   Exit Sub

ERRH:
10220   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
10230   Case Else
10240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10250   End Select
10260   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10320   intShiftDown = (Shift And acShiftMask) > 0
10330   intAltDown = (Shift And acAltMask) > 0
10340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10360     Select Case intRetVal
          Case vbKeyTab
10370       With Me
10380         intRetVal = 0
10390         .cmdWord.SetFocus
10400       End With
10410     End Select
10420   End If

        ' ** Shift keys.
10430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10440     Select Case intRetVal
          Case vbKeyTab
10450       With Me
10460         intRetVal = 0
10470         .cmdPreview.SetFocus
10480       End With
10490     End Select
10500   End If

EXITP:
10510   KeyCode = intRetVal
10520   Exit Sub

ERRH:
10530   intRetVal = 0
10540   Select Case ERR.Number
        Case Else
10550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10560   End Select
10570   Resume EXITP

End Sub

Private Sub cmdWord_Enter()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Enter"

10610   If blnDateChangeMsg = True Then
10620     blnDateChangeMsg = False
10630     DateChangeMsg  ' ** Procedure: Below.
10640   End If

EXITP:
10650   Exit Sub

ERRH:
10660   Select Case ERR.Number
        Case Else
10670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10680   End Select
10690   Resume EXITP

End Sub

Private Sub cmdWord_Click()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

10710   With Me

10720     If DoReport = True Then  ' ** Function: Below.

10730       If IsNull(.UserReportPath) = True Then
10740         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
10750       Else
10760         strRptPath = .UserReportPath
10770       End If

10780       strRptCap = "rptStatementOfCondition_"
10790       Select Case .opgAccountType
            Case .opgAccountType_optAll.OptionValue
10800         strRptCap = strRptCap & "All_"
10810       Case .opgAccountType_optDisc.OptionValue
10820         strRptCap = strRptCap & "Discretionary_"
10830       Case .opgAccountType_optNonDisc.OptionValue
10840         strRptCap = strRptCap & "NonDiscretionary_"
10850       End Select
10860       strRptCap = strRptCap & Format(Date, "yyyymmdd")

10870       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

10880       If strRptPathFile <> vbNullString Then
10890         gstrFormQuerySpec = THIS_NAME
10900         strRpt = "rptStatementOfCondition"
10910         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
10920         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
10930       End If

10940     End If

10950   End With

EXITP:
10960   Exit Sub

ERRH:
10970   Select Case ERR.Number
        Case Else
10980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10990   End Select
11000   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11120   intShiftDown = (Shift And acShiftMask) > 0
11130   intAltDown = (Shift And acAltMask) > 0
11140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11160     Select Case intRetVal
          Case vbKeyTab
11170       With Me
11180         intRetVal = 0
11190         If .cmdExcel.Enabled = True Then
11200           .cmdExcel.SetFocus
11210         Else
11220           .cmdClose.SetFocus
11230         End If
11240       End With
11250     End Select
11260   End If

        ' ** Shift keys.
11270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11280     Select Case intRetVal
          Case vbKeyTab
11290       With Me
11300         intRetVal = 0
11310         .cmdPrint.SetFocus
11320       End With
11330     End Select
11340   End If

EXITP:
11350   KeyCode = intRetVal
11360   Exit Sub

ERRH:
11370   intRetVal = 0
11380   Select Case ERR.Number
        Case Else
11390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11400   End Select
11410   Resume EXITP

End Sub

Private Sub cmdExcel_Enter()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Enter"

11510   If blnDateChangeMsg = True Then
11520     blnDateChangeMsg = False
11530     DateChangeMsg  ' ** Procedure: Below.
11540   End If

EXITP:
11550   Exit Sub

ERRH:
11560   Select Case ERR.Number
        Case Else
11570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11580   End Select
11590   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String
        Dim strRptCap As String, strRptPath As String, strRptPathFile As String

11610   With Me

11620     If DoReport = True Then  ' ** Function: Below.

11630       If IsNull(.UserReportPath) = True Then
11640         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
11650       Else
11660         strRptPath = .UserReportPath
11670       End If
11680       strRptCap = "rptStatementOfCondition_"
11690       Select Case .opgAccountType
            Case .opgAccountType_optAll.OptionValue
11700         strRptCap = strRptCap & "All_"
11710       Case .opgAccountType_optDisc.OptionValue
11720         strRptCap = strRptCap & "Discretionary_"
11730       Case .opgAccountType_optNonDisc.OptionValue
11740         strRptCap = strRptCap & "NonDiscretionary_"
11750       End Select
11760       strRptCap = strRptCap & Format(Date, "yyyymmdd")

11770       strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

11780       If strRptPathFile <> vbNullString Then
11790         gstrFormQuerySpec = THIS_NAME
11800         Select Case .opgAccountType
              Case .opgAccountType_optAll.OptionValue
                ' ** qryStatementOfCondition_39c (xx), For Export.
11810           strQry = "qryStatementOfCondition_40c"
11820         Case .opgAccountType_optDisc.OptionValue
                ' ** qryStatementOfCondition_39b (xx), For Export.
11830           strQry = "qryStatementOfCondition_40b"
11840         Case .opgAccountType_optNonDisc.OptionValue
                ' ** qryStatementOfCondition_39a (xx), For Export.
11850           strQry = "qryStatementOfCondition_40a"
11860         End Select
11870         DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
11880         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
11890       End If

11900     End If

11910   End With

EXITP:
11920   Exit Sub

ERRH:
11930   Select Case ERR.Number
        Case Else
11940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11950   End Select
11960   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12020   intShiftDown = (Shift And acShiftMask) > 0
12030   intAltDown = (Shift And acAltMask) > 0
12040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12060     Select Case intRetVal
          Case vbKeyTab
12070       With Me
12080         intRetVal = 0
12090         .cmdClose.SetFocus
12100       End With
12110     End Select
12120   End If

        ' ** Shift keys.
12130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12140     Select Case intRetVal
          Case vbKeyTab
12150       With Me
12160         intRetVal = 0
12170         .cmdWord.SetFocus
12180       End With
12190     End Select
12200   End If

EXITP:
12210   KeyCode = intRetVal
12220   Exit Sub

ERRH:
12230   intRetVal = 0
12240   Select Case ERR.Number
        Case Else
12250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12260   End Select
12270   Resume EXITP

End Sub

Private Function DoReport() As Boolean

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim blnRetVal As Boolean

12310   blnRetVal = True

12320   With Me
12330     If .opgAccountType = .opgAccountType_optDisc.OptionValue And .cnt_yes = 0 Then
12340       blnRetVal = False
12350       MsgBox "There are no Discretionary accounts.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
12360     Else
12370       If .opgAccountType = .opgAccountType_optNonDisc.OptionValue And .cnt_no = 0 Then
12380         blnRetVal = False
12390         MsgBox "There are no Non-Discretionary accounts.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
12400       End If
12410     End If
12420   End With

EXITP:
12430   DoReport = blnRetVal
12440   Exit Function

ERRH:
12450   blnRetVal = False
12460   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
12470   Case Else
12480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12490   End Select
12500   Resume EXITP

End Function

Private Sub DateChangeMsg()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "DateChangeMsg"

12610   With Me
12620     If IsNull(.DateAsOf) = False Then
12630       If .DateAsOf < Date Then
12640         MsgBox "Data is not rolled back to the entered date." & vbCrLf & _
                "The report reflects current information.", vbInformation + vbOKOnly, "Current Data Only"
12650       ElseIf .DateAsOf > Date Then
12660         MsgBox "Report will reflect current information.", vbInformation + vbOKOnly, "Current Data Only"
12670       End If
12680     End If
12690   End With

EXITP:
12700   Exit Sub

ERRH:
12710   Select Case ERR.Number
        Case Else
12720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12730   End Select
12740   Resume EXITP

End Sub
