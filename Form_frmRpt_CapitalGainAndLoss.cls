VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CapitalGainAndLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CapitalGainAndLoss"

'VGC 09/23/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   From:             F {DateStart}
' **   Show All Sales:   H {chkShowAllSales}
' **   All:              L {opgAccountNumber_optAll}
' **   Name:             M {opgAccountSource_optName}
' **   Sort By Name:     N {opgOrderBy_optName}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Sort By Number:   R {opgOrderBy_optNumber}
' **   To:               T {DateEnd}
' **   Specified:        U {opgAccountNumber_optSpecified}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean

Private blnExcel As Boolean
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = .OpenArgs

150           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
160           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False

170           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

180           gstrFormQuerySpec = THIS_NAME

190           blnExcel = False

      #If NoExcel Then
200           .cmdExcel.Enabled = False
      #Else
210           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!
              ' ** EVENT CHECK: chkRememberDates!
220           If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
230             DateStart_Exit Cancel  ' ** Procedure: Below.
240             DateEnd_Exit Cancel  ' ** Procedure: Below.
250           End If

260           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
270           chkShowAllSales_AfterUpdate  ' ** Procedure: Below.
280           chkRememberMe_AfterUpdate  ' ** Procedure: Below.
290           chkRememberDates_AfterUpdate  ' ** Procedure: Below.
300           opgOrderBy_AfterUpdate  ' ** Procedure: Below.
310           chkPageOf_AfterUpdate  ' ** Procedure: Below.
320           opgAccountNumber_AfterUpdate  ' ** Procedure: Below.

330         Else
340           Cancel = -1
350         End If
360       Else
370         Cancel = -1
380       End If
390     End With

400     If Cancel = -1 Then
410       Beep
420       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
430     End If

440     DoCmd.Hourglass False

EXITP:
450     Exit Sub

ERRH:
460     DoCmd.Hourglass False
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
620       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
630       clsMonthClass.hWndForm = .hwnd
          ' ** Let's default to PositionAtCursor.
640       clsMonthClass.PositionAtCursor = True

650       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

660     End With

EXITP:
670     Exit Sub

ERRH:
680     DoCmd.Hourglass False
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Number:           B {opgAccountSource_optNumber}
        ' **   Name:             M {opgAccountSource_optName}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       Case vbKeyF7
900         intRetVal = 0
910         cmdCalendar1_Click  ' ** Procedure: Below.
920       Case vbKeyF8
930         intRetVal = 0
940         cmdCalendar2_Click  ' ** Procedure: Below.
950       End Select
960     End If

        ' ** Alt keys.
970     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
980       Select Case intRetVal
          Case vbKeyB
990         With Me
1000          intRetVal = 0
1010          If .opgAccountSource.Enabled = True Then
1020            .opgAccountSource = .opgAccountSource_optNumber.OptionValue
1030            .opgAccountSource.SetFocus
1040            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1050          Else
1060            Beep
1070          End If
1080        End With
1090      Case vbKeyM
1100        With Me
1110          intRetVal = 0
1120          If .opgAccountSource.Enabled = True Then
1130            .opgAccountSource = .opgAccountSource_optName.OptionValue
1140            .opgAccountSource.SetFocus
1150            opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1160          Else
1170            Beep
1180          End If
1190        End With
1200      Case vbKeyX
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
1420          .DateStart = Null
1430        End With
1440      End Select
1450    End If

EXITP:
1460    KeyCode = intRetVal
1470    Exit Sub

ERRH:
1480    intRetVal = 0
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1520    Resume EXITP

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

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2010    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
2020      Response = acDataErrContinue
2030    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2040      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
2050        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
2060        Response = acDataErrContinue
2070      End If
2080    Case 2237  ' ** The text you entered isn't an item in the list.
2090      Response = acDataErrDisplay
2100    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
2110      Beep
2120      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
2130      Response = acDataErrContinue
2140    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
2150    Case Else
2160      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2170      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select

EXITP:
2190    Exit Sub

ERRH:
2200    Select Case ERR.Number
        Case Else
2210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2220    End Select
2230    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

2310    DoCmd.Hourglass True
2320    DoEvents

2330    With Me
2340      If .chkRememberDates = False Then
2350        .DateStart = Null  ' ** So it won't be there when Pref_Save() runs.
2360        .DateEnd = Null
2370      End If
2380      If .chkRememberMe = False Then
2390        .cmbAccounts = Null
2400      End If
2410    End With

2420    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2430    Do While Reports.Count > 0
2440      DoCmd.Close acReport, Reports(0).Name
2450    Loop

2460    Set dbs = CurrentDb
2470    With dbs
          ' ** Empty tmpCapitalGainsAndLosses.
2480      Set qdf = .QueryDefs("qryCapitalGainsAndLoss_Excel_08")
2490      qdf.Execute
2500      .Close
2510    End With
2520    Set dbs = Nothing

2530    gstrFormQuerySpec = vbNullString
2540    gblnSetFocus = True
2550    gstrReturningForm = THIS_NAME
2560    Select Case strCallingForm
        Case "frmMenu_Report"
2570      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2580    End Select

EXITP:
2590    Set qdf = Nothing
2600    Set dbs = Nothing
2610    Exit Sub

ERRH:
2620    DoCmd.Hourglass False
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

2710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2810    With Me
2820      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
2830        Select Case blnCalendar1_Focus
            Case True
2840          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
2850          .cmdCalendar1_raised_img.Visible = False
2860        Case False
2870          .cmdCalendar1_raised_img.Visible = True
2880          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
2890        End Select
2900        .cmdCalendar1_raised_focus_dots_img.Visible = False
2910        .cmdCalendar1_raised_focus_img.Visible = False
2920        .cmdCalendar1_sunken_focus_dots_img.Visible = False
2930        .cmdCalendar1_raised_img_dis.Visible = False
2940      End If
2950      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
2960        Select Case blnCalendar2_Focus
            Case True
2970          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
2980          .cmdCalendar2_raised_img.Visible = False
2990        Case False
3000          .cmdCalendar2_raised_img.Visible = True
3010          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
3020        End Select
3030        .cmdCalendar2_raised_focus_dots_img.Visible = False
3040        .cmdCalendar2_raised_focus_img.Visible = False
3050        .cmdCalendar2_sunken_focus_dots_img.Visible = False
3060        .cmdCalendar2_raised_img_dis.Visible = False
3070      End If
3080    End With

EXITP:
3090    Exit Sub

ERRH:
3100    Select Case ERR.Number
        Case Else
3110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3120    End Select
3130    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

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

Private Sub cmdClose_Click()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3310    DoCmd.Hourglass True
3320    DoEvents

3330    DoCmd.Close acForm, THIS_NAME

EXITP:
3340    Exit Sub

ERRH:
3350    DoCmd.Hourglass False
3360    Select Case ERR.Number
        Case Else
3370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3380    End Select
3390    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3420    intShiftDown = (Shift And acShiftMask) > 0
3430    intAltDown = (Shift And acAltMask) > 0
3440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3460      Select Case intRetVal
          Case vbKeyTab
3470        With Me
3480          intRetVal = 0
3490          .DateStart.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab
3550        With Me
3560          intRetVal = 0
3570          If .cmdExcel.Enabled = True Then
3580            .cmdExcel.SetFocus
3590          Else
3600            .cmdWord.SetFocus
3610          End If
3620        End With
3630      End Select
3640    End If

EXITP:
3650    KeyCode = intRetVal
3660    Exit Sub

ERRH:
3670    intRetVal = 0
3680    THAT_PROC = THIS_PROC
3690    That_Erl = Erl
3700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3710    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

3810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3920    intShiftDown = (Shift And acShiftMask) > 0
3930    intAltDown = (Shift And acAltMask) > 0
3940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3970        With Me
3980          intRetVal = 0
3990          .DateEnd.SetFocus
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4050        With Me
4060          intRetVal = 0
4070          .cmdClose.SetFocus
4080        End With
4090      End Select
4100    End If

        ' ** Ctrl keys.
4110    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4120      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4130        With Me
4140          intRetVal = 0
4150          .cmdCalendar1.SetFocus
4160        End With
4170      End Select
4180    End If

EXITP:
4190    KeyCode = intRetVal
4200    Exit Sub

ERRH:
4210    intRetVal = 0
4220    THAT_PROC = THIS_PROC
4230    That_Erl = Erl
4240    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4250    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

4310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub DateStart_Exit(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

4410    With Me
4420      If IsDate(.DateEnd) = True And IsDate(.DateStart) = True Then
4430        If CDate(.DateStart) > CDate(.DateEnd) Then
4440          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
4450          .DateEnd = Null
4460          .DateEnd.SetFocus
4470        End If
4480      End If
4490    End With

EXITP:
4500    Exit Sub

ERRH:
4510    THAT_PROC = THIS_PROC
4520    That_Erl = Erl
4530    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4540    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

4610    With Me
4620      blnCalendar1_Focus = True
4630      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4640      .cmdCalendar1_raised_img.Visible = False
4650      .cmdCalendar1_raised_focus_img.Visible = False
4660      .cmdCalendar1_raised_focus_dots_img.Visible = False
4670      .cmdCalendar1_sunken_focus_dots_img.Visible = False
4680      .cmdCalendar1_raised_img_dis.Visible = False
4690    End With

EXITP:
4700    Exit Sub

ERRH:
4710    Select Case ERR.Number
        Case Else
4720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4730    End Select
4740    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

4810    With Me
4820      blnCalendar1_MouseDown = True
4830      .cmdCalendar1_sunken_focus_dots_img.Visible = True
4840      .cmdCalendar1_raised_img.Visible = False
4850      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
4860      .cmdCalendar1_raised_focus_img.Visible = False
4870      .cmdCalendar1_raised_focus_dots_img.Visible = False
4880      .cmdCalendar1_raised_img_dis.Visible = False
4890    End With

EXITP:
4900    Exit Sub

ERRH:
4910    Select Case ERR.Number
        Case Else
4920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4930    End Select
4940    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

5010    With Me
5020      datStartDate = Date
5030      datEndDate = 0
5040      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
5050      If blnRetVal = True Then
5060        .DateStart = datStartDate
5070      Else
5080        .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
5090      End If
5100      .DateStart.SetFocus
5110    End With

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
5140    Case Else
5150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5160    End Select
5170    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

5210    With Me
5220      If blnCalendar1_MouseDown = False Then
5230        Select Case blnCalendar1_Focus
            Case True
5240          .cmdCalendar1_raised_focus_dots_img.Visible = True
5250          .cmdCalendar1_raised_focus_img.Visible = False
5260        Case False
5270          .cmdCalendar1_raised_focus_img.Visible = True
5280          .cmdCalendar1_raised_focus_dots_img.Visible = False
5290        End Select
5300        .cmdCalendar1_raised_img.Visible = False
5310        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5320        .cmdCalendar1_sunken_focus_dots_img.Visible = False
5330        .cmdCalendar1_raised_img_dis.Visible = False
5340      End If
5350    End With

EXITP:
5360    Exit Sub

ERRH:
5370    Select Case ERR.Number
        Case Else
5380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5390    End Select
5400    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

5510    With Me
5520      .cmdCalendar1_raised_focus_dots_img.Visible = True
5530      .cmdCalendar1_raised_img.Visible = False
5540      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5550      .cmdCalendar1_raised_focus_img.Visible = False
5560      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5570      .cmdCalendar1_raised_img_dis.Visible = False
5580      blnCalendar1_MouseDown = False
5590    End With

EXITP:
5600    Exit Sub

ERRH:
5610    Select Case ERR.Number
        Case Else
5620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5630    End Select
5640    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

5710    With Me
5720      .cmdCalendar1_raised_img.Visible = True
5730      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5740      .cmdCalendar1_raised_focus_img.Visible = False
5750      .cmdCalendar1_raised_focus_dots_img.Visible = False
5760      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5770      .cmdCalendar1_raised_img_dis.Visible = False
5780      blnCalendar1_Focus = False
5790    End With

EXITP:
5800    Exit Sub

ERRH:
5810    Select Case ERR.Number
        Case Else
5820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5830    End Select
5840    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

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
5990          .opgAccountNumber.SetFocus
6000        End With
6010      End Select
6020    End If

        ' ** Shift keys.
6030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6050        With Me
6060          intRetVal = 0
6070          .DateStart.SetFocus
6080        End With
6090      End Select
6100    End If

        ' ** Ctrl keys.
6110    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6120      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6130        With Me
6140          intRetVal = 0
6150          .cmdCalendar2.SetFocus
6160        End With
6170      End Select
6180    End If

        ' ** Ctrl-Shift keys.
6190    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6200      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6210        With Me
6220          intRetVal = 0
6230          .cmdCalendar1.SetFocus
6240        End With
6250      End Select
6260    End If

EXITP:
6270    KeyCode = intRetVal
6280    Exit Sub

ERRH:
6290    intRetVal = 0
6300    THAT_PROC = THIS_PROC
6310    That_Erl = Erl
6320    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6330    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

6410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

6510    With Me

6520      If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
6530        .DateEnd = Date
6540      Else
6550        If IsNull(.DateStart) = False Then
6560          If CDate(.DateStart) > CDate(.DateEnd) Then
6570            MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
6580            .DateStart = Null
6590            .DateStart.SetFocus
6600          End If
6610        End If
6620      End If
6630    End With

EXITP:
6640    Exit Sub

ERRH:
6650    THAT_PROC = THIS_PROC
6660    That_Erl = Erl
6670    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6680    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

6710    With Me
6720      blnCalendar2_Focus = True
6730      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
6740      .cmdCalendar2_raised_img.Visible = False
6750      .cmdCalendar2_raised_focus_img.Visible = False
6760      .cmdCalendar2_raised_focus_dots_img.Visible = False
6770      .cmdCalendar2_sunken_focus_dots_img.Visible = False
6780      .cmdCalendar2_raised_img_dis.Visible = False
6790    End With

EXITP:
6800    Exit Sub

ERRH:
6810    Select Case ERR.Number
        Case Else
6820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6830    End Select
6840    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

6910    With Me
6920      blnCalendar2_MouseDown = True
6930      .cmdCalendar2_sunken_focus_dots_img.Visible = True
6940      .cmdCalendar2_raised_img.Visible = False
6950      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
6960      .cmdCalendar2_raised_focus_img.Visible = False
6970      .cmdCalendar2_raised_focus_dots_img.Visible = False
6980      .cmdCalendar2_raised_img_dis.Visible = False
6990    End With

EXITP:
7000    Exit Sub

ERRH:
7010    Select Case ERR.Number
        Case Else
7020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7030    End Select
7040    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

7110    With Me
7120      datStartDate = Date
7130      datEndDate = 0
7140      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
7150      If blnRetVal = True Then
7160        .DateEnd = datStartDate
7170      Else
7180        .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
7190      End If
7200      .DateEnd.SetFocus
7210      Cancel = 0
7220      DateEnd_Exit Cancel  ' ** Procedure: Above.
7230      If Cancel = 0 Then
7240        .opgAccountNumber.SetFocus
7250      End If
7260    End With

EXITP:
7270    Exit Sub

ERRH:
7280    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7290    Case Else
7300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7310    End Select
7320    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7420    intShiftDown = (Shift And acShiftMask) > 0
7430    intAltDown = (Shift And acAltMask) > 0
7440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7470        With Me
7480          intRetVal = 0
7490          .chkRememberDates.SetFocus
7500        End With
7510      End Select
7520    End If

        ' ** Shift keys.
7530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7550        With Me
7560          intRetVal = 0
7570          .DateEnd.SetFocus
7580        End With
7590      End Select
7600    End If

EXITP:
7610    KeyCode = intRetVal
7620    Exit Sub

ERRH:
7630    intRetVal = 0
7640    THAT_PROC = THIS_PROC
7650    That_Erl = Erl
7660    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7670    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

7710    With Me
7720      If blnCalendar2_MouseDown = False Then
7730        Select Case blnCalendar2_Focus
            Case True
7740          .cmdCalendar2_raised_focus_dots_img.Visible = True
7750          .cmdCalendar2_raised_focus_img.Visible = False
7760        Case False
7770          .cmdCalendar2_raised_focus_img.Visible = True
7780          .cmdCalendar2_raised_focus_dots_img.Visible = False
7790        End Select
7800        .cmdCalendar2_raised_img.Visible = False
7810        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
7820        .cmdCalendar2_sunken_focus_dots_img.Visible = False
7830        .cmdCalendar2_raised_img_dis.Visible = False
7840      End If
7850    End With

EXITP:
7860    Exit Sub

ERRH:
7870    Select Case ERR.Number
        Case Else
7880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7890    End Select
7900    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

8010    With Me
8020      .cmdCalendar2_raised_focus_dots_img.Visible = True
8030      .cmdCalendar2_raised_img.Visible = False
8040      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8050      .cmdCalendar2_raised_focus_img.Visible = False
8060      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8070      .cmdCalendar2_raised_img_dis.Visible = False
8080      blnCalendar2_MouseDown = False
8090    End With

EXITP:
8100    Exit Sub

ERRH:
8110    Select Case ERR.Number
        Case Else
8120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8130    End Select
8140    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

8210    With Me
8220      .cmdCalendar2_raised_img.Visible = True
8230      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8240      .cmdCalendar2_raised_focus_img.Visible = False
8250      .cmdCalendar2_raised_focus_dots_img.Visible = False
8260      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8270      .cmdCalendar2_raised_img_dis.Visible = False
8280      blnCalendar2_Focus = False
8290    End With

EXITP:
8300    Exit Sub

ERRH:
8310    Select Case ERR.Number
        Case Else
8320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8330    End Select
8340    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

8410    With Me
8420      Select Case .chkRememberDates
          Case True
8430        .chkRememberDates_lbl.FontBold = True
8440      Case False
8450        .chkRememberDates_lbl.FontBold = False
8460      End Select
8470    End With

EXITP:
8480    Exit Sub

ERRH:
8490    Select Case ERR.Number
        Case Else
8500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8510    End Select
8520    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8620    intShiftDown = (Shift And acShiftMask) > 0
8630    intAltDown = (Shift And acAltMask) > 0
8640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8670        With Me
8680          intRetVal = 0
8690          .opgAccountNumber.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8750        With Me
8760          intRetVal = 0
8770          .cmdCalendar2.SetFocus
8780        End With
8790      End Select
8800    End If

EXITP:
8810    KeyCode = intRetVal
8820    Exit Sub

ERRH:
8830    intRetVal = 0
8840    THAT_PROC = THIS_PROC
8850    That_Erl = Erl
8860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8870    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

8910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
8920    Exit Sub

ERRH:
8930    Select Case ERR.Number
        Case Else
8940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8950    End Select
8960    Resume EXITP

End Sub

Private Sub opgAccountNumber_AfterUpdate()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

9010    With Me
9020      Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
9030        .opgAccountNumber_optSpecified_lbl.FontBold = True
9040        .opgAccountNumber_optAll_lbl.FontBold = False
9050        .opgAccountNumber_optSpecified_lbl_box.Visible = True
9060        .opgAccountNumber_optAll_lbl_box.Visible = False
9070        .cmbAccounts.Enabled = True
9080        .cmbAccounts.BorderColor = CLR_LTBLU2
9090        .cmbAccounts.BackStyle = acBackStyleNormal
9100        .opgAccountSource.Enabled = True
9110        .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
9120        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
9130        .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
9140        .opgAccountSource_optName_lbl2_dim_hi.Visible = False
9150        .chkRememberMe.Enabled = True
9160        .chkRememberMe_lbl.Visible = True
9170        .chkRememberMe_lbl2_dim.Visible = False
9180        .chkRememberMe_lbl2_dim_hi.Visible = False
9190      Case .opgAccountNumber_optAll.OptionValue
9200        .opgAccountNumber_optSpecified_lbl.FontBold = False
9210        .opgAccountNumber_optAll_lbl.FontBold = True
9220        .opgAccountNumber_optSpecified_lbl_box.Visible = False
9230        .opgAccountNumber_optAll_lbl_box.Visible = True
9240        .cmbAccounts.Enabled = False
9250        .cmbAccounts.BorderColor = WIN_CLR_DISR
9260        .cmbAccounts.BackStyle = acBackStyleTransparent
9270        .opgAccountSource.Enabled = False
9280        .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
9290        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
9300        .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
9310        .opgAccountSource_optName_lbl2_dim_hi.Visible = True
9320        .chkRememberMe.Enabled = False
9330        .chkRememberMe_lbl.Visible = False
9340        .chkRememberMe_lbl2_dim.Visible = True
9350        .chkRememberMe_lbl2_dim_hi.Visible = True
9360      End Select
9370    End With

EXITP:
9380    Exit Sub

ERRH:
9390    Select Case ERR.Number
        Case Else
9400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9410    End Select
9420    Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9520    intShiftDown = (Shift And acShiftMask) > 0
9530    intAltDown = (Shift And acAltMask) > 0
9540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9570        With Me
9580          intRetVal = 0
9590          If .cmbAccounts.Enabled = True Then
9600            .cmbAccounts.SetFocus
9610          Else
9620            .cmdPreview.SetFocus
9630          End If
9640        End With
9650      End Select
9660    End If

        ' ** Shift keys.
9670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9680      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9690        With Me
9700          intRetVal = 0
9710          .DateEnd.SetFocus
9720        End With
9730      End Select
9740    End If

        ' ** Ctrl keys.
9750    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          intRetVal = 0
9790          .chkShowAllSales.SetFocus
9800        End With
9810      End Select
9820    End If

        ' ** Ctrl-Shift keys.
9830    If intCtrlDown And (Not intAltDown) And intShiftDown Then
9840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9850        With Me
9860          intRetVal = 0
9870          .chkRememberDates.SetFocus
9880        End With
9890      End Select
9900    End If

EXITP:
9910    KeyCode = intRetVal
9920    Exit Sub

ERRH:
9930    intRetVal = 0
9940    THAT_PROC = THIS_PROC
9950    That_Erl = Erl
9960    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9970    Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10020   intShiftDown = (Shift And acShiftMask) > 0
10030   intAltDown = (Shift And acAltMask) > 0
10040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10070       With Me
10080         intRetVal = 0
10090         If .cmbAccounts.Enabled = True Then
10100           .cmbAccounts.SetFocus
10110         Else
10120           .cmdPreview.SetFocus
10130         End If
10140       End With
10150     End Select
10160   End If

        ' ** Shift keys.
10170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10190       With Me
10200         intRetVal = 0
10210         .DateEnd.SetFocus
10220       End With
10230     End Select
10240   End If

        ' ** Ctrl keys.
10250   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10270       With Me
10280         intRetVal = 0
10290         .chkShowAllSales.SetFocus
10300       End With
10310     End Select
10320   End If

        ' ** Ctrl-Shift keys.
10330   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10350       With Me
10360         intRetVal = 0
10370         .chkRememberDates.SetFocus
10380       End With
10390     End Select
10400   End If

EXITP:
10410   KeyCode = intRetVal
10420   Exit Sub

ERRH:
10430   intRetVal = 0
10440   THAT_PROC = THIS_PROC
10450   That_Erl = Erl
10460   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10470   Resume EXITP

End Sub

Private Sub chkShowAllSales_AfterUpdate()

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllSales_AfterUpdate"

10510   With Me
10520     Select Case .chkShowAllSales
          Case True
10530       .chkShowAllSales_lbl.FontBold = True
10540     Case False
10550       .chkShowAllSales_lbl.FontBold = False
10560     End Select
10570   End With

EXITP:
10580   Exit Sub

ERRH:
10590   Select Case ERR.Number
        Case Else
10600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10610   End Select
10620   Resume EXITP

End Sub

Private Sub chkShowAllSales_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAllSales_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         intRetVal = 0
10790         .opgAccountSource.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Shift keys.
10830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10850       With Me
10860         intRetVal = 0
10870         .opgAccountNumber.SetFocus
10880       End With
10890     End Select
10900   End If

EXITP:
10910   KeyCode = intRetVal
10920   Exit Sub

ERRH:
10930   intRetVal = 0
10940   THAT_PROC = THIS_PROC
10950   That_Erl = Erl
10960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10970   Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

11010   strAccountNo = vbNullString

11020   With Me
11030     If IsNull(.cmbAccounts) = False Then
11040       If Len(.cmbAccounts.Column(0)) > 0 Then
11050         strAccountNo = .cmbAccounts.Column(0)
11060       End If
11070     End If
11080     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
11090       .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
11100       .opgAccountSource_optNumber_lbl2.FontBold = True
11110       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
11120       .opgAccountSource_optName_lbl2.FontBold = False
11130       .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
11140     Case .opgAccountSource_optName.OptionValue
11150       .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
11160       .opgAccountSource_optNumber_lbl2.FontBold = False
11170       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
11180       .opgAccountSource_optName_lbl2.FontBold = True
11190       .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
11200     End Select
11210     DoEvents
11220     If strAccountNo <> vbNullString Then
11230       .cmbAccounts = strAccountNo
11240     End If
11250   End With

EXITP:
11260   Exit Sub

ERRH:
11270   Select Case ERR.Number
        Case Else
11280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11290   End Select
11300   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11420   intShiftDown = (Shift And acShiftMask) > 0
11430   intAltDown = (Shift And acAltMask) > 0
11440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11470       With Me
11480         intRetVal = 0
11490         .chkRememberMe.SetFocus
11500       End With
11510     End Select
11520   End If

        ' ** Shift keys.
11530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11550       With Me
11560         intRetVal = 0
11570         .chkShowAllSales.SetFocus
11580       End With
11590     End Select
11600   End If

EXITP:
11610   KeyCode = intRetVal
11620   Exit Sub

ERRH:
11630   intRetVal = 0
11640   THAT_PROC = THIS_PROC
11650   That_Erl = Erl
11660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11670   Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11720   intShiftDown = (Shift And acShiftMask) > 0
11730   intAltDown = (Shift And acAltMask) > 0
11740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11770       With Me
11780         intRetVal = 0
11790         .chkRememberMe.SetFocus
11800       End With
11810     End Select
11820   End If

        ' ** Shift keys.
11830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11850       With Me
11860         intRetVal = 0
11870         .chkShowAllSales.SetFocus
11880       End With
11890     End Select
11900   End If

EXITP:
11910   KeyCode = intRetVal
11920   Exit Sub

ERRH:
11930   intRetVal = 0
11940   THAT_PROC = THIS_PROC
11950   That_Erl = Erl
11960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11970   Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

12010   With Me
12020     Select Case .chkRememberMe
          Case True
12030       .chkRememberMe_lbl.FontBold = True
12040       .chkRememberMe_lbl2_dim.FontBold = True
12050       .chkRememberMe_lbl2_dim_hi.FontBold = True
12060     Case False
12070       .chkRememberMe_lbl.FontBold = False
12080       .chkRememberMe_lbl2_dim.FontBold = False
12090       .chkRememberMe_lbl2_dim_hi.FontBold = False
12100     End Select
12110   End With

EXITP:
12120   Exit Sub

ERRH:
12130   Select Case ERR.Number
        Case Else
12140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12150   End Select
12160   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12220   intShiftDown = (Shift And acShiftMask) > 0
12230   intAltDown = (Shift And acAltMask) > 0
12240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12270       With Me
12280         intRetVal = 0
12290         .cmbAccounts.SetFocus
12300       End With
12310     End Select
12320   End If

        ' ** Shift keys.
12330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12350       With Me
12360         intRetVal = 0
12370         .opgAccountSource.SetFocus
12380       End With
12390     End Select
12400   End If

EXITP:
12410   KeyCode = intRetVal
12420   Exit Sub

ERRH:
12430   intRetVal = 0
12440   THAT_PROC = THIS_PROC
12450   That_Erl = Erl
12460   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12470   Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

12510   With Me
12520     If .cmbAccounts.text = vbNullString Then
12530       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
12540       .cmbAccounts = .cmbAccounts.ItemData(0)
12550     End If
12560   End With

EXITP:
12570   Exit Sub

ERRH:
12580   Select Case ERR.Number
        Case Else
12590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12600   End Select
12610   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12720   intShiftDown = (Shift And acShiftMask) > 0
12730   intAltDown = (Shift And acAltMask) > 0
12740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12770       With Me
12780         intRetVal = 0
12790         .cmdPreview.SetFocus
12800       End With
12810     End Select
12820   End If

        ' ** Shift keys.
12830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12850       With Me
12860         intRetVal = 0
12870         .opgAccountNumber.SetFocus
12880       End With
12890     End Select
12900   End If

        ' ** Ctrl keys.
12910   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
12920     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12930       With Me
12940         intRetVal = 0
12950         .opgOrderBy.SetFocus
12960       End With
12970     End Select
12980   End If

        ' ** Ctrl-Shift keys.
12990   If intCtrlDown And (Not intAltDown) And intShiftDown Then
13000     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13010       With Me
13020         intRetVal = 0
13030         .chkRememberMe.SetFocus
13040       End With
13050     End Select
13060   End If

EXITP:
13070   KeyCode = intRetVal
13080   Exit Sub

ERRH:
13090   intRetVal = 0
13100   THAT_PROC = THIS_PROC
13110   That_Erl = Erl
13120   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13130   Resume EXITP

End Sub

Private Sub opgOrderBy_AfterUpdate()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_AfterUpdate"

13210   With Me
13220     Select Case .opgOrderBy
          Case .opgOrderBy_optNumber.OptionValue
13230       .opgOrderBy_optNumber_lbl.FontBold = True
13240       .opgOrderBy_optName_lbl.FontBold = False
13250     Case .opgOrderBy_optName.OptionValue
13260       .opgOrderBy_optNumber_lbl.FontBold = False
13270       .opgOrderBy_optName_lbl.FontBold = True
13280     End Select
13290   End With

EXITP:
13300   Exit Sub

ERRH:
13310   Select Case ERR.Number
        Case Else
13320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13330   End Select
13340   Resume EXITP

End Sub

Private Sub opgOrderBy_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13420   intShiftDown = (Shift And acShiftMask) > 0
13430   intAltDown = (Shift And acAltMask) > 0
13440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13470       With Me
13480         intRetVal = 0
13490         .chkPageOf.SetFocus
13500       End With
13510     End Select
13520   End If

        ' ** Shift keys.
13530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13550       With Me
13560         intRetVal = 0
13570         .cmbAccounts.SetFocus
13580       End With
13590     End Select
13600   End If

EXITP:
13610   KeyCode = intRetVal
13620   Exit Sub

ERRH:
13630   intRetVal = 0
13640   THAT_PROC = THIS_PROC
13650   That_Erl = Erl
13660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13670   Resume EXITP

End Sub

Private Sub opgOrderBy_optName_KeyDown(KeyCode As Integer, Shift As Integer)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13720   intShiftDown = (Shift And acShiftMask) > 0
13730   intAltDown = (Shift And acAltMask) > 0
13740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13770       With Me
13780         intRetVal = 0
13790         .chkPageOf.SetFocus
13800       End With
13810     End Select
13820   End If

        ' ** Shift keys.
13830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13850       With Me
13860         intRetVal = 0
13870         .cmbAccounts.SetFocus
13880       End With
13890     End Select
13900   End If

EXITP:
13910   KeyCode = intRetVal
13920   Exit Sub

ERRH:
13930   intRetVal = 0
13940   THAT_PROC = THIS_PROC
13950   That_Erl = Erl
13960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13970   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

14010   With Me
14020     Select Case .chkPageOf
          Case True
14030       .chkPageOf_lbl.FontBold = True
14040     Case False
14050       .chkPageOf_lbl.FontBold = False
14060     End Select
14070   End With

EXITP:
14080   Exit Sub

ERRH:
14090   Select Case ERR.Number
        Case Else
14100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14110   End Select
14120   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14220   intShiftDown = (Shift And acShiftMask) > 0
14230   intAltDown = (Shift And acAltMask) > 0
14240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14270       With Me
14280         intRetVal = 0
14290         .cmdPreview.SetFocus
14300       End With
14310     End Select
14320   End If

        ' ** Shift keys.
14330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14350       With Me
14360         intRetVal = 0
14370         .opgOrderBy.SetFocus
14380       End With
14390     End Select
14400   End If

EXITP:
14410   KeyCode = intRetVal
14420   Exit Sub

ERRH:
14430   intRetVal = 0
14440   THAT_PROC = THIS_PROC
14450   That_Erl = Erl
14460   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14470   Resume EXITP

End Sub

Private Sub cmdPreview_Click()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

14510   With Me
14520     If DoReport = True Then  ' ** Function: Below.
14530       Select Case .opgOrderBy
            Case .opgOrderBy_optNumber.OptionValue
14540         strDocName = "rptCapitalGainAndLoss_SortNumber"
14550       Case .opgOrderBy_optName.OptionValue
14560         strDocName = "rptCapitalGainAndLoss_SortName"
14570       End Select
14580       DoCmd.OpenReport strDocName, acViewPreview
14590       DoCmd.Maximize
14600       DoCmd.RunCommand acCmdFitToWindow
14610     End If
14620   End With

EXITP:
14630   Exit Sub

ERRH:
14640   THAT_PROC = THIS_PROC
14650   That_Erl = Erl
14660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14670   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14720   intShiftDown = (Shift And acShiftMask) > 0
14730   intAltDown = (Shift And acAltMask) > 0
14740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14760     Select Case intRetVal
          Case vbKeyTab
14770       With Me
14780         intRetVal = 0
14790         .cmdPrint.SetFocus
14800       End With
14810     End Select
14820   End If

        ' ** Shift keys.
14830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14840     Select Case intRetVal
          Case vbKeyTab
14850       With Me
14860         intRetVal = 0
14870         .cmbAccounts.SetFocus
14880       End With
14890     End Select
14900   End If

        ' ** Ctrl-Shift keys.
14910   If intCtrlDown And (Not intAltDown) And intShiftDown Then
14920     Select Case intRetVal
          Case vbKeyTab
14930       With Me
14940         intRetVal = 0
14950         .chkPageOf.SetFocus
14960       End With
14970     End Select
14980   End If

EXITP:
14990   KeyCode = intRetVal
15000   Exit Sub

ERRH:
15010   intRetVal = 0
15020   THAT_PROC = THIS_PROC
15030   That_Erl = Erl
15040   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15050   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

15110   With Me
15120     If DoReport = True Then  ' ** Function: Below.
15130       Select Case .opgOrderBy
            Case .opgOrderBy_optNumber.OptionValue
15140         strDocName = "rptCapitalGainAndLoss_SortNumber"
15150       Case .opgOrderBy_optName.OptionValue
15160         strDocName = "rptCapitalGainAndLoss_SortName"
15170       End Select
15180       If gblnDev_Debug = True Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
15190         DoCmd.OpenReport strDocName, acViewPreview
15200         DoCmd.Maximize
15210         DoCmd.RunCommand acCmdFitToWindow
15220       Else
              '##GTR_Ref: rptCapitalGainAndLoss_SortNumber
              '##GTR_Ref: rptCapitalGainAndLoss_SortName
15230         DoCmd.OpenReport strDocName, acViewNormal
15240       End If
15250     End If
15260   End With

EXITP:
15270   Exit Sub

ERRH:
15280   THAT_PROC = THIS_PROC
15290   That_Erl = Erl
15300   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15310   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15420   intShiftDown = (Shift And acShiftMask) > 0
15430   intAltDown = (Shift And acAltMask) > 0
15440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15460     Select Case intRetVal
          Case vbKeyTab
15470       With Me
15480         intRetVal = 0
15490         .cmdWord.SetFocus
15500       End With
15510     End Select
15520   End If

        ' ** Shift keys.
15530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15540     Select Case intRetVal
          Case vbKeyTab
15550       With Me
15560         intRetVal = 0
15570         .cmdPreview.SetFocus
15580       End With
15590     End Select
15600   End If

EXITP:
15610   KeyCode = intRetVal
15620   Exit Sub

ERRH:
15630   intRetVal = 0
15640   THAT_PROC = THIS_PROC
15650   That_Erl = Erl
15660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15670   Resume EXITP

End Sub

Private Sub cmdWord_Click()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String, strRptAcct As String
        Dim strRptPath As String, strRptPathFile As String

15710   With Me
15720     If DoReport = True Then  ' ** Function: Below.

15730       If IsNull(.UserReportPath) = True Then
15740         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
15750       Else
15760         strRptPath = .UserReportPath
15770       End If

15780       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
15790         strRptAcct = .cmbAccounts.Column(0)
15800       Case .opgAccountNumber_optAll.OptionValue
15810         strRptAcct = "All"
15820       End Select
15830       strRptCap = "rptCapitalGainAndLoss_" & strRptAcct & "_" & _
              Format(.DateStart, "yymmdd") & "-" & Format(.DateEnd, "yymmdd")

15840       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

15850       If strRptPathFile <> vbNullString Then
15860         Select Case .opgOrderBy
              Case .opgOrderBy_optNumber.OptionValue
15870           strRpt = "rptCapitalGainAndLoss_SortNumber"
15880         Case .opgOrderBy_optName.OptionValue
15890           strRpt = "rptCapitalGainAndLoss_SortName"
15900         End Select
15910         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
15920         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
15930       End If

15940     End If
15950   End With

EXITP:
15960   Exit Sub

ERRH:
15970   Select Case ERR.Number
        Case Else
15980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15990   End Select
16000   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16120   intShiftDown = (Shift And acShiftMask) > 0
16130   intAltDown = (Shift And acAltMask) > 0
16140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16160     Select Case intRetVal
          Case vbKeyTab
16170       With Me
16180         intRetVal = 0
16190         If .cmdExcel.Enabled = True Then
16200           .cmdExcel.SetFocus
16210         Else
16220           .cmdClose.SetFocus
16230         End If
16240       End With
16250     End Select
16260   End If

        ' ** Shift keys.
16270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16280     Select Case intRetVal
          Case vbKeyTab
16290       With Me
16300         intRetVal = 0
16310         .cmdPrint.SetFocus
16320       End With
16330     End Select
16340   End If

EXITP:
16350   KeyCode = intRetVal
16360   Exit Sub

ERRH:
16370   intRetVal = 0
16380   THAT_PROC = THIS_PROC
16390   That_Erl = Erl
16400   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16410   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String, strRptAcct As String
        Dim strRptPath As String, strRptPathFile As String

16510   blnExcel = True
16520   DoCmd.Hourglass True
16530   DoEvents

16540   With Me
16550     If DoReport = True Then  ' ** Function: Below.
16560       If BuildExcel = True Then  ' ** Function: Below.

16570         If IsNull(.UserReportPath) = True Then
16580           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
16590         Else
16600           strRptPath = .UserReportPath
16610         End If

16620         Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
16630           strRptAcct = .cmbAccounts.Column(0)
16640         Case .opgAccountNumber_optAll.OptionValue
16650           strRptAcct = "All"
16660         End Select
16670         strRptCap = "rptCapitalGainAndLoss_" & strRptAcct & "_" & _
                Format(.DateStart, "yymmdd") & "-" & Format(.DateEnd, "yymmdd")

16680         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

16690         If strRptPathFile <> vbNullString Then
                ' ** Sort order handled in BuildExcel(), below.
16700           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
                  ' ** qryCapitalGainsAndLoss_Excel_12 (Union of qryCapitalGainsAndLoss_Excel_10
                  ' ** (qryCapitalGainsAndLoss_Excel_09 (tmpCapitalGainsAndLosses, with add'l fields);
                  ' ** with column headers), qryCapitalGainsAndLoss_Excel_11a (Report title),
                  ' ** qryCapitalGainsAndLoss_Excel_11b (Report period)); For Export,
                  ' ** by specified FormRef('AccountNo').
16710             strQry = "qryCapitalGainsAndLoss_Excel_14"
16720           Case .opgAccountNumber_optAll.OptionValue
                  ' ** qryCapitalGainsAndLoss_Excel_12 (Union of qryCapitalGainsAndLoss_Excel_10
                  ' ** (qryCapitalGainsAndLoss_Excel_09 (tmpCapitalGainsAndLosses, with add'l fields);
                  ' ** with column headers), qryCapitalGainsAndLoss_Excel_11a (Report title),
                  ' ** qryCapitalGainsAndLoss_Excel_11b (Report period)); For Export.
16730             strQry = "qryCapitalGainsAndLoss_Excel_13"
16740           End Select
16750           DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
16760           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
16770         End If

16780       End If
16790     End If
16800   End With

16810   blnExcel = False
16820   DoCmd.Hourglass False

EXITP:
16830   Exit Sub

ERRH:
16840   DoCmd.Hourglass False
16850   blnExcel = False
16860   THAT_PROC = THIS_PROC
16870   That_Erl = Erl
16880   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16890   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16920   intShiftDown = (Shift And acShiftMask) > 0
16930   intAltDown = (Shift And acAltMask) > 0
16940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16960     Select Case intRetVal
          Case vbKeyTab
16970       With Me
16980         intRetVal = 0
16990         .cmdClose.SetFocus
17000       End With
17010     End Select
17020   End If

        ' ** Shift keys.
17030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17040     Select Case intRetVal
          Case vbKeyTab
17050       With Me
17060         intRetVal = 0
17070         .cmdWord.SetFocus
17080       End With
17090     End Select
17100   End If

EXITP:
17110   KeyCode = intRetVal
17120   Exit Sub

ERRH:
17130   intRetVal = 0
17140   THAT_PROC = THIS_PROC
17150   That_Erl = Erl
17160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17170   Resume EXITP

End Sub

Private Function DoReport() As Boolean

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

17210   blnRetVal = True

17220   With Me

17230     If IsNull(.DateStart) = True Or IsNull(.DateEnd) = True Then
17240       blnRetVal = False
17250       MsgBox "You must enter appropriate dates to continue.", vbInformation + vbOKOnly, "Entry Required"
17260     Else
17270       If .cmbAccounts.Enabled = True Then
17280         .cmbAccounts.SetFocus
17290         If .cmbAccounts.text = vbNullString Then
17300           blnRetVal = False
17310           MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
17320         End If
17330       End If
17340     End If

17350     If blnRetVal = True Then

17360       Set dbs = CurrentDb

17370       Select Case .chkShowAllSales
            Case True
17380         Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
17390           Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_02c")
17400         Case .opgAccountNumber_optAll.OptionValue
17410           Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_04c")
17420         End Select
17430       Case False
17440         Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
17450           Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_02a")
17460         Case .opgAccountNumber_optAll.OptionValue
17470           Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_04a")
17480         End Select
17490       End Select

17500       Set rst = qdf.OpenRecordset
17510       With rst
17520         If .BOF = True And .EOF Then
17530           blnRetVal = False
17540           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
17550         End If
17560         .Close
17570       End With
17580       dbs.Close

17590     End If
17600   End With

EXITP:
17610   Set rst = Nothing
17620   Set qdf = Nothing
17630   Set dbs = Nothing
17640   DoReport = blnRetVal
17650   Exit Function

ERRH:
17660   blnRetVal = False
17670   blnExcel = False
17680   Select Case ERR.Number
        Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
17690     THAT_PROC = THIS_PROC
17700     That_Erl = Erl
17710     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17720   Case Else
17730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17740   End Select
17750   Resume EXITP

End Function

Public Function BuildExcel() As Boolean

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "BuildExcel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strLastAccountNo As String, lngLastAccountNo As Long, strLastHolding As String, lngLastHolding As Long
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

17810   blnRetVal = True

17820   With Me

17830     Set dbs = CurrentDb

          ' ** Empty tmpCapitalGainsAndLosses.
17840     Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_Excel_08")
17850     qdf.Execute
17860     Set qdf = Nothing

17870     gstrFormQuerySpec = THIS_NAME

17880     Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
17890       Select Case .chkShowAllSales
            Case True
              ' ** Append qryCapitalGainsAndLoss_Excel_06c (xx) to tmpCapitalGainsAndLosses; for one, sort both, ShowAllSales.
17900         Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_Excel_07a")
17910       Case False
              ' ** Append qryCapitalGainsAndLoss_Excel_06a (xx) to tmpCapitalGainsAndLosses; for one, sort both.
17920         Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_Excel_07")
17930       End Select
17940     Case .opgAccountNumber_optAll.OptionValue
17950       Select Case .chkShowAllSales
            Case True
17960         Select Case .opgOrderBy
              Case .opgOrderBy_optNumber.OptionValue
                ' ** Append qryCapitalGainsAndLoss_Excel_06c (xx) to tmpCapitalGainsAndLosses; for all, sort number, ShowAllSales.
17970           Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_Excel_07k") '_07a]
17980         Case .opgOrderBy_optName.OptionValue
                ' ** Append qryCapitalGainsAndLoss_Excel_06c (xx) to tmpCapitalGainsAndLosses; for all, sort name, ShowAllSales.
17990           Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_Excel_07m") '_07a]
18000         End Select
18010       Case False
18020         Select Case .opgOrderBy
              Case .opgOrderBy_optNumber.OptionValue
                ' ** Append qryCapitalGainsAndLoss_Excel_06a (xx) to tmpCapitalGainsAndLosses; for all, sort number.
18030           Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_Excel_07j") '_07]
18040         Case .opgOrderBy_optName.OptionValue
                ' ** Append qryCapitalGainsAndLoss_Excel_06a (xx) to tmpCapitalGainsAndLosses; for all, sort name.
18050           Set qdf = dbs.QueryDefs("qryCapitalGainsAndLoss_Excel_07l") '_07]
18060         End Select
18070       End Select
18080     End Select
18090     qdf.Execute
18100     Set qdf = Nothing
18110     DoEvents

18120     strLastHolding = vbNullString: lngLastHolding = 0&
18130     strLastAccountNo = vbNullString: lngLastAccountNo = 0&

18140     With dbs
18150       .TableDefs.Refresh

18160       Set rst = .OpenRecordset("tmpCapitalGainsAndLosses", dbOpenDynaset, dbConsistent)
18170       With rst
18180         If .BOF = True And .EOF = True Then
                ' ** No records.
18190           blnRetVal = False
18200         Else
18210           .MoveLast
18220           lngRecs = .RecordCount
18230           .MoveFirst
18240           For lngX = 1& To lngRecs
18250             If ![accountno] <> strLastAccountNo Then
18260               .Edit
18270               ![capgl_par1] = CLng(0)
18280               ![capgl_par2] = CLng(0)
18290               .Update
18300               strLastAccountNo = ![accountno]
18310               lngLastAccountNo = ![capgl_id]
18320               strLastHolding = ![HoldingPeriod]
18330               lngLastHolding = ![capgl_id]
18340             Else
18350               If ![HoldingPeriod] <> strLastHolding Then
18360                 .Edit
18370                 ![capgl_par1] = lngLastAccountNo
18380                 ![capgl_par2] = CLng(0)
18390                 .Update
18400                 strLastHolding = ![HoldingPeriod]
18410                 lngLastHolding = ![capgl_id]
18420               Else
18430                 .Edit
18440                 ![capgl_par1] = lngLastAccountNo
18450                 ![capgl_par2] = lngLastHolding
18460                 .Update
18470               End If
18480             End If
18490             If lngX = 1& Then
18500               .Edit
18510               ![capgl_first] = True
18520               .Update
18530             ElseIf lngX = lngRecs Then
18540               .Edit
18550               ![capgl_last] = True
18560               .Update
18570             End If
18580             DoEvents
18590             If lngX < lngRecs Then .MoveNext
18600           Next
18610         End If
18620         .Close
18630       End With
18640       .Close
18650     End With

18660   End With

EXITP:
18670   Set rst = Nothing
18680   Set qdf = Nothing
18690   Set dbs = Nothing
18700   BuildExcel = blnRetVal
18710   Exit Function

ERRH:
18720   blnRetVal = False
18730   Select Case ERR.Number
        Case Else
18740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18750   End Select
18760   Resume EXITP

End Function
