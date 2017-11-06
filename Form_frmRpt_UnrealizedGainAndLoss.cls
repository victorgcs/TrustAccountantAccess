VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_UnrealizedGainAndLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_UnrealizedGainAndLoss"

'VGC 10/08/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   As Of:            F {DateAsOf}
' **   All Accts:        L {opgAccountNumber_optAll}
' **   Name:             M {opgAccountSource_optName}
' **   Sort By Name:     N {opgOrderBy_optName}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Sort By Number:   R {opgOrderBy_optNumber}
' **   Specified Acct:   U {opgAccountNumber_optSpecified}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

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

160           gstrFormQuerySpec = THIS_NAME

170           .DateAsOf = Null
180           DoEvents

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200           If IsNull(.DateAsOf) = True Then
210             .DateAsOf = TodaysDate  ' ** Module Function: modUtilities.
220           End If

      #If NoExcel Then
230           .cmdExcel.Enabled = False
      #Else
240           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!
              ' ** EVENT CHECK: chkRememberDates!
250           If IsNull(.DateAsOf) = False Then
260             DateAsOf_Exit Cancel  ' ** Procedure: Below.
270           End If

280           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
290           chkRememberMe_AfterUpdate  ' ** Procedure: Below.
300           chkRememberDates_AfterUpdate  ' ** Procedure: Below.
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

          ' ** Create an instance of the Calendar class.
620       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
630       clsMonthClass.hWndForm = Me.hwnd
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

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       Case vbKeyF7
900         With Me
910           intRetVal = 0
920           .cmdCalendar.SetFocus
930           cmdCalendar_Click  ' ** Procedure: Below.
940         End With
950       End Select
960     End If

        ' ** Alt keys.
970     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
980       Select Case intRetVal
          Case vbKeyX
990         intRetVal = 0
1000        cmdClose_Click  ' ** Procedure: Below.
1010      End Select
1020    End If

        ' ** Ctrl keys.
1030    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1040      Select Case intRetVal
          Case vbKeyG
1050        With Me
1060          intRetVal = 0
1070          If .chkPageOf.Enabled = True Then
1080            .chkPageOf.SetFocus
1090            .chkPageOf = (Not .chkPageOf)
1100            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1110          End If
1120        End With
1130      End Select
1140    End If

        ' ** Ctrl-Shift keys.
1150    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1160      Select Case intRetVal
          Case vbKeyF
1170        With Me
1180          intRetVal = 0
1190          .FocusHolder.SetFocus
1200          .DateAsOf = Null
1210        End With
1220      End Select
1230    End If

EXITP:
1240    KeyCode = intRetVal
1250    Exit Sub

ERRH:
1260    intRetVal = 0
1270    THAT_PROC = THIS_PROC
1280    That_Erl = Erl
1290    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1300    Resume EXITP

End Sub

Private Sub Form_Timer()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1410    With Me
1420      .TimerInterval = 0&
1430      If gblnSetFocus = True Then
1440        gblnSetFocus = False
1450        strFormName = vbNullString
1460  On Error Resume Next
1470        Do Until strFormName = THIS_NAME
1480          DoCmd.SelectObject acForm, THIS_NAME, False
1490          strFormName = Screen.ActiveForm.Name
1500        Loop
1510  On Error GoTo ERRH
1520      End If
1530      If gblnGoToReport = True Then
1540        If .cmdPrint.Enabled = True Then
1550          .cmdPrint.SetFocus
1560        Else
1570          .cmdClose.SetFocus
1580        End If
1590        gblnGoToReport = False
1600        Beep
1610        DoCmd.Hourglass False
1620        DoEvents
1630      Else
1640        DoCmd.Hourglass False
1650      End If
1660    End With

EXITP:
1670    Exit Sub

ERRH:
1680    DoCmd.Hourglass False
1690    Select Case ERR.Number
        Case Else
1700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1710    End Select
1720    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1810    Select Case DataErr
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
1820    Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1830      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateAsOf.
1840        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1850        Response = acDataErrContinue
1860      End If
1870    Case 2237  ' ** The text you entered isn't an item in the list.
1880      Response = acDataErrDisplay
1890    Case Else
1900      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1910      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1920    End Select

EXITP:
1930    Exit Sub

ERRH:
1940    Select Case ERR.Number
        Case Else
1950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1960    End Select
1970    Resume EXITP

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
2080      If .chkRememberMe = False Then
2090        .cmbAccounts = Null
2100        DoEvents
2110      End If
2120    End With

2130    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2140    Do While Reports.Count > 0
2150      DoCmd.Close acReport, Reports(0).Name
2160    Loop

2170    gstrFormQuerySpec = vbNullString
2180    gblnSetFocus = True
2190    gstrReturningForm = THIS_NAME
2200    Select Case strCallingForm
        Case "frmMenu_Report"
2210      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2220    End Select

EXITP:
2230    Exit Sub

ERRH:
2240    Select Case ERR.Number
        Case Else
2250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2260    End Select
2270    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

2310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

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

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2810    DoCmd.Hourglass True
2820    DoEvents

2830    DoCmd.Close acForm, THIS_NAME

EXITP:
2840    Exit Sub

ERRH:
2850    Select Case ERR.Number
        Case Else
2860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2870    End Select
2880    Resume EXITP

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
2990          .DateAsOf.SetFocus
3000        End With
3010      End Select
3020    End If

        ' ** Shift keys.
3030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3040      Select Case intRetVal
          Case vbKeyTab
3050        With Me
3060          intRetVal = 0
3070          If .cmdExcel.Enabled = True Then
3080            .cmdExcel.SetFocus
3090          Else
3100            .cmdWord.SetFocus
3110          End If
3120        End With
3130      End Select
3140    End If

EXITP:
3150    KeyCode = intRetVal
3160    Exit Sub

ERRH:
3170    intRetVal = 0
3180    THAT_PROC = THIS_PROC
3190    That_Erl = Erl
3200    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3210    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

3310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub DateAsOf_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
3470        With Me
3480          intRetVal = 0
3490          .opgAccountNumber.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3550        With Me
3560          intRetVal = 0
3570          .cmdClose.SetFocus
3580        End With
3590      End Select
3600    End If

        ' ** Ctrl keys.
3610    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3620      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3630        With Me
3640          intRetVal = 0
3650          .cmdCalendar.SetFocus
3660        End With
3670      End Select
3680    End If

EXITP:
3690    KeyCode = intRetVal
3700    Exit Sub

ERRH:
3710    intRetVal = 0
3720    THAT_PROC = THIS_PROC
3730    That_Erl = Erl
3740    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3750    Resume EXITP

End Sub

Private Sub DateAsOf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_MouseMove"

3810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub DateAsOf_Exit(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_Exit"

3910    With Me
3920      If IsNull(.DateAsOf) = True Then
            ' ** Populate it with today's date.
3930        .DateAsOf = Date
3940      Else
3950        If Trim(.DateAsOf) = vbNullString Then
              ' ** Populate it with today's date.
3960          .DateAsOf = Date
3970        Else
3980          If .DateAsOf < Date Then
3990            MsgBox "Data is not rolled back to the entered date." & vbCrLf & _
                  "The report reflects current information.", vbInformation + vbOKOnly, "Current Data Only"
4000          ElseIf .DateAsOf > Date Then
4010            MsgBox "Report will reflect current information.", vbInformation + vbOKOnly, "Current Data Only"
4020          End If
4030        End If
4040      End If
4050    End With

EXITP:
4060    Exit Sub

ERRH:
4070    THAT_PROC = THIS_PROC
4080    That_Erl = Erl
4090    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4100    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

4210    With Me
4220      blnCalendar1_Focus = True
4230      .cmdCalendar_raised_semifocus_dots_img.Visible = True
4240      .cmdCalendar_raised_img.Visible = False
4250      .cmdCalendar_raised_focus_img.Visible = False
4260      .cmdCalendar_raised_focus_dots_img.Visible = False
4270      .cmdCalendar_sunken_focus_dots_img.Visible = False
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

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

4410    With Me
4420      blnCalendar1_MouseDown = True
4430      .cmdCalendar_sunken_focus_dots_img.Visible = True
4440      .cmdCalendar_raised_img.Visible = False
4450      .cmdCalendar_raised_semifocus_dots_img.Visible = False
4460      .cmdCalendar_raised_focus_img.Visible = False
4470      .cmdCalendar_raised_focus_dots_img.Visible = False
4480      .cmdCalendar_raised_img_dis.Visible = False
4490    End With

EXITP:
4500    Exit Sub

ERRH:
4510    Select Case ERR.Number
        Case Else
4520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4530    End Select
4540    Resume EXITP

End Sub

Private Sub cmdCalendar_Click()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

4610    With Me
4620      datStartDate = Date
4630      datEndDate = 0
4640      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
4650      If blnRetVal = True Then
4660        .DateAsOf = datStartDate
4670      Else
4680        .DateAsOf = CDate(Format(Date, "mm/dd/yyyy"))
4690      End If
4700      .DateAsOf.SetFocus
4710    End With

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4820    intShiftDown = (Shift And acShiftMask) > 0
4830    intAltDown = (Shift And acAltMask) > 0
4840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4860      Select Case intRetVal
          Case vbKeyTab
4870        With Me
4880          intRetVal = 0
4890          .chkRememberDates.SetFocus
4900        End With
4910      End Select
4920    End If

        ' ** Shift keys.
4930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4940      Select Case intRetVal
          Case vbKeyTab
4950        With Me
4960          intRetVal = 0
4970          .DateAsOf.SetFocus
4980        End With
4990      End Select
5000    End If

EXITP:
5010    KeyCode = intRetVal
5020    Exit Sub

ERRH:
5030    intRetVal = 0
5040    Select Case ERR.Number
        Case Else
5050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5060    End Select
5070    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

5110    With Me
5120      If blnCalendar1_MouseDown = False Then
5130        Select Case blnCalendar1_Focus
            Case True
5140          .cmdCalendar_raised_focus_dots_img.Visible = True
5150          .cmdCalendar_raised_focus_img.Visible = False
5160        Case False
5170          .cmdCalendar_raised_focus_img.Visible = True
5180          .cmdCalendar_raised_focus_dots_img.Visible = False
5190        End Select
5200        .cmdCalendar_raised_img.Visible = False
5210        .cmdCalendar_raised_semifocus_dots_img.Visible = False
5220        .cmdCalendar_sunken_focus_dots_img.Visible = False
5230        .cmdCalendar_raised_img_dis.Visible = False
5240      End If
5250    End With

EXITP:
5260    Exit Sub

ERRH:
5270    Select Case ERR.Number
        Case Else
5280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5290    End Select
5300    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

5410    With Me
5420      .cmdCalendar_raised_focus_dots_img.Visible = True
5430      .cmdCalendar_raised_img.Visible = False
5440      .cmdCalendar_raised_semifocus_dots_img.Visible = False
5450      .cmdCalendar_raised_focus_img.Visible = False
5460      .cmdCalendar_sunken_focus_dots_img.Visible = False
5470      .cmdCalendar_raised_img_dis.Visible = False
5480      blnCalendar1_MouseDown = False
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

Private Sub cmdCalendar_LostFocus()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

5610    With Me
5620      .cmdCalendar_raised_img.Visible = True
5630      .cmdCalendar_raised_semifocus_dots_img.Visible = False
5640      .cmdCalendar_raised_focus_img.Visible = False
5650      .cmdCalendar_raised_focus_dots_img.Visible = False
5660      .cmdCalendar_sunken_focus_dots_img.Visible = False
5670      .cmdCalendar_raised_img_dis.Visible = False
5680      blnCalendar1_Focus = False
5690    End With

EXITP:
5700    Exit Sub

ERRH:
5710    Select Case ERR.Number
        Case Else
5720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730    End Select
5740    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

5810    With Me
5820      Select Case .chkRememberDates
          Case True
5830        .chkRememberDates_lbl.FontBold = True
5840      Case False
5850        .chkRememberDates_lbl.FontBold = False
5860      End Select
5870    End With

EXITP:
5880    Exit Sub

ERRH:
5890    Select Case ERR.Number
        Case Else
5900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5910    End Select
5920    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6020    intShiftDown = (Shift And acShiftMask) > 0
6030    intAltDown = (Shift And acAltMask) > 0
6040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6070        With Me
6080          intRetVal = 0
6090          .opgAccountNumber.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6150        With Me
6160          intRetVal = 0
6170          .cmdCalendar.SetFocus
6180        End With
6190      End Select
6200    End If

EXITP:
6210    KeyCode = intRetVal
6220    Exit Sub

ERRH:
6230    intRetVal = 0
6240    THAT_PROC = THIS_PROC
6250    That_Erl = Erl
6260    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6270    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

6310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub opgAccountNumber_AfterUpdate()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

6410    With Me
6420      Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
6430        .opgAccountNumber_optSpecified_lbl.FontBold = True
6440        .opgAccountNumber_optAll_lbl.FontBold = False
6450        .opgAccountNumber_optSpecified_lbl_box.Visible = True
6460        .opgAccountNumber_optAll_lbl_box.Visible = False
6470        .cmbAccounts.Enabled = True
6480        .cmbAccounts.BorderColor = CLR_LTBLU2
6490        .cmbAccounts.BackStyle = acBackStyleNormal
6500        .chkRememberMe.Enabled = True
6510        .chkRememberMe_lbl.Visible = True
6520        .chkRememberMe_lbl2_dim.Visible = False
6530        .chkRememberMe_lbl2_dim_hi.Visible = False
6540        .opgAccountSource.Enabled = True
6550        .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
6560        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
6570        .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
6580        .opgAccountSource_optName_lbl2_dim_hi.Visible = False
6590      Case .opgAccountNumber_optAll.OptionValue
6600        .opgAccountNumber_optSpecified_lbl.FontBold = False
6610        .opgAccountNumber_optAll_lbl.FontBold = True
6620        .opgAccountNumber_optSpecified_lbl_box.Visible = False
6630        .opgAccountNumber_optAll_lbl_box.Visible = True
6640        .cmbAccounts.Enabled = False
6650        .cmbAccounts.BorderColor = WIN_CLR_DISR
6660        .cmbAccounts.BackStyle = acBackStyleTransparent
6670        .chkRememberMe.Enabled = False
6680        .chkRememberMe_lbl.Visible = False
6690        .chkRememberMe_lbl2_dim.Visible = True
6700        .chkRememberMe_lbl2_dim_hi.Visible = True
6710        .opgAccountSource.Enabled = False
6720        .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
6730        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
6740        .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
6750        .opgAccountSource_optName_lbl2_dim_hi.Visible = True
6760      End Select
6770    End With

EXITP:
6780    Exit Sub

ERRH:
6790    Select Case ERR.Number
        Case Else
6800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6810    End Select
6820    Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6920    intShiftDown = (Shift And acShiftMask) > 0
6930    intAltDown = (Shift And acAltMask) > 0
6940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6970        With Me
6980          intRetVal = 0
6990          If .cmbAccounts.Enabled = True Then
7000            .cmbAccounts.SetFocus
7010          Else
7020            .cmdPreview.SetFocus
7030          End If
7040        End With
7050      End Select
7060    End If

        ' ** Shift keys.
7070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7080      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7090        With Me
7100          intRetVal = 0
7110          .DateAsOf.SetFocus
7120        End With
7130      End Select
7140    End If

        ' ** Ctrl keys.
7150    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7170        With Me
7180          intRetVal = 0
7190          If .cmbAccounts.Enabled = True Then
7200            .opgAccountSource.SetFocus
7210          Else
7220            .chkPageOf.SetFocus
7230          End If
7240        End With
7250      End Select
7260    End If

        ' ** Ctrl-Shift keys.
7270    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7280      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7290        With Me
7300          intRetVal = 0
7310          .chkRememberDates.SetFocus
7320        End With
7330      End Select
7340    End If

EXITP:
7350    KeyCode = intRetVal
7360    Exit Sub

ERRH:
7370    intRetVal = 0
7380    THAT_PROC = THIS_PROC
7390    That_Erl = Erl
7400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7410    Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7520    intShiftDown = (Shift And acShiftMask) > 0
7530    intAltDown = (Shift And acAltMask) > 0
7540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7570        With Me
7580          intRetVal = 0
7590          If .cmbAccounts.Enabled = True Then
7600            .cmbAccounts.SetFocus
7610          Else
7620            .cmdPreview.SetFocus
7630          End If
7640        End With
7650      End Select
7660    End If

        ' ** Shift keys.
7670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7680      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7690        With Me
7700          intRetVal = 0
7710          .DateAsOf.SetFocus
7720        End With
7730      End Select
7740    End If

        ' ** Ctrl keys.
7750    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7770        With Me
7780          intRetVal = 0
7790          If .cmbAccounts.Enabled = True Then
7800            .opgAccountSource.SetFocus
7810          Else
7820            .chkPageOf.SetFocus
7830          End If
7840        End With
7850      End Select
7860    End If

        ' ** Ctrl-Shift keys.
7870    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7880      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7890        With Me
7900          intRetVal = 0
7910          .chkRememberDates.SetFocus
7920        End With
7930      End Select
7940    End If

EXITP:
7950    KeyCode = intRetVal
7960    Exit Sub

ERRH:
7970    intRetVal = 0
7980    THAT_PROC = THIS_PROC
7990    That_Erl = Erl
8000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8010    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

8110    strAccountNo = vbNullString

8120    With Me
8130      If IsNull(.cmbAccounts) = False Then
8140        If Len(.cmbAccounts.Column(0)) > 0 Then
8150          strAccountNo = .cmbAccounts.Column(0)
8160        End If
8170      End If
8180      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
8190        .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
8200        .opgAccountSource_optNumber_lbl2.FontBold = True
8210        .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
8220        .opgAccountSource_optName_lbl2.FontBold = False
8230        .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
8240      Case .opgAccountSource_optName.OptionValue
8250        .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
8260        .opgAccountSource_optNumber_lbl2.FontBold = False
8270        .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
8280        .opgAccountSource_optName_lbl2.FontBold = True
8290        .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
8300      End Select
8310      DoEvents
8320      If strAccountNo <> vbNullString Then
8330        .cmbAccounts = strAccountNo
8340      End If
8350    End With

EXITP:
8360    Exit Sub

ERRH:
8370    Select Case ERR.Number
        Case Else
8380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8390    End Select
8400    Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          intRetVal = 0
8590          .chkRememberMe.SetFocus
8600        End With
8610      End Select
8620    End If

        ' ** Shift keys.
8630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8650        With Me
8660          intRetVal = 0
8670          .opgAccountNumber.SetFocus
8680        End With
8690      End Select
8700    End If

EXITP:
8710    KeyCode = intRetVal
8720    Exit Sub

ERRH:
8730    intRetVal = 0
8740    THAT_PROC = THIS_PROC
8750    That_Erl = Erl
8760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8770    Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8820    intShiftDown = (Shift And acShiftMask) > 0
8830    intAltDown = (Shift And acAltMask) > 0
8840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8870        With Me
8880          intRetVal = 0
8890          .chkRememberMe.SetFocus
8900        End With
8910      End Select
8920    End If

        ' ** Shift keys.
8930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8950        With Me
8960          intRetVal = 0
8970          .opgAccountNumber.SetFocus
8980        End With
8990      End Select
9000    End If

EXITP:
9010    KeyCode = intRetVal
9020    Exit Sub

ERRH:
9030    intRetVal = 0
9040    THAT_PROC = THIS_PROC
9050    That_Erl = Erl
9060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9070    Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

9110    With Me
9120      Select Case .chkRememberMe
          Case True
9130        .chkRememberMe_lbl.FontBold = True
9140        .chkRememberMe_lbl2_dim.FontBold = True
9150        .chkRememberMe_lbl2_dim_hi.FontBold = True
9160      Case False
9170        .chkRememberMe_lbl.FontBold = False
9180        .chkRememberMe_lbl2_dim.FontBold = False
9190        .chkRememberMe_lbl2_dim_hi.FontBold = False
9200      End Select
9210    End With

EXITP:
9220    Exit Sub

ERRH:
9230    Select Case ERR.Number
        Case Else
9240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9250    End Select
9260    Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9320    intShiftDown = (Shift And acShiftMask) > 0
9330    intAltDown = (Shift And acAltMask) > 0
9340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9370        With Me
9380          intRetVal = 0
9390          .cmbAccounts.SetFocus
9400        End With
9410      End Select
9420    End If

        ' ** Shift keys.
9430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9450        With Me
9460          intRetVal = 0
9470          .opgAccountSource.SetFocus
9480        End With
9490      End Select
9500    End If

EXITP:
9510    KeyCode = intRetVal
9520    Exit Sub

ERRH:
9530    intRetVal = 0
9540    THAT_PROC = THIS_PROC
9550    That_Erl = Erl
9560    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9570    Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

9610    With Me
9620      If .cmbAccounts.text = vbNullString Then
9630        ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
9640        .cmbAccounts = .cmbAccounts.ItemData(0)
9650      End If
9660    End With

EXITP:
9670    Exit Sub

ERRH:
9680    Select Case ERR.Number
        Case Else
9690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9700    End Select
9710    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

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
9890          .cmdPreview.SetFocus
9900        End With
9910      End Select
9920    End If

        ' ** Shift keys.
9930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9950        With Me
9960          intRetVal = 0
9970          .opgAccountNumber.SetFocus
9980        End With
9990      End Select
10000   End If

        ' ** Ctrl keys.
10010   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10030       With Me
10040         intRetVal = 0
10050         opgOrderBy.SetFocus
10060       End With
10070     End Select
10080   End If

        ' ** Ctrl-Shift keys.
10090   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10110       With Me
10120         intRetVal = 0
10130         .chkRememberMe.SetFocus
10140       End With
10150     End Select
10160   End If

EXITP:
10170   KeyCode = intRetVal
10180   Exit Sub

ERRH:
10190   intRetVal = 0
10200   THAT_PROC = THIS_PROC
10210   That_Erl = Erl
10220   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10230   Resume EXITP

End Sub

Private Sub opgOrderBy_AfterUpdate()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_AfterUpdate"

10310   With Me
10320     Select Case .opgOrderBy
          Case .opgOrderBy_optNumber.OptionValue
10330       .opgOrderBy_optNumber_lbl.FontBold = True
10340       .opgOrderBy_optName_lbl.FontBold = False
10350     Case .opgOrderBy_optName.OptionValue
10360       .opgOrderBy_optNumber_lbl.FontBold = False
10370       .opgOrderBy_optName_lbl.FontBold = True
10380     End Select
10390   End With

EXITP:
10400   Exit Sub

ERRH:
10410   Select Case ERR.Number
        Case Else
10420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10430   End Select
10440   Resume EXITP

End Sub

Private Sub opgOrderBy_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optNumber_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
10570       With Me
10580         intRetVal = 0
10590         .chkPageOf.SetFocus
10600       End With
10610     End Select
10620   End If

        ' ** Shift keys.
10630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10650       With Me
10660         intRetVal = 0
10670         .cmbAccounts.SetFocus
10680       End With
10690     End Select
10700   End If

EXITP:
10710   KeyCode = intRetVal
10720   Exit Sub

ERRH:
10730   intRetVal = 0
10740   THAT_PROC = THIS_PROC
10750   That_Erl = Erl
10760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10770   Resume EXITP

End Sub

Private Sub opgOrderBy_optName_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10820   intShiftDown = (Shift And acShiftMask) > 0
10830   intAltDown = (Shift And acAltMask) > 0
10840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
10890         .chkPageOf.SetFocus
10900       End With
10910     End Select
10920   End If

        ' ** Shift keys.
10930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10950       With Me
10960         intRetVal = 0
10970         .cmbAccounts.SetFocus
10980       End With
10990     End Select
11000   End If

EXITP:
11010   KeyCode = intRetVal
11020   Exit Sub

ERRH:
11030   intRetVal = 0
11040   THAT_PROC = THIS_PROC
11050   That_Erl = Erl
11060   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11070   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

11110   With Me
11120     Select Case .chkPageOf
          Case True
11130       .chkPageOf_lbl.FontBold = True
11140     Case False
11150       .chkPageOf_lbl.FontBold = False
11160     End Select
11170   End With

EXITP:
11180   Exit Sub

ERRH:
11190   DoCmd.Hourglass False
11200   Select Case ERR.Number
        Case Else
11210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11220   End Select
11230   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11320   intShiftDown = (Shift And acShiftMask) > 0
11330   intAltDown = (Shift And acAltMask) > 0
11340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11370       With Me
11380         intRetVal = 0
11390         .cmdPreview.SetFocus
11400       End With
11410     End Select
11420   End If

        ' ** Shift keys.
11430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11450       With Me
11460         intRetVal = 0
11470         .opgOrderBy.SetFocus
11480       End With
11490     End Select
11500   End If

EXITP:
11510   KeyCode = intRetVal
11520   Exit Sub

ERRH:
11530   intRetVal = 0
11540   THAT_PROC = THIS_PROC
11550   That_Erl = Erl
11560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11570   Resume EXITP

End Sub

Private Sub cmdPreview_Click()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

11610   With Me
11620     gdatStartDate = 0
11630     If DoReport = True Then  ' ** Function: Below.
11640       Select Case .opgOrderBy
            Case .opgOrderBy_optNumber.OptionValue
11650         strDocName = "rptUnrealizedGainAndLoss_SortNumber"
11660       Case .opgOrderBy_optName.OptionValue
11670         strDocName = "rptUnrealizedGainAndLoss_SortName"
11680       End Select
11690       DoCmd.OpenReport strDocName, acViewPreview
11700       DoCmd.Maximize
11710       DoCmd.RunCommand acCmdFitToWindow
11720     End If
11730   End With

EXITP:
11740   Exit Sub

ERRH:
11750   Select Case ERR.Number
        Case Else
11760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11770   End Select
11780   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
11890         .cmdPrint.SetFocus
11900       End With
11910     End Select
11920   End If

        ' ** Shift keys.
11930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11940     Select Case intRetVal
          Case vbKeyTab
11950       With Me
11960         intRetVal = 0
11970         .cmbAccounts.SetFocus
11980       End With
11990     End Select
12000   End If

        ' ** Ctrl-Shift keys.
12010   If intCtrlDown And (Not intAltDown) And intShiftDown Then
12020     Select Case intRetVal
          Case vbKeyTab
12030       With Me
12040         intRetVal = 0
12050         .chkPageOf.SetFocus
12060       End With
12070     End Select
12080   End If

EXITP:
12090   KeyCode = intRetVal
12100   Exit Sub

ERRH:
12110   intRetVal = 0
12120   THAT_PROC = THIS_PROC
12130   That_Erl = Erl
12140   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12150   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

12210   With Me
12220     gdatStartDate = 0
12230     If DoReport = True Then  ' ** Function: Below.
12240       Select Case .opgOrderBy
            Case .opgOrderBy_optNumber.OptionValue
12250         strDocName = "rptUnrealizedGainAndLoss_SortNumber"
12260       Case .opgOrderBy_optName.OptionValue
12270         strDocName = "rptUnrealizedGainAndLoss_SortName"
12280       End Select
12290       If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
12300         DoCmd.OpenReport strDocName, acViewPreview
12310         DoCmd.Maximize
12320         DoCmd.RunCommand acCmdFitToWindow
12330       Else
              '##GTR_Ref: rptUnrealizedGainAndLoss_SortNumber
              '##GTR_Ref: rptUnrealizedGainAndLoss_SortName
12340         DoCmd.OpenReport strDocName, acViewNormal
12350       End If
12360     End If
12370   End With

EXITP:
12380   Exit Sub

ERRH:
12390   THAT_PROC = THIS_PROC
12400   That_Erl = Erl
12410   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12420   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12520   intShiftDown = (Shift And acShiftMask) > 0
12530   intAltDown = (Shift And acAltMask) > 0
12540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12560     Select Case intRetVal
          Case vbKeyTab
12570       With Me
12580         intRetVal = 0
12590         .cmdWord.SetFocus
12600       End With
12610     End Select
12620   End If

        ' ** Shift keys.
12630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12640     Select Case intRetVal
          Case vbKeyTab
12650       With Me
12660         intRetVal = 0
12670         .cmdPreview.SetFocus
12680       End With
12690     End Select
12700   End If

EXITP:
12710   KeyCode = intRetVal
12720   Exit Sub

ERRH:
12730   intRetVal = 0
12740   THAT_PROC = THIS_PROC
12750   That_Erl = Erl
12760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12770   Resume EXITP

End Sub

Private Sub cmdWord_Click()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String, strRptAcct As String
        Dim strRptPath As String, strRptPathFile As String

12810   With Me
12820     If DoReport = True Then  ' ** Procedure: Below.

12830       If IsNull(.UserReportPath) = True Then
12840         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
12850       Else
12860         strRptPath = .UserReportPath
12870       End If

12880       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
12890         Select Case .opgOrderBy
              Case .opgOrderBy_optNumber.OptionValue
12900           strRpt = "rptUnrealizedGainAndLoss_SortNumber"
12910         Case .opgOrderBy_optName.OptionValue
12920           strRpt = "rptUnrealizedGainAndLoss_SortName"
12930         End Select
12940         strRptAcct = .cmbAccounts
12950         strRptCap = "Unrealized_Gains_and_Losses_" & strRptAcct & "_" & Format(.DateAsOf, "yymmdd")
12960       Case .opgAccountNumber_optAll.OptionValue
12970         Select Case .opgOrderBy
              Case .opgOrderBy_optNumber.OptionValue
12980           strRpt = "rptUnrealizedGainAndLoss_SortNumber"
12990         Case .opgOrderBy_optName.OptionValue
13000           strRpt = "rptUnrealizedGainAndLoss_SortName"
13010         End Select
13020         strRptCap = "Unrealized_Gains_and_Losses_All_" & Format(.DateAsOf, "yymmdd")
13030       End Select

13040       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

13050       If strRptPathFile <> vbNullString Then
13060         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
13070         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
13080       End If

13090     End If
13100   End With

EXITP:
13110   Exit Sub

ERRH:
13120   Select Case ERR.Number
        Case Else
13130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13140   End Select
13150   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13220   intShiftDown = (Shift And acShiftMask) > 0
13230   intAltDown = (Shift And acAltMask) > 0
13240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13260     Select Case intRetVal
          Case vbKeyTab
13270       With Me
13280         intRetVal = 0
13290         If .cmdExcel.Enabled = True Then
13300           .cmdExcel.SetFocus
13310         Else
13320           .cmdClose.SetFocus
13330         End If
13340       End With
13350     End Select
13360   End If

        ' ** Shift keys.
13370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13380     Select Case intRetVal
          Case vbKeyTab
13390       With Me
13400         intRetVal = 0
13410         .cmdPrint.SetFocus
13420       End With
13430     End Select
13440   End If

EXITP:
13450   KeyCode = intRetVal
13460   Exit Sub

ERRH:
13470   intRetVal = 0
13480   THAT_PROC = THIS_PROC
13490   That_Erl = Erl
13500   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13510   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQry As String, strRptCap As String, strRptAcct As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

13610   With Me

13620     DoCmd.Hourglass True
13630     DoEvents

13640     blnContinue = True

13650     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13660       DoCmd.Hourglass False
13670       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
13680       If msgResponse <> vbRetry Then
13690         blnContinue = False
13700       End If
13710     End If

13720     If blnContinue = True Then

13730       DoCmd.Hourglass True
13740       DoEvents

13750       If DoReport = True Then  ' ** Procedure: Below.

13760         blnContinue = True

13770         If .opgAccountNumber = .opgAccountNumber_optAll.OptionValue Then
13780           Set dbs = CurrentDb
                ' ** Empty tmpUnrealizedGainAndLoss.
13790           Set qdf = dbs.QueryDefs("qryRpt_UnrealizedGainAndLoss_49c")
13800           qdf.Execute
13810           Set qdf = Nothing
13820           DoEvents
                ' ** Append qryRpt_UnrealizedGainAndLoss_49a (xx) to tmpUnrealizedGainAndLoss.
13830           Set qdf = dbs.QueryDefs("qryRpt_UnrealizedGainAndLoss_49b")
13840           qdf.Execute
13850           Set qdf = Nothing
13860           DoEvents
13870           dbs.Close
13880           Set dbs = Nothing
13890         End If
13900         DoEvents

13910         If IsNull(.UserReportPath) = True Then
13920           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
13930         Else
13940           strRptPath = .UserReportPath
13950         End If

13960         Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
13970           Select Case .opgOrderBy
                Case .opgOrderBy_optNumber.OptionValue
                  ' ** qryRpt_UnrealizedGainAndLoss_29 (xx), For Export, one account, sort number.
13980             strQry = "qryRpt_UnrealizedGainAndLoss_30_01"  ' ** These 2 queries are identical.
13990           Case .opgOrderBy_optName.OptionValue
                  ' ** qryRpt_UnrealizedGainAndLoss_29 (xx), For Export, one account, sort name.
14000             strQry = "qryRpt_UnrealizedGainAndLoss_30_02"
14010           End Select
14020           strRptAcct = .cmbAccounts
14030           strRptCap = "Unrealized_Gains_and_Losses_" & strRptAcct & "_" & Format(.DateAsOf, "yymmdd")
14040         Case .opgAccountNumber_optAll.OptionValue
14050           Select Case .opgOrderBy
                Case .opgOrderBy_optNumber.OptionValue
                  ' ** qryRpt_UnrealizedGainAndLoss_50 (xx), For Export, all accounts, sort number.
14060             strQry = "qryRpt_UnrealizedGainAndLoss_51_01"
14070           Case .opgOrderBy_optName.OptionValue
                  ' ** qryRpt_UnrealizedGainAndLoss_50 (xx), For Export, all accounts, sort name.
14080             strQry = "qryRpt_UnrealizedGainAndLoss_51_02"
14090           End Select
14100           strRptAcct = vbNullString
14110           strRptCap = "Unrealized_Gains_and_Losses_All_" & Format(.DateAsOf, "yymmdd")
14120         End Select

              ' ** Short Term
              ' ** Mid Term
              ' ** Long Term

14130         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

14140         If strRptPathFile <> vbNullString Then
14150           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
14160             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
14170           End If
14180           DoEvents
14190           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
14200             Kill strRptPathFile
14210           End If
14220           DoEvents
14230           gstrFormQuerySpec = THIS_NAME
14240           DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
14250           If Excel_NameOnly(strRptPathFile, "Unrealized") = True Then  ' ** Module Function: modExcelFuncs.
14260             If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
14270               EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
14280             End If
14290             OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
14300           End If
14310           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
14320         End If

14330       End If  ' ** DoReport().
14340     End If  ' ** blnContinue.

14350     DoCmd.Hourglass False

14360   End With

EXITP:
14370   Set qdf = Nothing
14380   Set dbs = Nothing
14390   Exit Sub

ERRH:
14400   DoCmd.Hourglass False
14410   Select Case ERR.Number
        Case Else
14420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14430   End Select
14440   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14520   intShiftDown = (Shift And acShiftMask) > 0
14530   intAltDown = (Shift And acAltMask) > 0
14540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14560     Select Case intRetVal
          Case vbKeyTab
14570       With Me
14580         intRetVal = 0
14590         .cmdClose.SetFocus
14600       End With
14610     End Select
14620   End If

        ' ** Shift keys.
14630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14640     Select Case intRetVal
          Case vbKeyTab
14650       With Me
14660         intRetVal = 0
14670         .cmdWord.SetFocus
14680       End With
14690     End Select
14700   End If

EXITP:
14710   KeyCode = intRetVal
14720   Exit Sub

ERRH:
14730   intRetVal = 0
14740   THAT_PROC = THIS_PROC
14750   That_Erl = Erl
14760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14770   Resume EXITP

End Sub

Private Function DoReport() As Boolean

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

14810   blnRetVal = True

14820   With Me

14830     If IsNull(.DateAsOf) = True Then
14840       blnRetVal = False
14850       MsgBox "You must enter appropriate dates to continue.", vbInformation + vbOKOnly, "Entry Required"
14860     Else
14870       If .cmbAccounts.Enabled = True Then
14880         .cmbAccounts.SetFocus
14890         If .cmbAccounts.text = vbNullString Then
14900           blnRetVal = False
14910           MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
14920         End If
14930       End If
14940     End If

14950     If blnRetVal = True Then

14960       Set dbs = CurrentDb

14970       Select Case .opgAccountNumber.Value
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** ActiveAssets, with add'l fields, by specified FormRef('DateAsOf'), FormRef('DateAsOfTime'); one account.
14980         Set qdf = dbs.QueryDefs("qryRpt_UnrealizedGainAndLoss_01")
14990       Case .opgAccountNumber_optAll.OptionValue
              ' ** ActiveAssets, with add'l fields, by specified FormRef('DateAsOf'), FormRef('DateAsOfTime'); all accounts.
15000         Set qdf = dbs.QueryDefs("qryRpt_UnrealizedGainAndLoss_02")
15010       End Select

15020       Set rst = qdf.OpenRecordset
15030       With rst
15040         If .BOF = True And .EOF = True Then
15050           blnRetVal = False
15060           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
15070         Else
15080           .MoveFirst
15090           gdatStartDate = ![currentDate]  ' ** Used by the report.
15100         End If
15110         .Close
15120       End With

15130       dbs.Close

15140     End If

15150   End With

EXITP:
15160   Set rst = Nothing
15170   Set qdf = Nothing
15180   Set dbs = Nothing
15190   DoReport = blnRetVal
15200   Exit Function

ERRH:
15210   blnRetVal = False
15220   Select Case ERR.Number
        Case Else
15230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15240   End Select
15250   Resume EXITP

End Function
