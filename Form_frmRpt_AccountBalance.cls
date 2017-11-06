VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_AccountBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_AccountBalance"

'VGC 07/16/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   As Of:            F {DateAsOf}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Focus Holder:     F {FocusHolder}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False

160           .DateAsOf = Null
170           DoEvents

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
190           .cmdExcel.Enabled = False
      #Else
200           .cmdExcel.Enabled = True
      #End If

210           chkRememberDates_AfterUpdate  ' ** Procedure: Below.
220           chkPageOf_AfterUpdate  ' ** Procedure: Below.

              ' ** EVENT CHECK: chkRememberDates!
230           Select Case IsNull(.DateAsOf)
              Case True
240             .DateAsOf = TodaysDate  ' ** Module Function: modUtilities.
250           Case False
260             DateAsOf_Exit Cancel  ' ** Procedure: Below.
270           End Select

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

390     DoCmd.Hourglass False

EXITP:
400     Exit Sub

ERRH:
410     DoCmd.Hourglass False
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

          ' ** Create an instance of the Calendar class.
520       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
530       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
540       clsMonthClass.PositionAtCursor = True

550       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

560     End With

EXITP:
570     Exit Sub

ERRH:
580     DoCmd.Hourglass False
590     Select Case ERR.Number
        Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

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

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Focus Holder:     F {FocusHolder}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdClose_Click  ' ** Procedure: Below.
790       Case vbKeyF7
800         With Me
810           intRetVal = 0
820           If .cmdCalendar.Enabled = True Then
830             .cmdCalendar.SetFocus
840             cmdCalendar_Click  ' ** Procedure: Below.
850           End If
860         End With
870       End Select
880     End If

        ' ** Alt keys.
890     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyX
910         intRetVal = 0
920         cmdClose_Click  ' ** Procedure: Below.
930       End Select
940     End If

        ' ** Ctrl keys.
950     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyG
970         With Me
980           intRetVal = 0
990           If .chkPageOf.Enabled = True Then
1000            .chkPageOf.SetFocus
1010            .chkPageOf = (Not .chkPageOf)
1020            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1030          End If
1040        End With
1050      End Select
1060    End If

        ' ** Ctrl-Shift keys.
1070    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1080      Select Case intRetVal
          Case vbKeyF
1090        With Me
1100          intRetVal = 0
1110          .FocusHolder.SetFocus
1120          .DateAsOf = Null
1130        End With
1140      End Select
1150    End If

EXITP:
1160    KeyCode = intRetVal
1170    Exit Sub

ERRH:
1180    intRetVal = 0
1190    Select Case ERR.Number
        Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1310    With Me
1320      .TimerInterval = 0&
1330      If gblnSetFocus = True Then
1340        gblnSetFocus = False
1350        strFormName = vbNullString
1360  On Error Resume Next
1370        Do Until strFormName = THIS_NAME
1380          DoCmd.SelectObject acForm, THIS_NAME, False
1390          strFormName = Screen.ActiveForm.Name
1400        Loop
1410  On Error GoTo ERRH
1420      End If
1430      If gblnGoToReport = True Then
1440        If .cmdPrint.Enabled = True Then
1450          .cmdPrint.SetFocus
1460        Else
1470          .cmdClose.SetFocus
1480        End If
1490        gblnGoToReport = False
1500        Beep
1510        DoCmd.Hourglass False
1520        DoEvents
1530      Else
1540        DoCmd.Hourglass False
1550      End If
1560    End With

EXITP:
1570    Exit Sub

ERRH:
1580    DoCmd.Hourglass False
1590    Select Case ERR.Number
        Case Else
1600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1610    End Select
1620    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1710    Select Case DataErr
        Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1720      If Left(Me.ActiveControl.Name, 4) = "Date" Then
1730        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1740        Response = acDataErrContinue
1750      End If
1760    Case Else
1770      zErrorHandler THIS_NAME, "Undefined Form Error", DataErr  ' ** Module Function: modErrorHandler.
1780    End Select

EXITP:
1790    Exit Sub

ERRH:
1800    Select Case ERR.Number
        Case Else
1810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1820    End Select
1830    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1910    DoCmd.Hourglass True
1920    DoEvents

1930    With Me
1940      If .chkRememberDates = False Then
1950        .DateAsOf = Null  ' ** So it won't be there when Pref_Save() runs.
1960        DoEvents
1970      End If
1980    End With

1990    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2000    Do While Reports.Count > 0
2010      DoCmd.Close acReport, Reports(0).Name
2020    Loop

2030    gstrFormQuerySpec = vbNullString
2040    gblnSetFocus = True
2050    gstrReturningForm = THIS_NAME
2060    Select Case strCallingForm
        Case "frmMenu_Report"
2070      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2080    End Select

EXITP:
2090    Exit Sub

ERRH:
2100    Select Case ERR.Number
        Case Else
2110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2120    End Select
2130    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2210    With Me
2220      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
2230        Select Case blnCalendar1_Focus
            Case True
2240          .cmdCalendar_raised_semifocus_dots_img.Visible = True
2250          .cmdCalendar_raised_img.Visible = False
2260        Case False
2270          .cmdCalendar_raised_img.Visible = True
2280          .cmdCalendar_raised_semifocus_dots_img.Visible = False
2290        End Select
2300        .cmdCalendar_raised_focus_img.Visible = False
2310        .cmdCalendar_raised_focus_dots_img.Visible = False
2320        .cmdCalendar_sunken_focus_dots_img.Visible = False
2330        .cmdCalendar_raised_img_dis.Visible = False
2340      End If
2350    End With

EXITP:
2360    Exit Sub

ERRH:
2370    Select Case ERR.Number
        Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2510    DoCmd.Hourglass True
2520    DoEvents

2530    DoCmd.Close acForm, THIS_NAME

EXITP:
2540    Exit Sub

ERRH:
2550    Select Case ERR.Number
        Case Else
2560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2570    End Select
2580    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2620    intShiftDown = (Shift And acShiftMask) > 0
2630    intAltDown = (Shift And acAltMask) > 0
2640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyTab
2670        With Me
2680          intRetVal = 0
2690          .DateAsOf.SetFocus
2700        End With
2710      End Select
2720    End If

        ' ** Shift keys.
2730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2740      Select Case intRetVal
          Case vbKeyTab
2750        With Me
2760          intRetVal = 0
2770          If .cmdExcel.Enabled = True Then
2780            .cmdExcel.SetFocus
2790          Else
2800            .cmdWord.SetFocus
2810          End If
2820        End With
2830      End Select
2840    End If

EXITP:
2850    KeyCode = intRetVal
2860    Exit Sub

ERRH:
2870    intRetVal = 0
2880    Select Case ERR.Number
        Case Else
2890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2900    End Select
2910    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

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

Private Sub DateAsOf_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          .cmdPreview.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3250        With Me
3260          intRetVal = 0
3270          .cmdClose.SetFocus
3280        End With
3290      End Select
3300    End If

        ' ** Ctrl keys.
3310    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3330        With Me
3340          intRetVal = 0
3350          .cmdCalendar.SetFocus
3360        End With
3370      End Select
3380    End If

EXITP:
3390    KeyCode = intRetVal
3400    Exit Sub

ERRH:
3410    intRetVal = 0
3420    Select Case ERR.Number
        Case Else
3430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3440    End Select
3450    Resume EXITP

End Sub

Private Sub DateAsOf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_MouseMove"

3510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub DateAsOf_Exit(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_Exit"

3610    With Me
3620      If IsNull(.DateAsOf) = True Then
            ' ** Populate it with today's date.
3630        .DateAsOf = Date
3640      Else
3650        If Trim(.DateAsOf) = vbNullString Then
              ' ** Populate it with today's date.
3660          .DateAsOf = Date
3670        Else
3680          If .DateAsOf < Date Then
3690            MsgBox "Data is not rolled back to the entered date." & vbCrLf & _
                  "The report reflects current information.", vbInformation + vbOKOnly, "Current Data Only"
3700          ElseIf .DateAsOf > Date Then
3710            MsgBox "Report will reflect current information.", vbInformation + vbOKOnly, "Current Data Only"
3720          End If
3730        End If
3740      End If
3750    End With

EXITP:
3760    Exit Sub

ERRH:
3770    Select Case ERR.Number
        Case Else
3780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3790    End Select
3800    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

3910    With Me
3920      blnCalendar1_Focus = True
3930      .cmdCalendar_raised_semifocus_dots_img.Visible = True
3940      .cmdCalendar_raised_img.Visible = False
3950      .cmdCalendar_raised_focus_img.Visible = False
3960      .cmdCalendar_raised_focus_dots_img.Visible = False
3970      .cmdCalendar_sunken_focus_dots_img.Visible = False
3980      .cmdCalendar_raised_img_dis.Visible = False
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case Else
4020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4030    End Select
4040    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

4110    With Me
4120      blnCalendar1_MouseDown = True
4130      .cmdCalendar_sunken_focus_dots_img.Visible = True
4140      .cmdCalendar_raised_img.Visible = False
4150      .cmdCalendar_raised_semifocus_dots_img.Visible = False
4160      .cmdCalendar_raised_focus_img.Visible = False
4170      .cmdCalendar_raised_focus_dots_img.Visible = False
4180      .cmdCalendar_raised_img_dis.Visible = False
4190    End With

EXITP:
4200    Exit Sub

ERRH:
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Private Sub cmdCalendar_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

4310    With Me
4320      datStartDate = Date
4330      datEndDate = 0
4340      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
4350      If blnRetVal = True Then
4360        .DateAsOf = datStartDate
4370      Else
4380        .DateAsOf = CDate(Format(Date, "mm/dd/yyyy"))
4390      End If
4400      .DateAsOf.SetFocus
4410    End With

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab
4570        With Me
4580          intRetVal = 0
4590          .cmdPreview.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab
4650        With Me
4660          intRetVal = 0
4670          .DateAsOf.SetFocus
4680        End With
4690      End Select
4700    End If

        ' ** Ctrl keys.
4710    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4720      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4730        With Me
4740          intRetVal = 0
4750          .chkRememberDates.SetFocus
4760        End With
4770      End Select
4780    End If

EXITP:
4790    KeyCode = intRetVal
4800    Exit Sub

ERRH:
4810    intRetVal = 0
4820    Select Case ERR.Number
        Case Else
4830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4840    End Select
4850    Resume EXITP

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

Private Sub chkRememberDates_AfterUpdate()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

5610    With Me
5620      Select Case .chkRememberDates
          Case True
5630        .chkRememberDates_lbl.FontBold = True
5640      Case False
5650        .chkRememberDates_lbl.FontBold = False
5660      End Select
5670    End With

EXITP:
5680    Exit Sub

ERRH:
5690    Select Case ERR.Number
        Case Else
5700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5710    End Select
5720    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5820    intShiftDown = (Shift And acShiftMask) > 0
5830    intAltDown = (Shift And acAltMask) > 0
5840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5860      Select Case intRetVal
          Case vbKeyTab
5870        With Me
5880          intRetVal = 0
5890          .chkPageOf.SetFocus
5900        End With
5910      End Select
5920    End If

        ' ** Shift keys.
5930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5940      Select Case intRetVal
          Case vbKeyTab
5950        With Me
5960          intRetVal = 0
5970          .cmdCalendar.SetFocus
5980        End With
5990      End Select
6000    End If

EXITP:
6010    KeyCode = intRetVal
6020    Exit Sub

ERRH:
6030    intRetVal = 0
6040    Select Case ERR.Number
        Case Else
6050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6060    End Select
6070    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

6110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6120    Exit Sub

ERRH:
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

6210    With Me
6220      Select Case .chkPageOf
          Case True
6230        .chkPageOf_lbl.FontBold = True
6240      Case False
6250        .chkPageOf_lbl.FontBold = False
6260      End Select
6270    End With

EXITP:
6280    Exit Sub

ERRH:
6290    Select Case ERR.Number
        Case Else
6300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6310    End Select
6320    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6420    intShiftDown = (Shift And acShiftMask) > 0
6430    intAltDown = (Shift And acAltMask) > 0
6440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6460      Select Case intRetVal
          Case vbKeyTab
6470        With Me
6480          intRetVal = 0
6490          .cmdPreview.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab
6550        With Me
6560          intRetVal = 0
6570          .chkRememberDates.SetFocus
6580        End With
6590      End Select
6600    End If

EXITP:
6610    KeyCode = intRetVal
6620    Exit Sub

ERRH:
6630    intRetVal = 0
6640    Select Case ERR.Number
        Case Else
6650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6660    End Select
6670    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

6710    If DoReport = True Then  ' ** Function: Below.
          ' ** Report also has 'No Data' MsgBox that cancels opening.
6720      strDocName = "rptAccountBalance"
6730      DoCmd.OpenReport strDocName, acViewPreview
6740      DoCmd.Maximize
6750      DoCmd.RunCommand acCmdFitToWindow
6760    End If

EXITP:
6770    Exit Sub

ERRH:
6780    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
6790    Case Else
6800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6810    End Select
6820    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
          Case vbKeyTab
6970        With Me
6980          intRetVal = 0
6990          .cmdPrint.SetFocus
7000        End With
7010      End Select
7020    End If

        ' ** Shift keys.
7030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7040      Select Case intRetVal
          Case vbKeyTab
7050        With Me
7060          intRetVal = 0
7070          .DateAsOf.SetFocus
7080        End With
7090      End Select
7100    End If

        ' ** Ctrl-Shift keys.
7110    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7120      Select Case intRetVal
          Case vbKeyTab
7130        With Me
7140          intRetVal = 0
7150          .chkPageOf.SetFocus
7160        End With
7170      End Select
7180    End If

EXITP:
7190    KeyCode = intRetVal
7200    Exit Sub

ERRH:
7210    intRetVal = 0
7220    Select Case ERR.Number
        Case Else
7230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7240    End Select
7250    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

7310    If DoReport = True Then  ' ** Function: Below.
          ' ** Report also has 'No Data' MsgBox that cancels opening.
7320      strDocName = "rptAccountBalance"
7330      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
            '##GTR_Ref: rptAccountBalance
7340        DoCmd.OpenReport strDocName, acViewPreview
7350        DoCmd.Maximize
7360        DoCmd.RunCommand acCmdFitToWindow
7370      Else
7380        DoCmd.OpenReport strDocName, acViewNormal
7390      End If
7400    End If

EXITP:
7410    Exit Sub

ERRH:
7420    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
7430    Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

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
          Case vbKeyTab
7570        With Me
7580          intRetVal = 0
7590          .cmdWord.SetFocus
7600        End With
7610      End Select
7620    End If

        ' ** Shift keys.
7630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7640      Select Case intRetVal
          Case vbKeyTab
7650        With Me
7660          intRetVal = 0
7670          .cmdPreview.SetFocus
7680        End With
7690      End Select
7700    End If

EXITP:
7710    KeyCode = intRetVal
7720    Exit Sub

ERRH:
7730    intRetVal = 0
7740    Select Case ERR.Number
        Case Else
7750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7760    End Select
7770    Resume EXITP

End Sub

Private Sub cmdWord_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

7810    With Me
7820      If DoReport = True Then  ' ** Function: Below.

7830        If IsNull(.UserReportPath) = True Then
7840          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
7850        Else
7860          strRptPath = .UserReportPath
7870        End If
7880        strRptCap = "rptAccountBalance_" & Format$(Date, "yyyymmdd")

7890        strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

7900        If strRptPathFile <> vbNullString Then
7910          strRpt = "rptAccountBalance"
7920          DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
7930          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
7940        End If

7950      End If
7960    End With

EXITP:

7970    Exit Sub

ERRH:
7980    Select Case ERR.Number
        Case Else
7990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8000    End Select
8010    Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8120    intShiftDown = (Shift And acShiftMask) > 0
8130    intAltDown = (Shift And acAltMask) > 0
8140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8160      Select Case intRetVal
          Case vbKeyTab
8170        With Me
8180          intRetVal = 0
8190          If .cmdExcel.Enabled = True Then
8200            .cmdExcel.SetFocus
8210          Else
8220            .cmdClose.SetFocus
8230          End If
8240        End With
8250      End Select
8260    End If

        ' ** Shift keys.
8270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8280      Select Case intRetVal
          Case vbKeyTab
8290        With Me
8300          intRetVal = 0
8310          .cmdPrint.SetFocus
8320        End With
8330      End Select
8340    End If

EXITP:
8350    KeyCode = intRetVal
8360    Exit Sub

ERRH:
8370    intRetVal = 0
8380    Select Case ERR.Number
        Case Else
8390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8400    End Select
8410    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

8510    With Me
8520      If DoReport = True Then  ' ** Function: Below.

8530        If IsNull(.UserReportPath) = True Then
8540          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
8550        Else
8560          strRptPath = .UserReportPath
8570        End If
8580        strRptCap = "rptAccountBalance_" & Format$(Date, "yyyymmdd")

8590        strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

8600        If strRptPathFile <> vbNullString Then

8610          gstrFormQuerySpec = THIS_NAME
8620          gdatStartDate = .DateAsOf

              ' ** qryRpt_AccountBalance_13_05 (xx), For Export.
8630          strQry = "qryRpt_AccountBalance_14"

8640          DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False

8650          If Excel_NameOnly(strRptPathFile, "Account Balance") = True Then  ' ** Module Function: modExcelFuncs.
8660            If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
8670              EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
8680            End If
8690            OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
8700          End If

8710          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

8720        End If

8730      End If
8740    End With

EXITP:
8750    Exit Sub

ERRH:
8760    Select Case ERR.Number
        Case Else
8770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8780    End Select
8790    Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

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
          Case vbKeyTab
8870        With Me
8880          intRetVal = 0
8890          .cmdClose.SetFocus
8900        End With
8910      End Select
8920    End If

        ' ** Shift keys.
8930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8940      Select Case intRetVal
          Case vbKeyTab
8950        With Me
8960          intRetVal = 0
8970          .cmdWord.SetFocus
8980        End With
8990      End Select
9000    End If

EXITP:
9010    KeyCode = intRetVal
9020    Exit Sub

ERRH:
9030    intRetVal = 0
9040    Select Case ERR.Number
        Case Else
9050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9060    End Select
9070    Resume EXITP

End Sub

Private Function DoReport() As Boolean

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

9110    blnRetVal = True

9120    With Me
9130      If IsNull(.DateAsOf) = True Then
9140        blnRetVal = False
9150        MsgBox "You must enter a date to continue.", vbInformation + vbOKOnly, "Entry Required"
9160      End If
9170    End With

9180    If blnRetVal = True Then
9190      Set dbs = CurrentDb
9200      With dbs
            ' ** qryRpt_AccountBalance_11a (Account, linked to ActiveAssets), grouped and summed.
9210        Set qdf = .QueryDefs("qryRpt_AccountBalance_11")
9220        Set rst = qdf.OpenRecordset
9230        With rst
9240          If .BOF = True And .EOF = True Then
                ' ** Report also has 'No Data' MsgBox that cancels opening.
9250            blnRetVal = False
9260            MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
9270          End If
9280          .Close
9290        End With
9300        .Close
9310      End With
          'MV: Sum([qryassetlist].[totalshareface]*[marketvaluecurrentx])+First([qryassetlist].[icash])+First([qryassetlist].[pcash])
          'cost: Sum([qryassetlist].[totalcost])+First([qryassetlist].[icash])+First([qryassetlist].[pcash])
9320    End If

EXITP:
9330    Set rst = Nothing
9340    Set qdf = Nothing
9350    Set dbs = Nothing
9360    DoReport = blnRetVal
9370    Exit Function

ERRH:
9380    blnRetVal = False
9390    Select Case ERR.Number
        Case Else
9400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9410    End Select
9420    Resume EXITP

End Function
