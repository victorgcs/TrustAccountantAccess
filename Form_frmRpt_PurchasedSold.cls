VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_PurchasedSold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_PurchasedSold"

'VGC 09/29/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, opens, shows arrow, and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Sold Word:        D {cmdAssetPurchasedSold_SoldWord}
' **   Purch Excel:      E {cmdAssetPurchasedSold_PurchasedExcel}
' **   From:             F {DateStart}
' **   Sold Print:       N {cmdAssetPurchasedSold_SoldPrint}
' **   Purch Word:       O {cmdAssetPurchasedSold_PurchasedWord}
' **   Purch Print:      P {cmdAssetPurchasedSold_PurchasedPrint}
' **   To:               T {DateEnd}
' **   Sold Preview:     V {cmdAssetPurchasedSold_SoldPreview}
' **   Purch Preview:    W {cmdAssetPurchasedSold_PurchasedPreview}
' **   Exit:             X {cmdClose}
' **   Sold Excel:       Z {cmdAssetPurchasedSold_SoldExcel}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Print All:        A {cmdPrintAll}
' **   Excel All:        E {cmdExcelAll}
' **   Page Count:       G {chkPageOf}
' **   Open Excel:       N {chkOpenExcel}
' **   Open Word:        O {chkOpenWord}
' **   Word All:         W {cmdWordAll}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean
Private blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean

Private strFile1 As String, strFile2 As String
Private blnPrintBoth As Boolean, blnGoingToReport As Boolean
Private strCallingForm As String, lngTpp As Long
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
160           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
170           blnPrintAll_Focus = False: blnWordAll_Focus = False: blnExcelAll_Focus = False

180           gstrFormQuerySpec = THIS_NAME
190           blnPrintBoth = False

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
220           .cmdAssetPurchasedSold_PurchasedExcel.Enabled = False
230           .cmdAssetPurchasedSold_SoldExcel.Enabled = False
240           .cmdExcelAll.Enabled = False
250           .chkOpenExcel.Enabled = False
260           .chkOpenExcel_lbl2.ForeColor = WIN_CLR_DISF
270           .chkOpenExcel_lbl2_dim_hi.Visible = True
      #Else
280           .cmdAssetPurchasedSold_PurchasedExcel.Enabled = True
290           .cmdAssetPurchasedSold_SoldExcel.Enabled = True
300           .cmdExcelAll.Enabled = True
310           .chkOpenExcel.Enabled = True
320           .chkOpenExcel_lbl2.ForeColor = CLR_DKGRY
330           .chkOpenExcel_lbl2_dim_hi.Visible = False
      #End If

              ' ** EVENT CHECK: chkRememberDates!
340           If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
350             DateStart_Exit Cancel  ' ** Procedure: Below.
360             DateEnd_Exit Cancel  ' ** Procedure: Below.
370           End If

380           chkRememberDates_AfterUpdate  ' ** Procedure: Below.
390           chkPageOf_AfterUpdate  ' ** Procedure: Below.
400           chkOpenWord_AfterUpdate  ' ** Procedure: Below.
410           chkOpenExcel_AfterUpdate  ' ** Procedure: Below.

420         Else
430           Cancel = -1
440         End If
450       Else
460         Cancel = -1
470       End If
480     End With

490     If Cancel = -1 Then
500       Beep
510       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
520     End If

530     DoCmd.Hourglass False

EXITP:
540     Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
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

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Print All:        A {cmdPrintAll}
        ' **   Excel All:        E {cmdExcelAll}
        ' **   Page Count:       G {chkPageOf}
        ' **   Open Excel:       N {chkOpenExcel}
        ' **   Open Word:        O {chkOpenWord}
        ' **   Word All:         W {cmdWordAll}

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
          Case vbKeyX
990         intRetVal = 0
1000        cmdClose_Click  ' ** Procedure: Below.
1010      End Select
1020    End If

        ' ** Ctrl keys.
1030    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1040      Select Case intRetVal
          Case vbKeyA
1050        With Me
1060          intRetVal = 0
1070          .cmdPrintAll.SetFocus
1080          cmdPrintAll_Click  ' ** Procedure: Below.
1090        End With
1100      Case vbKeyE
1110        With Me
1120          intRetVal = 0
1130          .cmdExcelAll.SetFocus
1140          cmdExcelAll_Click  ' ** Procedure: Below.
1150        End With
1160      Case vbKeyG
1170        With Me
1180          intRetVal = 0
1190          If .chkPageOf.Enabled = True Then
1200            .chkPageOf.SetFocus
1210            .chkPageOf = (Not .chkPageOf)
1220            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1230          End If
1240        End With
1250      Case vbKeyN
1260        With Me
1270          intRetVal = 0
1280          .chkOpenExcel.SetFocus
1290          .chkOpenExcel = (Not .chkOpenExcel)
1300          chkOpenExcel_AfterUpdate  ' ** Procedure: Below.
1310        End With
1320      Case vbKeyO
1330        With Me
1340          intRetVal = 0
1350          .chkOpenWord.SetFocus
1360          .chkOpenWord = (Not .chkOpenWord)
1370          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
1380        End With
1390      Case vbKeyW
1400        With Me
1410          intRetVal = 0
1420          .cmdWordAll.SetFocus
1430          cmdWordAll_Click  ' ** Procedure: Below.
1440        End With
1450      End Select
1460    End If

        ' ** Ctrl-Shift keys.
1470    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1480      Select Case intRetVal
          Case vbKeyF
1490        With Me
1500          intRetVal = 0
1510          .FocusHolder.SetFocus
1520          .DateStart = Null
1530        End With
1540      End Select
1550    End If

EXITP:
1560    KeyCode = intRetVal
1570    Exit Sub

ERRH:
1580    intRetVal = 0
1590    THAT_PROC = THIS_PROC
1600    That_Erl = Erl
1610    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1620    Resume EXITP

End Sub

Private Sub Form_Timer()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1710    With Me
1720      .TimerInterval = 0&
1730      If gblnSetFocus = True Then
1740        gblnSetFocus = False
1750        strFormName = vbNullString
1760  On Error Resume Next
1770        Do Until strFormName = THIS_NAME
1780          DoCmd.SelectObject acForm, THIS_NAME, False
1790          strFormName = Screen.ActiveForm.Name
1800        Loop
1810  On Error GoTo ERRH
1820      End If
1830      If gblnGoToReport = True Then
1840        Select Case blnGoingToReport
            Case True
1850          DoCmd.Hourglass False  ' ** It should already be off.
1860          blnGoingToReport = False
1870          gblnGoToReport = False
1880          .GoToReport_arw_rpt_purch_img.Visible = False
1890          .GoToReport_arw_rpt_sold_img.Visible = False
1900          .cmdAssetPurchasedSold_PurchasedWord.Visible = True
1910          .cmdAssetPurchasedSold_SoldWord.Visible = True
1920          DoEvents
1930        Case False
1940          DoCmd.Hourglass True  ' ** Make sure it's still running.
1950          DoEvents
1960          blnGoingToReport = True
1970          Select Case garr_varGoToReport(GTR_RNAM)
              Case "rptPurchasedSold_Purchased"
1980            .cmdAssetPurchasedSold_PurchasedPrint.SetFocus
1990            .cmdAssetPurchasedSold_PurchasedWord.Visible = False
2000            .GoToReport_arw_rpt_purch_img.Visible = True
2010          Case "rptPurchasedSold_Sold"
2020            .cmdAssetPurchasedSold_SoldPrint.SetFocus
2030            .cmdAssetPurchasedSold_SoldWord.Visible = False
2040            .GoToReport_arw_rpt_sold_img.Visible = True
2050          End Select
2060          DoEvents
2070          Beep
2080          DoCmd.Hourglass False
2090          DoEvents
2100          .TimerInterval = (GTR_WAIT * 2&)
2110        End Select
2120      Else
2130        DoCmd.Hourglass False
2140      End If
2150    End With

EXITP:
2160    Exit Sub

ERRH:
2170    DoCmd.Hourglass False
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
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
2320    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2330      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
2340        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
2350        Response = acDataErrContinue
2360      End If
2370    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
2380    Case Else
2390      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2400      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2410    End Select

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2510    DoCmd.Hourglass True
2520    DoEvents

2530    With Me
2540      If .chkRememberDates = False Then
2550        .DateStart = Null  ' ** So it won't be there when Pref_Save() runs.
2560        .DateEnd = Null
2570      End If
2580    End With

2590    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2600    Do While Reports.Count > 0
2610      DoCmd.Close acReport, Reports(0).Name
2620    Loop

2630    gstrFormQuerySpec = vbNullString
2640    gblnSetFocus = True
2650    gstrReturningForm = THIS_NAME
2660    Select Case strCallingForm
        Case "frmMenu_Report"
2670      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2680    End Select

EXITP:
2690    Exit Sub

ERRH:
2700    Select Case ERR.Number
        Case Else
2710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2720    End Select
2730    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

2810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2910    With Me
2920      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
2930        Select Case blnCalendar1_Focus
            Case True
2940          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
2950          .cmdCalendar1_raised_img.Visible = False
2960        Case False
2970          .cmdCalendar1_raised_img.Visible = True
2980          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
2990        End Select
3000        .cmdCalendar1_raised_focus_dots_img.Visible = False
3010        .cmdCalendar1_raised_focus_img.Visible = False
3020        .cmdCalendar1_sunken_focus_dots_img.Visible = False
3030        .cmdCalendar1_raised_img_dis.Visible = False
3040      End If
3050      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
3060        Select Case blnCalendar2_Focus
            Case True
3070          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
3080          .cmdCalendar2_raised_img.Visible = False
3090        Case False
3100          .cmdCalendar2_raised_img.Visible = True
3110          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
3120        End Select
3130        .cmdCalendar2_raised_focus_dots_img.Visible = False
3140        .cmdCalendar2_raised_focus_img.Visible = False
3150        .cmdCalendar2_sunken_focus_dots_img.Visible = False
3160        .cmdCalendar2_raised_img_dis.Visible = False
3170      End If
3180      If blnPrintAll_Focus = False And (.cmdPrintAll_box01.Visible = True Or .cmdPrintAll_box02.Visible = True) Then
3190        .cmdPrintAll_box01.Visible = False
3200        .cmdPrintAll_box02.Visible = False
3210      End If
3220      If blnWordAll_Focus = False And (.cmdWordAll_box01.Visible = True Or .cmdWordAll_box02.Visible = True) Then
3230        .cmdWordAll_box01.Visible = False
3240        .cmdWordAll_box02.Visible = False
3250      End If
3260      If blnExcelAll_Focus = False And (.cmdExcelAll_box01.Visible = True Or .cmdExcelAll_box02.Visible = True) Then
3270        .cmdExcelAll_box01.Visible = False
3280        .cmdExcelAll_box02.Visible = False
3290      End If
3300    End With

EXITP:
3310    Exit Sub

ERRH:
3320    Select Case ERR.Number
        Case Else
3330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3340    End Select
3350    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub cmdClose_Click()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3510    DoCmd.Hourglass True
3520    DoEvents

3530    DoCmd.Close acForm, THIS_NAME

EXITP:
3540    Exit Sub

ERRH:
3550    Select Case ERR.Number
        Case Else
3560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3570    End Select
3580    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3620    intShiftDown = (Shift And acShiftMask) > 0
3630    intAltDown = (Shift And acAltMask) > 0
3640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3660      Select Case intRetVal
          Case vbKeyTab
3670        With Me
3680          intRetVal = 0
3690          .DateStart.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab
3750        With Me
3760          intRetVal = 0
3770          If .cmdExcelAll.Enabled = True Then
3780            .cmdExcelAll.SetFocus
3790          Else
3800            .cmdWordAll.SetFocus
3810          End If
3820        End With
3830      End Select
3840    End If

        ' ** Ctrl-Shift keys.
3850    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3860      Select Case intRetVal
          Case vbKeyTab
3870        With Me
3880          intRetVal = 0
3890          If .chkOpenExcel.Enabled = True Then
3900            .chkOpenExcel.SetFocus
3910          Else
3920            .chkOpenWord.SetFocus
3930          End If
3940        End With
3950      End Select
3960    End If

EXITP:
3970    KeyCode = intRetVal
3980    Exit Sub

ERRH:
3990    intRetVal = 0
4000    THAT_PROC = THIS_PROC
4010    That_Erl = Erl
4020    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4030    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

4110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4270        With Me
4280          intRetVal = 0
4290          .DateEnd.SetFocus
4300        End With
4310      End Select
4320    End If

        ' ** Shift keys.
4330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4350        With Me
4360          intRetVal = 0
4370          .cmdClose.SetFocus
4380        End With
4390      End Select
4400    End If

        ' ** Ctrl keys.
4410    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4420      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4430        With Me
4440          intRetVal = 0
4450          .cmdCalendar1.SetFocus
4460        End With
4470      End Select
4480    End If

EXITP:
4490    KeyCode = intRetVal
4500    Exit Sub

ERRH:
4510    intRetVal = 0
4520    THAT_PROC = THIS_PROC
4530    That_Erl = Erl
4540    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4550    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

4610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub DateStart_Exit(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

4710    With Me
4720      If IsDate(.DateEnd) = True And IsDate(.DateStart) = True Then
4730        If CDate(.DateStart) > CDate(.DateEnd) Then
4740          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
4750          Cancel = -1
4760        End If
4770      End If
4780    End With

EXITP:
4790    Exit Sub

ERRH:
4800    Select Case ERR.Number
        Case Else
4810      THAT_PROC = THIS_PROC
4820      That_Erl = Erl
4830      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4840    End Select
4850    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

4910    With Me
4920      blnCalendar1_Focus = True
4930      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4940      .cmdCalendar1_raised_img.Visible = False
4950      .cmdCalendar1_raised_focus_img.Visible = False
4960      .cmdCalendar1_raised_focus_dots_img.Visible = False
4970      .cmdCalendar1_sunken_focus_dots_img.Visible = False
4980      .cmdCalendar1_raised_img_dis.Visible = False
4990    End With

EXITP:
5000    Exit Sub

ERRH:
5010    Select Case ERR.Number
        Case Else
5020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5030    End Select
5040    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

5110    With Me
5120      blnCalendar1_MouseDown = True
5130      .cmdCalendar1_sunken_focus_dots_img.Visible = True
5140      .cmdCalendar1_raised_img.Visible = False
5150      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5160      .cmdCalendar1_raised_focus_img.Visible = False
5170      .cmdCalendar1_raised_focus_dots_img.Visible = False
5180      .cmdCalendar1_raised_img_dis.Visible = False
5190    End With

EXITP:
5200    Exit Sub

ERRH:
5210    Select Case ERR.Number
        Case Else
5220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5230    End Select
5240    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

5310    With Me
5320      datStartDate = Date
5330      datEndDate = 0
5340      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
5350      If blnRetVal = True Then
5360        .DateStart = datStartDate
5370      Else
5380        .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
5390      End If
5400      .DateStart.SetFocus
5410    End With

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
5440    Case Else
5450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5460    End Select
5470    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5520    intShiftDown = (Shift And acShiftMask) > 0
5530    intAltDown = (Shift And acAltMask) > 0
5540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5570        With Me
5580          intRetVal = 0
5590          .DateEnd.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5650        With Me
5660          intRetVal = 0
5670          .DateStart.SetFocus
5680        End With
5690      End Select
5700    End If

EXITP:
5710    KeyCode = intRetVal
5720    Exit Sub

ERRH:
5730    intRetVal = 0
5740    THAT_PROC = THIS_PROC
5750    That_Erl = Erl
5760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5770    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

5810    With Me
5820      If blnCalendar1_MouseDown = False Then
5830        Select Case blnCalendar1_Focus
            Case True
5840          .cmdCalendar1_raised_focus_dots_img.Visible = True
5850          .cmdCalendar1_raised_focus_img.Visible = False
5860        Case False
5870          .cmdCalendar1_raised_focus_img.Visible = True
5880          .cmdCalendar1_raised_focus_dots_img.Visible = False
5890        End Select
5900        .cmdCalendar1_raised_img.Visible = False
5910        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5920        .cmdCalendar1_sunken_focus_dots_img.Visible = False
5930        .cmdCalendar1_raised_img_dis.Visible = False
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

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

6110    With Me
6120      .cmdCalendar1_raised_focus_dots_img.Visible = True
6130      .cmdCalendar1_raised_img.Visible = False
6140      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6150      .cmdCalendar1_raised_focus_img.Visible = False
6160      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6170      .cmdCalendar1_raised_img_dis.Visible = False
6180      blnCalendar1_MouseDown = False
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

Private Sub cmdCalendar1_LostFocus()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

6310    With Me
6320      .cmdCalendar1_raised_img.Visible = True
6330      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6340      .cmdCalendar1_raised_focus_img.Visible = False
6350      .cmdCalendar1_raised_focus_dots_img.Visible = False
6360      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6370      .cmdCalendar1_raised_img_dis.Visible = False
6380      blnCalendar1_Focus = False
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

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6520    intShiftDown = (Shift And acShiftMask) > 0
6530    intAltDown = (Shift And acAltMask) > 0
6540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6570        With Me
6580          intRetVal = 0
6590          .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
6600        End With
6610      End Select
6620    End If

        ' ** Shift keys.
6630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6650        With Me
6660          intRetVal = 0
6670          .DateStart.SetFocus
6680        End With
6690      End Select
6700    End If

        ' ** Ctrl keys.
6710    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6720      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6730        With Me
6740          intRetVal = 0
6750          .cmdCalendar2.SetFocus
6760        End With
6770      End Select
6780    End If

        ' ** Ctrl-Shift keys.
6790    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6800      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6810        With Me
6820          intRetVal = 0
6830          .cmdCalendar1.SetFocus
6840        End With
6850      End Select
6860    End If

EXITP:
6870    KeyCode = intRetVal
6880    Exit Sub

ERRH:
6890    intRetVal = 0
6900    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
6910    Case Else
6920      THAT_PROC = THIS_PROC
6930      That_Erl = Erl
6940      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6950    End Select
6960    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

7010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

7110    With Me
7120      If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
7130        .DateEnd = Date
7140      End If
7150      If IsNull(.DateStart) Then
7160        MsgBox "The start date must be prior to or the same as the end date.", vbInformation + vbOKOnly, "Entry Required"
7170        .DateStart.SetFocus
7180      Else
7190        If CDate(.DateStart) > CDate(.DateEnd) Then
7200          MsgBox "The start date must be prior to or the same as the end date.", vbInformation + vbOKOnly, "Invalid Date"
7210          .DateStart.SetFocus
7220        End If
7230      End If
7240    End With

EXITP:
7250    Exit Sub

ERRH:
7260    Select Case ERR.Number
        Case Else
7270      THAT_PROC = THIS_PROC
7280      That_Erl = Erl
7290      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7300    End Select
7310    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

7410    With Me
7420      blnCalendar2_Focus = True
7430      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
7440      .cmdCalendar2_raised_img.Visible = False
7450      .cmdCalendar2_raised_focus_img.Visible = False
7460      .cmdCalendar2_raised_focus_dots_img.Visible = False
7470      .cmdCalendar2_sunken_focus_dots_img.Visible = False
7480      .cmdCalendar2_raised_img_dis.Visible = False
7490    End With

EXITP:
7500    Exit Sub

ERRH:
7510    Select Case ERR.Number
        Case Else
7520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7530    End Select
7540    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

7610    With Me
7620      blnCalendar2_MouseDown = True
7630      .cmdCalendar2_sunken_focus_dots_img.Visible = True
7640      .cmdCalendar2_raised_img.Visible = False
7650      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
7660      .cmdCalendar2_raised_focus_img.Visible = False
7670      .cmdCalendar2_raised_focus_dots_img.Visible = False
7680      .cmdCalendar2_raised_img_dis.Visible = False
7690    End With

EXITP:
7700    Exit Sub

ERRH:
7710    Select Case ERR.Number
        Case Else
7720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7730    End Select
7740    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

7810    With Me
7820      datStartDate = Date
7830      datEndDate = 0
7840      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
7850      If blnRetVal = True Then
7860        .DateEnd = datStartDate
7870      Else
7880        .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
7890      End If
7900      .DateEnd.SetFocus
7910      Cancel = 0
7920      DateEnd_Exit Cancel  ' ** Procedure: Above.
7930      If Cancel = 0 Then
7940        .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
7950      End If
7960    End With

EXITP:
7970    Exit Sub

ERRH:
7980    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7990    Case Else
8000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8010    End Select
8020    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
8170        With Me
8180          intRetVal = 0
8190          .chkRememberDates.SetFocus
8200        End With
8210      End Select
8220    End If

        ' ** Shift keys.
8230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8250        With Me
8260          intRetVal = 0
8270          .DateEnd.SetFocus
8280        End With
8290      End Select
8300    End If

EXITP:
8310    KeyCode = intRetVal
8320    Exit Sub

ERRH:
8330    intRetVal = 0
8340    THAT_PROC = THIS_PROC
8350    That_Erl = Erl
8360    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8370    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

8410    With Me
8420      If blnCalendar2_MouseDown = False Then
8430        Select Case blnCalendar2_Focus
            Case True
8440          .cmdCalendar2_raised_focus_dots_img.Visible = True
8450          .cmdCalendar2_raised_focus_img.Visible = False
8460        Case False
8470          .cmdCalendar2_raised_focus_img.Visible = True
8480          .cmdCalendar2_raised_focus_dots_img.Visible = False
8490        End Select
8500        .cmdCalendar2_raised_img.Visible = False
8510        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8520        .cmdCalendar2_sunken_focus_dots_img.Visible = False
8530        .cmdCalendar2_raised_img_dis.Visible = False
8540      End If
8550    End With

EXITP:
8560    Exit Sub

ERRH:
8570    Select Case ERR.Number
        Case Else
8580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8590    End Select
8600    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

8710    With Me
8720      .cmdCalendar2_raised_focus_dots_img.Visible = True
8730      .cmdCalendar2_raised_img.Visible = False
8740      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8750      .cmdCalendar2_raised_focus_img.Visible = False
8760      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8770      .cmdCalendar2_raised_img_dis.Visible = False
8780      blnCalendar2_MouseDown = False
8790    End With

EXITP:
8800    Exit Sub

ERRH:
8810    Select Case ERR.Number
        Case Else
8820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8830    End Select
8840    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

8910    With Me
8920      .cmdCalendar2_raised_img.Visible = True
8930      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8940      .cmdCalendar2_raised_focus_img.Visible = False
8950      .cmdCalendar2_raised_focus_dots_img.Visible = False
8960      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8970      .cmdCalendar2_raised_img_dis.Visible = False
8980      blnCalendar2_Focus = False
8990    End With

EXITP:
9000    Exit Sub

ERRH:
9010    Select Case ERR.Number
        Case Else
9020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9030    End Select
9040    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

9110    With Me
9120      Select Case .chkRememberDates
          Case True
9130        .chkRememberDates_lbl.FontBold = True
9140      Case False
9150        .chkRememberDates_lbl.FontBold = False
9160      End Select
9170    End With

EXITP:
9180    Exit Sub

ERRH:
9190    Select Case ERR.Number
        Case Else
9200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9210    End Select
9220    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

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
9390          .chkPageOf.SetFocus
9400        End With
9410      End Select
9420    End If

        ' ** Shift keys.
9430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9450        With Me
9460          intRetVal = 0
9470          .cmdCalendar2.SetFocus
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

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

9610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
9620    Exit Sub

ERRH:
9630    Select Case ERR.Number
        Case Else
9640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9650    End Select
9660    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

9710    With Me
9720      Select Case .chkPageOf
          Case True
9730        .chkPageOf_lbl.FontBold = True
9740      Case False
9750        .chkPageOf_lbl.FontBold = False
9760      End Select
9770    End With

EXITP:
9780    Exit Sub

ERRH:
9790    DoCmd.Hourglass False
9800    Select Case ERR.Number
        Case Else
9810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9820    End Select
9830    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9920    intShiftDown = (Shift And acShiftMask) > 0
9930    intAltDown = (Shift And acAltMask) > 0
9940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9970        With Me
9980          intRetVal = 0
9990          .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10050       With Me
10060         intRetVal = 0
10070         .chkRememberDates.SetFocus
10080       End With
10090     End Select
10100   End If

EXITP:
10110   KeyCode = intRetVal
10120   Exit Sub

ERRH:
10130   intRetVal = 0
10140   THAT_PROC = THIS_PROC
10150   That_Erl = Erl
10160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10170   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_Sold_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_Sold_box_MouseMove"

10210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
10220   Exit Sub

ERRH:
10230   Select Case ERR.Number
        Case Else
10240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10250   End Select
10260   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_PurchasedPreview_Click()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_PurchasedPreview_Click"

        Dim strDocName As String

10310   If DoReport("Purchased") = True Then  ' ** Function: Below.
10320     strDocName = "rptPurchasedSold_Purchased"
10330     DoCmd.OpenReport strDocName, acViewPreview
10340     DoCmd.Maximize
10350     DoCmd.RunCommand acCmdFitToWindow
10360   End If

EXITP:
10370   Exit Sub

ERRH:
10380   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
10390   Case Else
10400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10410   End Select
10420   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_PurchasedPreview_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_PurchasedPreview_KeyDown"

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
10590         .cmdAssetPurchasedSold_PurchasedPrint.SetFocus
10600       End With
10610     Case vbKeyUp
10620       With Me
10630         intRetVal = 0
10640         .cmdAssetPurchasedSold_SoldPreview.SetFocus
10650       End With
10660     Case vbKeyDown
10670       With Me
10680         intRetVal = 0
10690         .cmdAssetPurchasedSold_SoldPreview.SetFocus
10700       End With
10710     Case vbKeyLeft
10720       With Me
10730         intRetVal = 0
10740         If .cmdExcelAll.Enabled = True Then
10750           .cmdExcelAll.SetFocus
10760         Else
10770           .cmdWordAll.SetFocus
10780         End If
10790       End With
10800     Case vbKeyRight
10810       With Me
10820         intRetVal = 0
10830         .cmdAssetPurchasedSold_PurchasedPrint.SetFocus
10840       End With
10850     End Select
10860   End If

        ' ** Shift keys.
10870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10880     Select Case intRetVal
          Case vbKeyTab
10890       With Me
10900         intRetVal = 0
10910         .DateEnd.SetFocus
10920       End With
10930     End Select
10940   End If

        ' ** Ctrl-Shift keys.
10950   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10960     Select Case intRetVal
          Case vbKeyTab
10970       With Me
10980         intRetVal = 0
10990         .chkPageOf.SetFocus
11000       End With
11010     End Select
11020   End If

EXITP:
11030   KeyCode = intRetVal
11040   Exit Sub

ERRH:
11050   intRetVal = 0
11060   THAT_PROC = THIS_PROC
11070   That_Erl = Erl
11080   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11090   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_PurchasedPrint_Click()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_PurchasedPrint_Click"

        Dim strDocName As String

11110   If DoReport("Purchased") = True Then  ' ** Function: Below.
11120     strDocName = "rptPurchasedSold_Purchased"
11130     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
11140       DoCmd.OpenReport strDocName, acViewPreview
11150       DoCmd.Maximize
11160       DoCmd.RunCommand acCmdFitToWindow
11170     Else
            '##GTR_Ref: rptPurchasedSold_Purchased
11180       DoCmd.OpenReport strDocName, acViewNormal
11190     End If
11200   Else
11210     blnPrintBoth = False
11220   End If

EXITP:
11230   Exit Sub

ERRH:
11240   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
11250   Case Else
11260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11270   End Select
11280   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_PurchasedPrint_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_PurchasedPrint_KeyDown"

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
          Case vbKeyTab
11370       With Me
11380         intRetVal = 0
11390         .cmdAssetPurchasedSold_PurchasedWord.SetFocus
11400       End With
11410     Case vbKeyUp
11420       With Me
11430         intRetVal = 0
11440         .cmdPrintAll.SetFocus
11450       End With
11460     Case vbKeyDown
11470       With Me
11480         intRetVal = 0
11490         .cmdAssetPurchasedSold_SoldPrint.SetFocus
11500       End With
11510     Case vbKeyLeft
11520       With Me
11530         intRetVal = 0
11540         .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
11550       End With
11560     Case vbKeyRight
11570       With Me
11580         intRetVal = 0
11590         .cmdAssetPurchasedSold_PurchasedWord.SetFocus
11600       End With
11610     End Select
11620   End If

        ' ** Shift keys.
11630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11640     Select Case intRetVal
          Case vbKeyTab
11650       With Me
11660         intRetVal = 0
11670         .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
11680       End With
11690     End Select
11700   End If

EXITP:
11710   KeyCode = intRetVal
11720   Exit Sub

ERRH:
11730   intRetVal = 0
11740   THAT_PROC = THIS_PROC
11750   That_Erl = Erl
11760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11770   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_PurchasedWord_Click()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_PurchasedWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

11810   With Me
11820     If DoReport("Purchased") = True Then  ' ** Function: Below.

11830       If IsNull(.UserReportPath) = True Then
11840         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
11850       Else
11860         strRptPath = .UserReportPath
11870       End If

11880       strRpt = "rptPurchasedSoldAssets_Purchased"
11890       gdatStartDate = .DateStart
11900       gdatEndDate = .DateEnd
11910       strRptCap = "Assets_Purchased_" & _
              Format$(gdatStartDate, "yymmdd") & "_To_" & Format$(gdatEndDate, "yymmdd")

11920       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

11930       If strRptPathFile <> vbNullString Then
11940         strFile1 = strRptPathFile
11950         Select Case .chkOpenWord
              Case True
11960           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
11970         Case False
11980           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, False
11990         End Select
12000         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
12010       Else
12020         gblnPrintAll = False
12030       End If

12040     Else
12050       gblnPrintAll = False
12060     End If
12070     gdatStartDate = 0
12080     gdatEndDate = 0
12090   End With

EXITP:
12100   Exit Sub

ERRH:
12110   gblnPrintAll = False
12120   Select Case ERR.Number
        Case Else
12130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
12140   End Select
12150   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_PurchasedWord_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_PurchasedWord_KeyDown"

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
          Case vbKeyTab
12270       With Me
12280         intRetVal = 0
12290         If .cmdAssetPurchasedSold_PurchasedExcel.Enabled = True Then
12300           .cmdAssetPurchasedSold_PurchasedExcel.SetFocus
12310         Else
12320           .cmdAssetPurchasedSold_SoldPreview.SetFocus
12330         End If
12340       End With
12350     Case vbKeyUp
12360       With Me
12370         intRetVal = 0
12380         .cmdAssetPurchasedSold_SoldWord.SetFocus
12390       End With
12400     Case vbKeyDown
12410       With Me
12420         intRetVal = 0
12430         .cmdAssetPurchasedSold_SoldWord.SetFocus
12440       End With
12450     Case vbKeyLeft
12460       With Me
12470         intRetVal = 0
12480         .cmdAssetPurchasedSold_PurchasedPrint.SetFocus
12490       End With
12500     Case vbKeyRight
12510       With Me
12520         intRetVal = 0
12530         If .cmdAssetPurchasedSold_PurchasedExcel.Enabled = True Then
12540           .cmdAssetPurchasedSold_PurchasedExcel.SetFocus
12550         Else
12560           .cmdAssetPurchasedSold_SoldPreview.SetFocus
12570         End If
12580       End With
12590     End Select
12600   End If

        ' ** Shift keys.
12610   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12620     Select Case intRetVal
          Case vbKeyTab
12630       With Me
12640         intRetVal = 0
12650         .cmdAssetPurchasedSold_PurchasedPrint.SetFocus
12660       End With
12670     End Select
12680   End If

EXITP:
12690   KeyCode = intRetVal
12700   Exit Sub

ERRH:
12710   intRetVal = 0
12720   Select Case ERR.Number
        Case Else
12730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12740   End Select
12750   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_PurchasedExcel_Click()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_PurchasedExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

12810   With Me

12820     DoCmd.Hourglass True
12830     DoEvents

12840     blnContinue = True

12850     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12860       DoCmd.Hourglass False
12870       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
12880       If msgResponse <> vbRetry Then
12890         blnContinue = False
12900       End If
12910     End If

12920     If blnContinue = True Then

12930       DoCmd.Hourglass True
12940       DoEvents

12950       If DoReport("Purchased") = True Then  ' ** Procedure: Below.

12960         blnContinue = True

12970         If IsNull(.UserReportPath) = True Then
12980           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
12990         Else
13000           strRptPath = .UserReportPath
13010         End If

13020         strRptCap = "Assets_Purchased_" & Format$(.DateStart, "yymmdd") & "_To_" & Format$(.DateEnd, "yymmdd")

              ' ** qryRpt_PurchasedSold_Purchased_15, For Export.
13030         strQry = "qryRpt_PurchasedSold_Purchased_16"

13040         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

13050         If strRptPathFile <> vbNullString Then
13060           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13070             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
13080           End If
13090           DoEvents
13100           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
13110             Kill strRptPathFile
13120           End If
13130           DoEvents
13140           gstrFormQuerySpec = THIS_NAME
13150           Select Case gblnPrintAll
                Case True
13160             DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
13170             strFile1 = strRptPathFile
13180           Case False
13190             DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
13200             If Excel_NameOnly(strRptPathFile, "Assets Purchased") = True Then  ' ** Module Function: modExcelFuncs.
13210               Select Case .chkOpenExcel
                    Case True
13220                 If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13230                   EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
13240                 End If
13250                 OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
13260               Case False
                      ' ** Nothing, that's it.
13270               End Select
13280             End If
13290           End Select
13300           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
13310         End If

13320       End If  ' ** DoReport().
13330     End If  ' ** blnContinue.

13340     DoCmd.Hourglass False

13350   End With

EXITP:
13360   Exit Sub

ERRH:
13370   DoCmd.Hourglass False
13380   Select Case ERR.Number
        Case Else
13390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
13400   End Select
13410   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_PurchasedExcel_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_PurchasedExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13520   intShiftDown = (Shift And acShiftMask) > 0
13530   intAltDown = (Shift And acAltMask) > 0
13540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13560     Select Case intRetVal
          Case vbKeyTab
13570       With Me
13580         intRetVal = 0
13590         .cmdAssetPurchasedSold_SoldPreview.SetFocus
13600       End With
13610     Case vbKeyUp
13620       With Me
13630         intRetVal = 0
13640         .cmdAssetPurchasedSold_SoldExcel.SetFocus
13650       End With
13660     Case vbKeyDown
13670       With Me
13680         intRetVal = 0
13690         .cmdAssetPurchasedSold_SoldExcel.SetFocus
13700       End With
13710     Case vbKeyLeft
13720       With Me
13730         intRetVal = 0
13740         .cmdAssetPurchasedSold_PurchasedWord.SetFocus
13750       End With
13760     Case vbKeyRight
13770       With Me
13780         intRetVal = 0
13790         .cmdAssetPurchasedSold_SoldPreview.SetFocus
13800       End With
13810     End Select
13820   End If

        ' ** Shift keys.
13830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13840     Select Case intRetVal
          Case vbKeyTab
13850       With Me
13860         intRetVal = 0
13870         .cmdAssetPurchasedSold_PurchasedWord.SetFocus
13880       End With
13890     End Select
13900   End If

EXITP:
13910   KeyCode = intRetVal
13920   Exit Sub

ERRH:
13930   intRetVal = 0
13940   Select Case ERR.Number
        Case Else
13950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13960   End Select
13970   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_SoldPreview_Click()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_SoldPreview_Click"

        Dim strDocName As String

14010   If DoReport("Sold") = True Then  ' ** Function: Below.
14020     strDocName = "rptPurchasedSold_Sold"
14030     DoCmd.OpenReport strDocName, acViewPreview
14040     DoCmd.Maximize
14050     DoCmd.RunCommand acCmdFitToWindow
14060   End If

EXITP:
14070   Exit Sub

ERRH:
14080   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
14090   Case Else
14100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14110   End Select
14120   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_SoldPreview_KeyDown(KeyCode As Integer, Shift As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_SoldPreview_KeyDown"

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
          Case vbKeyTab
14270       With Me
14280         intRetVal = 0
14290         .cmdAssetPurchasedSold_SoldPrint.SetFocus
14300       End With
14310     Case vbKeyUp
14320       With Me
14330         intRetVal = 0
14340         .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
14350       End With
14360     Case vbKeyDown
14370       With Me
14380         intRetVal = 0
14390         .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
14400       End With
14410     Case vbKeyLeft
14420       With Me
14430         intRetVal = 0
14440         If .cmdAssetPurchasedSold_PurchasedExcel.Enabled = True Then
14450           .cmdAssetPurchasedSold_PurchasedExcel.SetFocus
14460         Else
14470           .cmdAssetPurchasedSold_PurchasedWord.SetFocus
14480         End If
14490       End With
14500     Case vbKeyRight
14510       With Me
14520         intRetVal = 0
14530         .cmdAssetPurchasedSold_SoldPrint.SetFocus
14540       End With
14550     End Select
14560   End If

        ' ** Shift keys.
14570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14580     Select Case intRetVal
          Case vbKeyTab
14590       With Me
14600         intRetVal = 0
14610         If .cmdAssetPurchasedSold_PurchasedExcel.Enabled = True Then
14620           .cmdAssetPurchasedSold_PurchasedExcel.SetFocus
14630         Else
14640           .cmdAssetPurchasedSold_PurchasedWord.SetFocus
14650         End If
14660       End With
14670     End Select
14680   End If

EXITP:
14690   KeyCode = intRetVal
14700   Exit Sub

ERRH:
14710   intRetVal = 0
14720   THAT_PROC = THIS_PROC
14730   That_Erl = Erl
14740   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14750   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_SoldPrint_Click()

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_SoldPrint_Click"

        Dim strDocName As String

14810   If DoReport("Sold") = True Then  ' ** Function: Below.
14820     strDocName = "rptPurchasedSold_Sold"
14830     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
14840       DoCmd.OpenReport strDocName, acViewPreview
14850       DoCmd.Maximize
14860       DoCmd.RunCommand acCmdFitToWindow
14870     Else
            '##GTR_Ref: rptPurchasedSold_Sold
14880       DoCmd.OpenReport strDocName, acViewNormal
14890     End If
14900   Else
14910     blnPrintBoth = False
14920   End If

EXITP:
14930   Exit Sub

ERRH:
14940   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
14950   Case Else
14960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14970   End Select
14980   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_SoldPrint_KeyDown(KeyCode As Integer, Shift As Integer)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_SoldPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15020   intShiftDown = (Shift And acShiftMask) > 0
15030   intAltDown = (Shift And acAltMask) > 0
15040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15060     Select Case intRetVal
          Case vbKeyTab
15070       With Me
15080         intRetVal = 0
15090         .cmdAssetPurchasedSold_SoldWord.SetFocus
15100       End With
15110     Case vbKeyUp
15120       With Me
15130         intRetVal = 0
15140         .cmdAssetPurchasedSold_PurchasedPrint.SetFocus
15150       End With
15160     Case vbKeyDown
15170       With Me
15180         intRetVal = 0
15190         .cmdPrintAll.SetFocus
15200       End With
15210     Case vbKeyLeft
15220       With Me
15230         intRetVal = 0
15240         .cmdAssetPurchasedSold_SoldPreview.SetFocus
15250       End With
15260     Case vbKeyRight
15270       With Me
15280         intRetVal = 0
15290         .cmdAssetPurchasedSold_SoldWord.SetFocus
15300       End With
15310     End Select
15320   End If

        ' ** Shift keys.
15330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15340     Select Case intRetVal
          Case vbKeyTab
15350       With Me
15360         intRetVal = 0
15370         .cmdAssetPurchasedSold_SoldPreview.SetFocus
15380       End With
15390     End Select
15400   End If

EXITP:
15410   KeyCode = intRetVal
15420   Exit Sub

ERRH:
15430   intRetVal = 0
15440   THAT_PROC = THIS_PROC
15450   That_Erl = Erl
15460   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15470   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_SoldWord_Click()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_SoldWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

15510   With Me
15520     If DoReport("Sold") = True Then  ' ** Function: Below.

15530       If IsNull(.UserReportPath) = True Then
15540         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
15550       Else
15560         strRptPath = .UserReportPath
15570       End If

15580       strRpt = "rptPurchasedSoldAssets_Sold"
15590       gdatStartDate = .DateStart
15600       gdatEndDate = .DateEnd
15610       strRptCap = "Assets_Sold_" & _
              Format$(gdatStartDate, "yymmdd") & "_To_" & Format$(gdatEndDate, "yymmdd")

15620       Select Case gblnPrintAll
            Case True
15630         strRptPathFile = strRptPath & LNK_SEP & strRptCap & ".rtf"
15640       Case False
15650         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
15660       End Select

15670       If strRptPathFile <> vbNullString Then
15680         strFile2 = strRptPathFile
15690         Select Case .chkOpenWord
              Case True
15700           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
15710         Case False
15720           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, False
15730         End Select
15740         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
15750       Else
15760         gblnPrintAll = False
15770       End If

15780     Else
15790       gblnPrintAll = False
15800     End If
15810     gdatStartDate = 0
15820     gdatEndDate = 0
15830   End With

EXITP:
15840   Exit Sub

ERRH:
15850   gblnPrintAll = False
15860   Select Case ERR.Number
        Case Else
15870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
15880   End Select
15890   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_SoldWord_KeyDown(KeyCode As Integer, Shift As Integer)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_SoldWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15920   intShiftDown = (Shift And acShiftMask) > 0
15930   intAltDown = (Shift And acAltMask) > 0
15940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15960     Select Case intRetVal
          Case vbKeyTab
15970       With Me
15980         intRetVal = 0
15990         If .cmdAssetPurchasedSold_SoldExcel.Enabled = True Then
16000           .cmdAssetPurchasedSold_SoldExcel.SetFocus
16010         Else
16020           .cmdPrintAll.SetFocus
16030         End If
16040       End With
16050     Case vbKeyUp
16060       With Me
16070         intRetVal = 0
16080         .cmdAssetPurchasedSold_PurchasedWord.SetFocus
16090       End With
16100     Case vbKeyDown
16110       With Me
16120         intRetVal = 0
16130         .cmdAssetPurchasedSold_PurchasedWord.SetFocus
16140       End With
16150     Case vbKeyLeft
16160       With Me
16170         intRetVal = 0
16180         .cmdAssetPurchasedSold_SoldPrint.SetFocus
16190       End With
16200     Case vbKeyRight
16210       With Me
16220         intRetVal = 0
16230         If .cmdAssetPurchasedSold_SoldExcel.Enabled = True Then
16240           .cmdAssetPurchasedSold_SoldExcel.SetFocus
16250         Else
16260           .cmdPrintAll.SetFocus
16270         End If
16280       End With
16290     End Select
16300   End If

        ' ** Shift keys.
16310   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16320     Select Case intRetVal
          Case vbKeyTab
16330       With Me
16340         intRetVal = 0
16350         .cmdAssetPurchasedSold_SoldPrint.SetFocus
16360       End With
16370     End Select
16380   End If

EXITP:
16390   KeyCode = intRetVal
16400   Exit Sub

ERRH:
16410   intRetVal = 0
16420   Select Case ERR.Number
        Case Else
16430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16440   End Select
16450   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_SoldExcel_Click()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_SoldExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

16510   With Me

16520     DoCmd.Hourglass True
16530     DoEvents

16540     blnContinue = True

16550     If EXE_IsRunning("EXCEL.EXE") = True And gblnPrintAll = False Then  ' ** Module Function: modProcessFuncs.
16560       DoCmd.Hourglass False
16570       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
16580       If msgResponse <> vbRetry Then
16590         blnContinue = False
16600       End If
16610     End If

16620     If blnContinue = True Then

16630       DoCmd.Hourglass True
16640       DoEvents

16650       If DoReport("Sold") = True Then  ' ** Procedure: Below.

16660         blnContinue = True

16670         If IsNull(.UserReportPath) = True Then
16680           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
16690         Else
16700           strRptPath = .UserReportPath
16710         End If

16720         strRptCap = "Assets_Sold_" & Format$(.DateStart, "yymmdd") & "_To_" & Format$(.DateEnd, "yymmdd")

              ' ** qryRpt_PurchasedSold_Sold_15, For Export.
16730         strQry = "qryRpt_PurchasedSold_Sold_16"

16740         Select Case gblnPrintAll
              Case True
16750           If strFile1 <> vbNullString Then
16760             strRptPathFile = StringReplace(strFile1, "Purchased_", "Sold_")  ' ** Module Function: modStringFuncs.
16770           Else
16780             strRptPathFile = strRptPath & LNK_SEP & strRptCap & ".xls"
16790           End If
16800         Case False
16810           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
16820         End Select

16830         If strRptPathFile <> vbNullString Then
16840           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
16850             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
16860           End If
16870           DoEvents
16880           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
16890             Kill strRptPathFile
16900           End If
16910           DoEvents
16920           gstrFormQuerySpec = THIS_NAME
16930           Select Case gblnPrintAll
                Case True
16940             DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
16950             strFile2 = strRptPathFile
16960           Case False
16970             DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
16980             If Excel_NameOnly(strRptPathFile, "Assets Sold") = True Then  ' ** Module Function: modExcelFuncs.
16990               Select Case .chkOpenExcel
                    Case True
17000                 If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
17010                   EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
17020                 End If
17030                 OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
17040               Case False
                      ' ** Nothing, that's it.
17050               End Select
17060             End If
17070           End Select
17080           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
17090         End If

17100       End If  ' ** DoReport().
17110     End If  ' ** blnContinue.

17120     DoCmd.Hourglass False

17130   End With

EXITP:
17140   Exit Sub

ERRH:
17150   DoCmd.Hourglass False
17160   Select Case ERR.Number
        Case Else
17170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
17180   End Select
17190   Resume EXITP

End Sub

Private Sub cmdAssetPurchasedSold_SoldExcel_KeyDown(KeyCode As Integer, Shift As Integer)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPurchasedSold_SoldExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17220   intShiftDown = (Shift And acShiftMask) > 0
17230   intAltDown = (Shift And acAltMask) > 0
17240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17260     Select Case intRetVal
          Case vbKeyTab
17270       With Me
17280         intRetVal = 0
17290         .cmdPrintAll.SetFocus
17300       End With
17310     Case vbKeyUp
17320       With Me
17330         intRetVal = 0
17340         .cmdAssetPurchasedSold_PurchasedExcel.SetFocus
17350       End With
17360     Case vbKeyDown
17370       With Me
17380         intRetVal = 0
17390         .cmdAssetPurchasedSold_PurchasedExcel.SetFocus
17400       End With
17410     Case vbKeyLeft
17420       With Me
17430         intRetVal = 0
17440         .cmdAssetPurchasedSold_SoldWord.SetFocus
17450       End With
17460     Case vbKeyRight
17470       With Me
17480         intRetVal = 0
17490         .cmdPrintAll.SetFocus
17500       End With
17510     End Select
17520   End If

        ' ** Shift keys.
17530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17540     Select Case intRetVal
          Case vbKeyTab
17550       With Me
17560         intRetVal = 0
17570         .cmdAssetPurchasedSold_SoldWord.SetFocus
17580       End With
17590     End Select
17600   End If

EXITP:
17610   KeyCode = intRetVal
17620   Exit Sub

ERRH:
17630   intRetVal = 0
17640   Select Case ERR.Number
        Case Else
17650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17660   End Select
17670   Resume EXITP

End Sub

Private Sub cmdPrintAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_box01_MouseMove"

17710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
17720   Exit Sub

ERRH:
17730   Select Case ERR.Number
        Case Else
17740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17750   End Select
17760   Resume EXITP

End Sub

Private Sub cmdPrintAll_GotFocus()

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_GotFocus"

17810   With Me
17820     blnPrintAll_Focus = True
17830     .cmdPrintAll_box01.Visible = True
17840     .cmdPrintAll_box02.Visible = True
17850   End With

EXITP:
17860   Exit Sub

ERRH:
17870   Select Case ERR.Number
        Case Else
17880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17890   End Select
17900   Resume EXITP

End Sub

Private Sub cmdPrintAll_Click()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_Click"

18010   blnPrintBoth = True
18020   cmdAssetPurchasedSold_PurchasedPrint_Click  ' ** Procedure: Above.
18030   DoEvents
18040   If blnPrintBoth = True Then
18050     cmdAssetPurchasedSold_SoldPrint_Click  ' ** Procedure: Above.
18060   End If
18070   blnPrintBoth = False

EXITP:
18080   Exit Sub

ERRH:
18090   blnPrintBoth = False
18100   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
18110   Case Else
18120     THAT_PROC = THIS_PROC
18130     That_Erl = Erl
18140     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18150   End Select
18160   Resume EXITP

End Sub

Private Sub cmdPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18220   intShiftDown = (Shift And acShiftMask) > 0
18230   intAltDown = (Shift And acAltMask) > 0
18240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18260     Select Case intRetVal
          Case vbKeyTab
18270       With Me
18280         intRetVal = 0
18290         .cmdWordAll.SetFocus
18300       End With
18310     Case vbKeyUp
18320       With Me
18330         intRetVal = 0
18340         .cmdAssetPurchasedSold_SoldPrint.SetFocus
18350       End With
18360     Case vbKeyDown
18370       With Me
18380         intRetVal = 0
18390         .cmdAssetPurchasedSold_PurchasedPrint.SetFocus
18400       End With
18410     Case vbKeyLeft
18420       With Me
18430         intRetVal = 0
18440         If .cmdAssetPurchasedSold_SoldExcel.Enabled = True Then
18450           .cmdAssetPurchasedSold_SoldExcel.SetFocus
18460         Else
18470           .cmdAssetPurchasedSold_SoldWord.SetFocus
18480         End If
18490       End With
18500     Case vbKeyRight
18510       With Me
18520         intRetVal = 0
18530         .cmdWordAll.SetFocus
18540       End With
18550     End Select
18560   End If

        ' ** Shift keys.
18570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18580     Select Case intRetVal
          Case vbKeyTab
18590       With Me
18600         intRetVal = 0
18610         If .cmdAssetPurchasedSold_SoldExcel.Enabled = True Then
18620           .cmdAssetPurchasedSold_SoldExcel.SetFocus
18630         Else
18640           .cmdAssetPurchasedSold_SoldWord.SetFocus
18650         End If
18660       End With
18670     End Select
18680   End If

EXITP:
18690   KeyCode = intRetVal
18700   Exit Sub

ERRH:
18710   intRetVal = 0
18720   THAT_PROC = THIS_PROC
18730   That_Erl = Erl
18740   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18750   Resume EXITP

End Sub

Private Sub cmdPrintAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_MouseMove"

18810   With Me
18820     If gblnPrintAll = False Then
18830       .cmdPrintAll_box01.Visible = True
18840       .cmdPrintAll_box02.Visible = True
18850       If blnWordAll_Focus = False Then
18860         .cmdWordAll_box01.Visible = False
18870         .cmdWordAll_box02.Visible = False
18880       End If
18890       If blnExcelAll_Focus = False Then
18900         .cmdExcelAll_box01.Visible = False
18910         .cmdExcelAll_box02.Visible = False
18920       End If
18930     End If
18940   End With

EXITP:
18950   Exit Sub

ERRH:
18960   Select Case ERR.Number
        Case Else
18970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18980   End Select
18990   Resume EXITP

End Sub

Private Sub cmdPrintAll_LostFocus()

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_LostFocus"

19010   With Me
19020     .cmdPrintAll_box01.Visible = False
19030     .cmdPrintAll_box02.Visible = False
19040     blnPrintAll_Focus = False
19050   End With

EXITP:
19060   Exit Sub

ERRH:
19070   Select Case ERR.Number
        Case Else
19080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19090   End Select
19100   Resume EXITP

End Sub

Private Sub cmdWordAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_box01_MouseMove"

19210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
19220   Exit Sub

ERRH:
19230   Select Case ERR.Number
        Case Else
19240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19250   End Select
19260   Resume EXITP

End Sub

Private Sub cmdWordAll_GotFocus()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_GotFocus"

19310   With Me
19320     blnWordAll_Focus = True
19330     .cmdWordAll_box01.Visible = True
19340     .cmdWordAll_box02.Visible = True
19350   End With

EXITP:
19360   Exit Sub

ERRH:
19370   Select Case ERR.Number
        Case Else
19380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19390   End Select
19400   Resume EXITP

End Sub

Private Sub cmdWordAll_Click()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_Click"

19510   With Me

19520     gblnPrintAll = True
19530     strFile1 = vbNullString: strFile2 = vbNullString

19540     cmdAssetPurchasedSold_PurchasedWord_Click  ' ** Procedure: Above.
19550     DoEvents
19560     If gblnPrintAll = True Then
19570       cmdAssetPurchasedSold_SoldWord_Click  ' ** Procedure: Above.
19580     End If

19590     gblnPrintAll = False

19600   End With

EXITP:
19610   Exit Sub

ERRH:
19620   gblnPrintAll = False
19630   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
19640   Case Else
19650     Select Case ERR.Number
          Case Else
19660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
19670     End Select
19680   End Select
19690   Resume EXITP

End Sub

Private Sub cmdWordAll_KeyDown(KeyCode As Integer, Shift As Integer)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19720   intShiftDown = (Shift And acShiftMask) > 0
19730   intAltDown = (Shift And acAltMask) > 0
19740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19760     Select Case intRetVal
          Case vbKeyTab
19770       With Me
19780         intRetVal = 0
19790         If .cmdExcelAll.Enabled = True Then
19800           .cmdExcelAll.SetFocus
19810         Else
19820           .cmdClose.SetFocus
19830         End If
19840       End With
19850     Case vbKeyUp
19860       With Me
19870         intRetVal = 0
19880         .cmdAssetPurchasedSold_SoldWord.SetFocus
19890       End With
19900     Case vbKeyDown
19910       With Me
19920         intRetVal = 0
19930         .cmdAssetPurchasedSold_PurchasedWord.SetFocus
19940       End With
19950     Case vbKeyLeft
19960       With Me
19970         intRetVal = 0
19980         .cmdPrintAll.SetFocus
19990       End With
20000     Case vbKeyRight
20010       With Me
20020         intRetVal = 0
20030         If .cmdExcelAll.Enabled = True Then
20040           .cmdExcelAll.SetFocus
20050         Else
20060           .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
20070         End If
20080       End With
20090     End Select
20100   End If

        ' ** Shift keys.
20110   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20120     Select Case intRetVal
          Case vbKeyTab
20130       With Me
20140         intRetVal = 0
20150         .cmdPrintAll.SetFocus
20160       End With
20170     End Select
20180   End If

        ' ** Ctrl keys.
20190   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
20200     Select Case intRetVal
          Case vbKeyTab
20210       With Me
20220         intRetVal = 0
20230         .chkOpenWord.SetFocus
20240       End With
20250     End Select
20260   End If

EXITP:
20270   KeyCode = intRetVal
20280   Exit Sub

ERRH:
20290   intRetVal = 0
20300   THAT_PROC = THIS_PROC
20310   That_Erl = Erl
20320   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20330   Resume EXITP

End Sub

Private Sub cmdWordAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_MouseMove"

20410   With Me
20420     If gblnPrintAll = False Then
20430       .cmdWordAll_box01.Visible = True
20440       .cmdWordAll_box02.Visible = True
20450       If blnPrintAll_Focus = False Then
20460         .cmdPrintAll_box01.Visible = False
20470         .cmdPrintAll_box02.Visible = False
20480       End If
20490       If blnExcelAll_Focus = False Then
20500         .cmdExcelAll_box01.Visible = False
20510         .cmdExcelAll_box02.Visible = False
20520       End If
20530     End If
20540   End With

EXITP:
20550   Exit Sub

ERRH:
20560   Select Case ERR.Number
        Case Else
20570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20580   End Select
20590   Resume EXITP

End Sub

Private Sub cmdWordAll_LostFocus()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_LostFocus"

20610   With Me
20620     .cmdWordAll_box01.Visible = False
20630     .cmdWordAll_box02.Visible = False
20640     blnWordAll_Focus = False
20650   End With

EXITP:
20660   Exit Sub

ERRH:
20670   Select Case ERR.Number
        Case Else
20680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20690   End Select
20700   Resume EXITP

End Sub

Private Sub chkOpenWord_AfterUpdate()

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_AfterUpdate"

20810   With Me
20820     Select Case .chkOpenWord
          Case True
20830       .chkOpenWord_lbl.FontBold = True
20840       .chkOpenWord_lbl_dim_hi.FontBold = True
20850       .chkOpenWord_lbl2.FontBold = True
20860       .chkOpenWord_lbl2_dim_hi.FontBold = True
20870     Case False
20880       .chkOpenWord_lbl.FontBold = False
20890       .chkOpenWord_lbl_dim_hi.FontBold = False
20900       .chkOpenWord_lbl2.FontBold = False
20910       .chkOpenWord_lbl2_dim_hi.FontBold = False
20920     End Select
20930   End With

EXITP:
20940   Exit Sub

ERRH:
20950   Select Case ERR.Number
        Case Else
20960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20970   End Select
20980   Resume EXITP

End Sub

Private Sub chkOpenWord_KeyDown(KeyCode As Integer, Shift As Integer)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21020   intShiftDown = (Shift And acShiftMask) > 0
21030   intAltDown = (Shift And acAltMask) > 0
21040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21060     Select Case intRetVal
          Case vbKeyTab
21070       With Me
21080         intRetVal = 0
21090         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
21100           .cmdExcelAll.SetFocus
21110         Else
21120           .cmdClose.SetFocus
21130         End If
21140       End With
21150     Case vbKeyUp
21160       With Me
21170         intRetVal = 0
21180         .cmdWordAll.SetFocus
21190       End With
21200     Case vbKeyDown
21210       With Me
21220         intRetVal = 0
21230         .cmdAssetPurchasedSold_PurchasedWord.SetFocus
21240       End With
21250     End Select
21260   End If

        ' ** Shift keys.
21270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21280     Select Case intRetVal
          Case vbKeyTab
21290       With Me
21300         intRetVal = 0
21310         .cmdWordAll.SetFocus
21320       End With
21330     End Select
21340   End If

        ' ** Ctrl keys.
21350   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
21360     Select Case intRetVal
          Case vbKeyTab
21370       With Me
21380         intRetVal = 0
21390         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
21400           .chkOpenExcel.SetFocus
21410         Else
21420           Beep
21430         End If
21440       End With
21450     End Select
21460   End If

EXITP:
21470   KeyCode = intRetVal
21480   Exit Sub

ERRH:
21490   intRetVal = 0
21500   THAT_PROC = THIS_PROC
21510   That_Erl = Erl
21520   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21530   Resume EXITP

End Sub

Private Sub chkOpenWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_MouseMove"

21610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
21620   Exit Sub

ERRH:
21630   Select Case ERR.Number
        Case Else
21640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21650   End Select
21660   Resume EXITP

End Sub

Private Sub cmdExcelAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_box01_MouseMove"

21710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
21720   Exit Sub

ERRH:
21730   Select Case ERR.Number
        Case Else
21740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21750   End Select
21760   Resume EXITP

End Sub

Private Sub cmdExcelAll_GotFocus()

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_GotFocus"

21810   With Me
21820     blnExcelAll_Focus = True
21830     .cmdExcelAll_box01.Visible = True
21840     .cmdExcelAll_box02.Visible = True
21850   End With

EXITP:
21860   Exit Sub

ERRH:
21870   Select Case ERR.Number
        Case Else
21880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21890   End Select
21900   Resume EXITP

End Sub

Private Sub cmdExcelAll_Click()

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_Click"

22010   With Me

22020     DoCmd.Hourglass True
22030     DoEvents

22040     gblnPrintAll = True
22050     strFile1 = vbNullString: strFile2 = vbNullString

          ' ** Excel not opened after export.
22060     cmdAssetPurchasedSold_PurchasedExcel_Click  ' ** Procedure: Above.
22070     DoEvents
22080     If gblnPrintAll = True Then
            ' ** Excel not opened after export.
22090       cmdAssetPurchasedSold_SoldExcel_Click  ' ** Procedure: Above.
22100     End If

22110     If gblnPrintAll = True Then
22120       Select Case .chkOpenExcel
            Case True
22130         DoCmd.Hourglass True
22140         DoEvents
22150         If Excel_NameOnly(strFile1, "Assets Purchased") = True Then  ' ** Module Function: modExcelFuncs.
22160           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22170             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22180           End If
22190           OpenExe strFile1  ' ** Module Function: modShellFuncs.
22200         End If
22210         DoEvents
22220         ForcePause 2  ' ** Module Function: modCodeUtilities.
22230         If Excel_NameOnly(strFile2, "Assets Sold") = True Then  ' ** Module Function: modExcelFuncs.
22240           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22250             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22260           End If
22270           OpenExe strFile2  ' ** Module Function: modShellFuncs.
22280         End If
22290         DoEvents
22300       Case False
              ' ** Nothing, that's it.
22310       End Select
22320     End If

22330     gblnPrintAll = False

22340     DoCmd.Hourglass False

22350   End With

EXITP:
22360   Exit Sub

ERRH:
22370   gblnPrintAll = False
22380   DoCmd.Hourglass False
22390   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
22400   Case Else
22410     Select Case ERR.Number
          Case Else
22420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
22430     End Select
22440   End Select
22450   Resume EXITP

End Sub

Private Sub cmdExcelAll_KeyDown(KeyCode As Integer, Shift As Integer)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22520   intShiftDown = (Shift And acShiftMask) > 0
22530   intAltDown = (Shift And acAltMask) > 0
22540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22560     Select Case intRetVal
          Case vbKeyTab
22570       With Me
22580         intRetVal = 0
22590         .cmdClose.SetFocus
22600       End With
22610     Case vbKeyUp
22620       With Me
22630         intRetVal = 0
22640         .cmdAssetPurchasedSold_SoldExcel.SetFocus
22650       End With
22660     Case vbKeyDown
22670       With Me
22680         intRetVal = 0
22690         .cmdAssetPurchasedSold_PurchasedExcel.SetFocus
22700       End With
22710     Case vbKeyLeft
22720       With Me
22730         intRetVal = 0
22740         .cmdWordAll.SetFocus
22750       End With
22760     Case vbKeyRight
22770       With Me
22780         intRetVal = 0
22790         .cmdAssetPurchasedSold_PurchasedPreview.SetFocus
22800       End With
22810     End Select
22820   End If

        ' ** Shift keys.
22830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22840     Select Case intRetVal
          Case vbKeyTab
22850       With Me
22860         intRetVal = 0
22870         .cmdWordAll.SetFocus
22880       End With
22890     End Select
22900   End If

        ' ** Ctrl keys.
22910   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
22920     Select Case intRetVal
          Case vbKeyTab
22930       With Me
22940         intRetVal = 0
22950         .chkOpenExcel.SetFocus
22960       End With
22970     End Select
22980   End If

EXITP:
22990   KeyCode = intRetVal
23000   Exit Sub

ERRH:
23010   intRetVal = 0
23020   THAT_PROC = THIS_PROC
23030   That_Erl = Erl
23040   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23050   Resume EXITP

End Sub

Private Sub cmdExcelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_MouseMove"

23110   With Me
23120     If gblnPrintAll = False Then
23130       .cmdExcelAll_box01.Visible = True
23140       .cmdExcelAll_box02.Visible = True
23150       If blnPrintAll_Focus = False Then
23160         .cmdPrintAll_box01.Visible = False
23170         .cmdPrintAll_box02.Visible = False
23180       End If
23190       If blnWordAll_Focus = False Then
23200         .cmdWordAll_box01.Visible = False
23210         .cmdWordAll_box02.Visible = False
23220       End If
23230     End If
23240   End With

EXITP:
23250   Exit Sub

ERRH:
23260   Select Case ERR.Number
        Case Else
23270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23280   End Select
23290   Resume EXITP

End Sub

Private Sub cmdExcelAll_LostFocus()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_LostFocus"

23310   With Me
23320     .cmdExcelAll_box01.Visible = False
23330     .cmdExcelAll_box02.Visible = False
23340     blnExcelAll_Focus = False
23350   End With

EXITP:
23360   Exit Sub

ERRH:
23370   Select Case ERR.Number
        Case Else
23380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23390   End Select
23400   Resume EXITP

End Sub

Private Sub chkOpenExcel_AfterUpdate()

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_AfterUpdate"

23510   With Me
23520     Select Case .chkOpenExcel
          Case True
23530       .chkOpenExcel_lbl.FontBold = True
23540       .chkOpenExcel_lbl_dim_hi.FontBold = True
23550       .chkOpenExcel_lbl2.FontBold = True
23560       .chkOpenExcel_lbl2_dim_hi.FontBold = True
23570     Case False
23580       .chkOpenExcel_lbl.FontBold = False
23590       .chkOpenExcel_lbl_dim_hi.FontBold = False
23600       .chkOpenExcel_lbl2.FontBold = False
23610       .chkOpenExcel_lbl2_dim_hi.FontBold = False
23620     End Select
23630   End With

EXITP:
23640   Exit Sub

ERRH:
23650   Select Case ERR.Number
        Case Else
23660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23670   End Select
23680   Resume EXITP

End Sub

Private Sub chkOpenExcel_KeyDown(KeyCode As Integer, Shift As Integer)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23720   intShiftDown = (Shift And acShiftMask) > 0
23730   intAltDown = (Shift And acAltMask) > 0
23740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23760     Select Case intRetVal
          Case vbKeyTab
23770       With Me
23780         intRetVal = 0
23790         .cmdClose.SetFocus
23800       End With
23810     Case vbKeyUp
23820       With Me
23830         intRetVal = 0
23840         .cmdExcelAll.SetFocus
23850       End With
23860     Case vbKeyDown
23870       With Me
23880         intRetVal = 0
23890         .cmdAssetPurchasedSold_PurchasedExcel.SetFocus
23900       End With
23910     End Select
23920   End If

        ' ** Shift keys.
23930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23940     Select Case intRetVal
          Case vbKeyTab
23950       With Me
23960         intRetVal = 0
23970         .cmdExcelAll.SetFocus
23980       End With
23990     End Select
24000   End If

EXITP:
24010   KeyCode = intRetVal
24020   Exit Sub

ERRH:
24030   intRetVal = 0
24040   THAT_PROC = THIS_PROC
24050   That_Erl = Erl
24060   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24070   Resume EXITP

End Sub

Private Sub chkOpenExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_MouseMove"

24110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
24120   Exit Sub

ERRH:
24130   Select Case ERR.Number
        Case Else
24140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24150   End Select
24160   Resume EXITP

End Sub

Public Function DoReport(strMode As String) As Boolean

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

24210   blnRetVal = True

24220   With Me
24230     If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
24240       .DateEnd = Date
24250       DoEvents
24260     End If
24270     If IsNull(.DateStart) Then
24280       blnRetVal = False
24290       MsgBox "The start date must be prior to or the same as the end date.", vbInformation + vbOKOnly, "Entry Required"
24300       .DateStart.SetFocus
24310     Else
24320       If CDate(.DateStart) > CDate(.DateEnd) Then
24330         blnRetVal = False
24340         MsgBox "The start date must be prior to or the same as the end date.", vbInformation + vbOKOnly, "Invalid Date"
24350         .DateStart.SetFocus
24360       End If
24370     End If
24380   End With

24390   If blnRetVal = True Then
24400     gstrFormQuerySpec = THIS_NAME
24410     Set dbs = CurrentDb
24420     With dbs
24430       Select Case strMode
            Case "Purchased"
              ' ** Ledger, linked to qryRpt_PurchasedSold_01 (MasterAsset, with totdesc), by specified FormRef(); Purchase.
24440         Set qdf = .QueryDefs("qryRpt_PurchasedSold_Purchased_01")
24450       Case "Sold"
              ' ** Ledger, linked to qryRpt_PurchasedSold_01 (MasterAsset, with totdesc), by specified FormRef(); Sold.
24460         Set qdf = .QueryDefs("qryRpt_PurchasedSold_Sold_01")
24470       End Select
24480       Set rst = qdf.OpenRecordset
24490       With rst
24500         If .BOF = True And .EOF = True Then
24510           blnRetVal = False
24520           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
24530         End If
24540         .Close
24550       End With
24560       .Close
24570     End With
24580   End If

EXITP:
24590   Set rst = Nothing
24600   Set qdf = Nothing
24610   Set dbs = Nothing
24620   DoReport = blnRetVal
24630   Exit Function

ERRH:
24640   blnRetVal = False
24650   Select Case ERR.Number
        Case Else
24660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24670   End Select
24680   Resume EXITP

End Function
