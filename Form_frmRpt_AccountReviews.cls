VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_AccountReviews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_AccountReviews"

'VGC 07/15/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, sets options, then stops.
'ARROWS OR HIGHLIGHTS?

' ** Shortcut Alt keys responsive from this form:
' **   Account Profiles: A {chkAccountProfiles}
' **   Close:            C {cmdClose}
' **   Admin Officer:    D {chkAdministrator}
' **   Excel:            E {cmdExcel}
' **   For:              F {cmbMonth}
' **   Word:             O {cmdWord}
' **   Specify All:      L {opgAccountReview_optAll}
' **   Specify Month:    M {opgAccountReview_optSpecified}
' **   Print:            P {cmdPrint}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

'rev_month: IIf([revjan]=True,'jan',IIf([revfeb]=True,'feb',IIf([revmar]=True,'mar',IIf([revapr]=True,'apr',IIf([revmay]=True,'may',IIf([revjun]=True,'jun',IIf([revjul]=True,'jul',IIf([revaug]=True,'aug',IIf([revsep]=True,'sep',IIf([revoct]=True,'oct',IIf([revnov]=True,'nov',IIf([revdec]=True,'dec',''))))))))))))

' ** cmbMonth combo box constants:
Private Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
Private Const CBX_MON_NAME  As Integer = 1  ' ** month_name
Private Const CBX_MON_SHORT As Integer = 2  ' ** month_short
'Private Const CBX_MON_CNT   As Integer = 3  ' ** cnt

Private blnIsWord As Boolean, blnIsExcel As Boolean
Private lngFiles As Long, arr_varFile() As Variant
' ** Array: arr_varFile().
Private Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const F_FNAM As Integer = 0
Private Const F_TYP1 As Integer = 1
Private Const F_TYP2 As Integer = 2

Private blnGoingToReport As Boolean, blnWasGTR As Boolean ', strGTR_Set As String, strGTR_Pref As String
Private strCallingForm As String, lngTpp As Long
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = .OpenArgs

150           blnWasGTR = gblnGoToReport

160           gstrFormQuerySpec = THIS_NAME
170           blnIsWord = False: blnIsExcel = False

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

      #If NoExcel Then
200           .cmdExcel.Enabled = False
      #Else
210           .cmdExcel.Enabled = True
      #End If

220           opgAccountReview_AfterUpdate  ' ** Procedure: Below.
230           chkAccountProfiles_AfterUpdate  ' ** Procedure: Below.
240           chkAdministrator_AfterUpdate  ' ** Procedure: Below.
250           chkPageOf_AfterUpdate  ' ** Procedure: Below.

260           If .cmbMonth.Enabled = True Then
270             .cmbMonth.SetFocus
280           Else
290             .opgAccountReview.SetFocus
300           End If

310         Else
320           Cancel = -1
330         End If
340       Else
350         Cancel = -1
360       End If
370     End With

380     If Cancel = -1 Then
390       Beep
400       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
410     End If

420     DoCmd.Hourglass False

EXITP:
430     Exit Sub

ERRH:
440     DoCmd.Hourglass False
450     Select Case ERR.Number
        Case Else
460       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
470     End Select
480     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me
520       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.
530     End With

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         intRetVal = 0
680         cmdClose_Click  ' ** Procedure: Below.
690       End Select
700     End If

        ' ** Alt keys.
710     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
720       Select Case intRetVal
          Case vbKeyX
730         intRetVal = 0
740         cmdClose_Click  ' ** Procedure: Below.
750       End Select
760     End If

        ' ** Ctrl keys.
770     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
780       Select Case intRetVal
          Case vbKeyG
790         With Me
800           intRetVal = 0
810           If .chkPageOf.Enabled = True Then
820             .chkPageOf.SetFocus
830             .chkPageOf = (Not .chkPageOf)
840             chkPageOf_AfterUpdate  ' ** Procedure: Below.
850           End If
860         End With
870       End Select
880     End If

        ' ** Ctrl-Shift keys.
890     If intCtrlDown And (Not intAltDown) And intShiftDown Then
900       Select Case intRetVal
          Case vbKeyF
910         With Me
920           intRetVal = 0
930           .FocusHolder.SetFocus
940         End With
950       End Select
960     End If

EXITP:
970     KeyCode = intRetVal
980     Exit Sub

ERRH:
990     intRetVal = 0
1000    THAT_PROC = THIS_PROC
1010    That_Erl = Erl
1020    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1030    Resume EXITP

End Sub

Private Sub Form_Timer()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1110    With Me
1120      .TimerInterval = 0&
1130      If gblnSetFocus = True Then
1140        gblnSetFocus = False
1150        strFormName = vbNullString
1160  On Error Resume Next
1170        Do Until Screen.ActiveForm.Name = strFormName
1180          DoCmd.SelectObject acForm, THIS_NAME, False
1190          strFormName = Screen.ActiveForm.Name
1200        Loop
1210  On Error GoTo ERRH
1220      End If
1230      If gblnGoToReport = True Then
1240        Select Case blnGoingToReport
            Case True
1250          .GoToReport_arw_ar_admin_img.Visible = False
1260          If .cmdPrint.Enabled = True Then
1270            .cmdPrint.SetFocus
1280          Else
1290            .cmdClose.SetFocus
1300          End If
1310          DoEvents
1320          blnGoingToReport = False
1330          gblnGoToReport = False
1340          Beep
1350          DoCmd.Hourglass False
1360          DoEvents
1370        Case False
1380          DoCmd.Hourglass True  ' ** Make sure it's still running.
1390          DoEvents
1400          Select Case garr_varGoToReport(GTR_RNAM)
              Case "rptAccountReviews_Admin"
1410            blnGoingToReport = True
1420            .GoToReport_arw_ar_admin_img.Visible = True
1430            DoEvents
1440            If .chkAdministrator = False Then
1450              .chkAdministrator = True
1460              chkAdministrator_AfterUpdate  ' ** Procedure: Below.
1470              DoEvents
1480            End If
1490            .TimerInterval = (GTR_WAIT * 2&)
1500          Case "rptAccountReviews"
1510            If .cmdPrint.Enabled = True Then
1520              .cmdPrint.SetFocus
1530            Else
1540              .cmdClose.SetFocus
1550            End If
1560            DoEvents
1570            If .chkAdministrator = True Then
1580              .chkAdministrator = False
1590              chkAdministrator_AfterUpdate  ' ** Procedure: Below.
1600              DoEvents
1610            End If
1620            blnGoingToReport = False
1630            gblnGoToReport = False
1640            Beep
1650            DoCmd.Hourglass False
1660            DoEvents
1670          End Select
1680        End Select
1690      Else
1700        DoCmd.Hourglass False
1710      End If
1720    End With

EXITP:
1730    Exit Sub

ERRH:
1740    DoCmd.Hourglass False
1750    Select Case ERR.Number
        Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1810    If DataErr <> 0 Then
1820      Select Case DataErr
          Case 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
1830      Case 3162
1840        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
1850        DoCmd.CancelEvent
1860      Case 3129
1870        MsgBox "You must select a valid month to continue.", vbInformation + vbOKOnly, "Invalid Month"
1880        Me.cmbMonth.SetFocus
1890      Case Else
1900        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form error"
1910        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1920      End Select
1930    End If

EXITP:
1940    Exit Sub

ERRH:
1950    Select Case ERR.Number
        Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2010    DoCmd.Hourglass True
2020    DoEvents

        ' ** Make sure all reports are closed.
2030    Do While Reports.Count > 0
2040      DoCmd.Close acReport, Reports(0).Name
2050    Loop

2060    Select Case blnWasGTR
        Case True
2070      blnWasGTR = False
2080    Case False
          ' ** Don't save settings imposed by GoToReport!
2090      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
2100    End Select

2110    gstrFormQuerySpec = vbNullString
2120    gblnSetFocus = True
2130    gstrReturningForm = THIS_NAME
2140    Select Case strCallingForm
        Case "frmMenu_Report"
2150      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2160    End Select

EXITP:
2170    Exit Sub

ERRH:
2180    Select Case ERR.Number
        Case Else
2190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2200    End Select
2210    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2310    DoCmd.Hourglass True
2320    DoEvents

2330    DoCmd.Close acForm, THIS_NAME

EXITP:
2340    Exit Sub

ERRH:
2350    Select Case ERR.Number
        Case Else
2360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2370    End Select
2380    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab
2470        With Me
2480          intRetVal = 0
2490          .opgAccountReview.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab
2550        With Me
2560          intRetVal = 0
2570          If .cmdExcel.Enabled = True Then
2580            .cmdExcel.SetFocus
2590          Else
2600            .cmdWord.SetFocus
2610          End If
2620        End With
2630      End Select
2640    End If

EXITP:
2650    KeyCode = intRetVal
2660    Exit Sub

ERRH:
2670    intRetVal = 0
2680    THAT_PROC = THIS_PROC
2690    That_Erl = Erl
2700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2710    Resume EXITP

End Sub

Public Sub opgAccountReview_AfterUpdate()
' ** Public for GoToReport.

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountReview_AfterUpdate"

2810    With Me
2820      Select Case .opgAccountReview
          Case .opgAccountReview_optSpecified.OptionValue
2830        .opgAccountReview_optSpecified_lbl.FontBold = True
2840        .opgAccountReview_optAll_lbl.FontBold = False
2850        .opgAccountReview_optSpecified_lbl_box.Visible = True
2860        .opgAccountReview_optAll_lbl_box.Visible = False
2870        .cmbMonth.Enabled = True
2880        .cmbMonth.BorderColor = CLR_LTBLU2
2890        .cmbMonth.BackStyle = acBackStyleNormal
2900      Case .opgAccountReview_optAll.OptionValue
2910        .opgAccountReview_optSpecified_lbl.FontBold = False
2920        .opgAccountReview_optAll_lbl.FontBold = True
2930        .opgAccountReview_optSpecified_lbl_box.Visible = False
2940        .opgAccountReview_optAll_lbl_box.Visible = True
2950        .cmbMonth.Enabled = False
2960        .cmbMonth.BorderColor = WIN_CLR_DISR
2970        .cmbMonth.BackStyle = acBackStyleTransparent
2980      End Select
2990    End With

EXITP:
3000    Exit Sub

ERRH:
3010    THAT_PROC = THIS_PROC
3020    That_Erl = Erl
3030    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3040    Resume EXITP

End Sub

Private Sub opgAccountReview_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountReview_optSpecified_KeyDown"

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
3190          If .cmbMonth.Enabled = True Then
3200            .cmbMonth.SetFocus
3210          ElseIf .cmdPreview.Enabled = True Then
3220            .cmdPreview.SetFocus
3230          Else
3240            .cmdPrint.SetFocus
3250          End If
3260        End With
3270      End Select
3280    End If

        ' ** Shift keys.
3290    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3300      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3310        With Me
3320          intRetVal = 0
3330          .cmdClose.SetFocus
3340        End With
3350      End Select
3360    End If

        ' ** Ctrl keys.
3370    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3380      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3390        With Me
3400          intRetVal = 0
3410          .chkAccountProfiles.SetFocus
3420        End With
3430      End Select
3440    End If

EXITP:
3450    KeyCode = intRetVal
3460    Exit Sub

ERRH:
3470    intRetVal = 0
3480    THAT_PROC = THIS_PROC
3490    That_Erl = Erl
3500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3510    Resume EXITP

End Sub

Private Sub opgAccountReview_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountReview_optAll_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
3670        With Me
3680          intRetVal = 0
3690          If .cmbMonth.Enabled = True Then
3700            .cmbMonth.SetFocus
3710          ElseIf .cmdPreview.Enabled = True Then
3720            .cmdPreview.SetFocus
3730          Else
3740            .cmdPrint.SetFocus
3750          End If
3760        End With
3770      End Select
3780    End If

        ' ** Shift keys.
3790    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3800      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3810        With Me
3820          intRetVal = 0
3830          .cmdClose.SetFocus
3840        End With
3850      End Select
3860    End If

        ' ** Ctrl keys.
3870    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3880      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3890        With Me
3900          intRetVal = 0
3910          .chkAccountProfiles.SetFocus
3920        End With
3930      End Select
3940    End If

EXITP:
3950    KeyCode = intRetVal
3960    Exit Sub

ERRH:
3970    intRetVal = 0
3980    THAT_PROC = THIS_PROC
3990    That_Erl = Erl
4000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4010    Resume EXITP

End Sub

Public Sub chkAccountProfiles_AfterUpdate()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkAccountProfiles_AfterUpdate"

4110    With Me
4120      Select Case .chkAccountProfiles
          Case True
4130        .chkAccountProfiles_lbl.FontBold = True
4140        .cmdPreview.Enabled = False
4150      Case False
4160        .chkAccountProfiles_lbl.FontBold = False
4170        .cmdPreview.Enabled = True
4180      End Select
4190    End With

EXITP:
4200    Exit Sub

ERRH:
4210    THAT_PROC = THIS_PROC
4220    That_Erl = Erl
4230    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4240    Resume EXITP

End Sub

Private Sub chkAccountProfiles_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkAccountProfiles_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4320    intShiftDown = (Shift And acShiftMask) > 0
4330    intAltDown = (Shift And acAltMask) > 0
4340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4370        With Me
4380          intRetVal = 0
4390          .chkAdministrator.SetFocus
4400        End With
4410      End Select
4420    End If

        ' ** Shift keys.
4430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4450        With Me
4460          intRetVal = 0
4470          .opgAccountReview.SetFocus
4480        End With
4490      End Select
4500    End If

EXITP:
4510    KeyCode = intRetVal
4520    Exit Sub

ERRH:
4530    intRetVal = 0
4540    THAT_PROC = THIS_PROC
4550    That_Erl = Erl
4560    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4570    Resume EXITP

End Sub

Public Sub chkAdministrator_AfterUpdate()
' ** Public for GoToReport.

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkAdministrator_AfterUpdate"

4610    With Me
4620      Select Case .chkAdministrator
          Case True
4630        .chkAdministrator_lbl.FontBold = True
4640      Case False
4650        .chkAdministrator_lbl.FontBold = False
4660      End Select
4670    End With

EXITP:
4680    Exit Sub

ERRH:
4690    THAT_PROC = THIS_PROC
4700    That_Erl = Erl
4710    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4720    Resume EXITP

End Sub

Private Sub chkAdministrator_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkAdministrator_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
4870        With Me
4880          intRetVal = 0
4890          If .cmbMonth.Enabled = True Then
4900            .cmbMonth.SetFocus
4910          ElseIf .cmdPreview.Enabled = True Then
4920            .cmdPreview.SetFocus
4930          Else
4940            .cmdPrint.SetFocus
4950          End If
4960        End With
4970      End Select
4980    End If

        ' ** Shift keys.
4990    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5000      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5010        With Me
5020          intRetVal = 0
5030          .chkAccountProfiles.SetFocus
5040        End With
5050      End Select
5060    End If

EXITP:
5070    KeyCode = intRetVal
5080    Exit Sub

ERRH:
5090    intRetVal = 0
5100    THAT_PROC = THIS_PROC
5110    That_Erl = Erl
5120    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5130    Resume EXITP

End Sub

Private Sub cmbMonth_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbMonth_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5270        With Me
5280          intRetVal = 0
5290          If .cmdPreview.Enabled = True Then
5300            .cmdPreview.SetFocus
5310          ElseIf .cmdPrint.Enabled = True Then
5320            .cmdPrint.SetFocus
5330          Else
5340            .cmdClose.SetFocus
5350          End If
5360        End With
5370      End Select
5380    End If

        ' ** Shift keys.
5390    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5400      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5410        With Me
5420          intRetVal = 0
5430          .opgAccountReview.SetFocus
5440        End With
5450      End Select
5460    End If

        ' ** Ctrl keys.
5470    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5480      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5490        With Me
5500          intRetVal = 0
5510          .chkPageOf.SetFocus
5520        End With
5530      End Select
5540    End If

        ' ** Ctrl-Shift keys.
5550    If intCtrlDown And (Not intAltDown) And intShiftDown Then
5560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5570        With Me
5580          intRetVal = 0
5590          .chkAdministrator.SetFocus
5600        End With
5610      End Select
5620    End If

EXITP:
5630    KeyCode = intRetVal
5640    Exit Sub

ERRH:
5650    intRetVal = 0
5660    THAT_PROC = THIS_PROC
5670    That_Erl = Erl
5680    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5690    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

5710    With Me
5720      Select Case .chkPageOf
          Case True
5730        .chkPageOf_lbl.FontBold = True
5740      Case False
5750        .chkPageOf_lbl.FontBold = False
5760      End Select
5770    End With

EXITP:
5780    Exit Sub

ERRH:
5790    THAT_PROC = THIS_PROC
5800    That_Erl = Erl
5810    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5820    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

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
5990          If .cmdPreview.Enabled = True Then
6000            .cmdPreview.SetFocus
6010          ElseIf .cmdPrint.Enabled = True Then
6020            .cmdPrint.SetFocus
6030          Else
6040            .cmdClose.SetFocus
6050          End If
6060        End With
6070      End Select
6080    End If

        ' ** Shift keys.
6090    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6100      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6110        With Me
6120          intRetVal = 0
6130          .cmbMonth.SetFocus
6140        End With
6150      End Select
6160    End If

EXITP:
6170    KeyCode = intRetVal
6180    Exit Sub

ERRH:
6190    intRetVal = 0
6200    THAT_PROC = THIS_PROC
6210    That_Erl = Erl
6220    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6230    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

6310    If DoReport(False) = True Then  ' ** Procedure: Below.
6320      gstrFormQuerySpec = THIS_NAME
6330      With Me
6340        Select Case .opgAccountReview
            Case .opgAccountReview_optSpecified.OptionValue
6350          Select Case .chkAdministrator
              Case True
                'qryRpt_AccountReview_01
6360            strDocName = "rptAccountReviews_Admin"
6370          Case False
                'qryRpt_AccountReview_01
6380            strDocName = "rptAccountReviews"
6390          End Select
6400        Case .opgAccountReview_optAll.OptionValue
6410          Select Case .chkAdministrator
              Case True
                'qryRpt_AccountReview_16
6420            strDocName = "rptAccountReviews_Admin_All"
6430          Case False
                'qryRpt_AccountReview_15
6440            strDocName = "rptAccountReviews_All"
6450          End Select
6460        End Select
6470      End With
6480      DoCmd.OpenReport strDocName, acViewPreview
6490      DoCmd.Maximize
6500      DoCmd.RunCommand acCmdFitToWindow
6510    End If

EXITP:
6520    Exit Sub

ERRH:
6530    THAT_PROC = THIS_PROC
6540    That_Erl = Erl
6550    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6560    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6620    intShiftDown = (Shift And acShiftMask) > 0
6630    intAltDown = (Shift And acAltMask) > 0
6640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab
6670        With Me
6680          intRetVal = 0
6690          .cmdPrint.SetFocus
6700        End With
6710      End Select
6720    End If

        ' ** Shift keys.
6730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6740      Select Case intRetVal
          Case vbKeyTab
6750        With Me
6760          intRetVal = 0
6770          If .cmbMonth.Enabled = True Then
6780            .cmbMonth.SetFocus
6790          Else
6800            .opgAccountReview.SetFocus
6810          End If
6820        End With
6830      End Select
6840    End If

        ' ** Ctrl-Shift keys.
6850    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6860      Select Case intRetVal
          Case vbKeyTab
6870        With Me
6880          intRetVal = 0
6890          .chkPageOf.SetFocus
6900        End With
6910      End Select
6920    End If

EXITP:
6930    KeyCode = intRetVal
6940    Exit Sub

ERRH:
6950    intRetVal = 0
6960    THAT_PROC = THIS_PROC
6970    That_Erl = Erl
6980    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6990    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

7010    If DoReport(False) = True Then  ' ** Procedure: Below.
7020      gstrFormQuerySpec = THIS_NAME
7030      With Me
7040        Select Case .opgAccountReview
            Case .opgAccountReview_optSpecified.OptionValue
7050          Select Case .chkAdministrator
              Case True
7060            strDocName = "rptAccountReviews_Admin"
7070          Case False
7080            strDocName = "rptAccountReviews"
7090          End Select
7100        Case .opgAccountReview_optAll.OptionValue
7110          Select Case .chkAdministrator
              Case True
7120            strDocName = "rptAccountReviews_Admin_All"
7130          Case False
7140            strDocName = "rptAccountReviews_All"
7150          End Select
7160        End Select
7170        If gblnDev_Debug = True Then 'Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
7180          DoCmd.OpenReport strDocName, acViewPreview
7190          DoCmd.Maximize
7200          DoCmd.RunCommand acCmdFitToWindow
7210        Else
              '##GTR_Ref: rptAccountReviews
              '##GTR_Ref: rptAccountReviews_All
              '##GTR_Ref: rptAccountReviews_Admin
              '##GTR_Ref: rptAccountReviews_Admin_All
7220          DoCmd.OpenReport strDocName, acViewNormal
7230          If .chkAccountProfiles = True Then
7240            DoReport True  ' ** Function: Below.
7250          End If
7260        End If
7270      End With
7280    End If

EXITP:
7290    Exit Sub

ERRH:
7300    THAT_PROC = THIS_PROC
7310    That_Erl = Erl
7320    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7330    Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

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
          Case vbKeyTab
7470        With Me
7480          intRetVal = 0
7490          .cmdWord.SetFocus
7500        End With
7510      End Select
7520    End If

        ' ** Shift keys.
7530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7540      Select Case intRetVal
          Case vbKeyTab
7550        With Me
7560          intRetVal = 0
7570          If .cmdPreview.Enabled = True Then
7580            .cmdPreview.SetFocus
7590          ElseIf .cmbMonth.Enabled = True Then
7600            .cmbMonth.SetFocus
7610          Else
7620            .opgAccountReview.SetFocus
7630          End If
7640        End With
7650      End Select
7660    End If

EXITP:
7670    KeyCode = intRetVal
7680    Exit Sub

ERRH:
7690    intRetVal = 0
7700    THAT_PROC = THIS_PROC
7710    That_Erl = Erl
7720    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7730    Resume EXITP

End Sub

Private Sub cmdWord_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String, strMon As String
        Dim lngX As Long, lngE As Long

7810    With Me
7820      blnIsWord = True
7830      If DoReport(False) = True Then  ' ** Function: Below.

7840        If IsNull(.UserReportPath) = True Then
7850          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
7860        Else
7870          strRptPath = .UserReportPath
7880        End If

7890        Select Case .opgAccountReview
            Case .opgAccountReview_optSpecified.OptionValue
7900          Select Case .chkAdministrator
              Case True
7910            strRpt = "rptAccountReviews_Admin"
7920          Case False
7930            strRpt = "rptAccountReviews"
7940          End Select
7950          strMon = FormatProperCase(.cmbMonth.Column(CBX_MON_SHORT))  ' ** Module Function: modStringFuncs.
7960          strRptCap = "AccountReview_" & strMon & "_" & Format$(Date, "yyyymmdd")
7970        Case .opgAccountReview_optAll.OptionValue
7980          Select Case .chkAdministrator
              Case True
7990            strRpt = "rptAccountReviews_Admin_All"
8000          Case False
8010            strRpt = "rptAccountReviews_All"
8020          End Select
8030          strRptCap = "AccountReview_All_" & Format$(Date, "yyyymmdd")
8040        End Select

8050        strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

8060        If strRptPathFile <> vbNullString Then

8070          lngFiles = 0&
8080          ReDim arr_varFile(F_ELEMS, 0)

8090          Select Case .chkAccountProfiles
              Case True
8100            lngFiles = lngFiles + 1&
8110            lngE = lngFiles - 1&
8120            ReDim Preserve arr_varFile(F_ELEMS, lngE)
8130            arr_varFile(F_FNAM, lngE) = strRptPathFile
8140            arr_varFile(F_TYP1, lngE) = "Word"
8150            arr_varFile(F_TYP2, lngE) = "Review"
8160            DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, False
8170          Case False
8180            DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
8190          End Select
8200          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

8210          If .chkAccountProfiles = True Then
8220            DoReport True  ' ** Function: Below.
8230            DoEvents
8240            For lngX = 0& To (lngFiles - 1&)
                  ' ** Can I feed it multiple files in one string?
8250              OpenExe arr_varFile(F_FNAM, lngX)  ' ** Module Function: modShellFuncs.
8260            Next
8270          End If

8280        End If  ' ** vbNullString.

8290      End If  ' ** DoReport().
8300      blnIsWord = False
8310    End With

EXITP:
8320    Exit Sub

ERRH:
8330    Select Case ERR.Number
        Case Else
8340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8350    End Select
8360    Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

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
          Case vbKeyTab
8470        With Me
8480          intRetVal = 0
8490          If .cmdExcel.Enabled = True Then
8500            .cmdExcel.SetFocus
8510          Else
8520            .cmdClose.SetFocus
8530          End If
8540        End With
8550      End Select
8560    End If

        ' ** Shift keys.
8570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8580      Select Case intRetVal
          Case vbKeyTab
8590        With Me
8600          intRetVal = 0
8610          .cmdPrint.SetFocus
8620        End With
8630      End Select
8640    End If

EXITP:
8650    KeyCode = intRetVal
8660    Exit Sub

ERRH:
8670    intRetVal = 0
8680    THAT_PROC = THIS_PROC
8690    That_Erl = Erl
8700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8710    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim strMon As String
        Dim lngX As Long, lngE As Long

8810    With Me
8820      blnIsExcel = True
8830      If DoReport(False) = True Then  ' ** Function: Below.

8840        If IsNull(.UserReportPath) = True Then
8850          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
8860        Else
8870          strRptPath = .UserReportPath
8880        End If

8890        Select Case .opgAccountReview
            Case .opgAccountReview_optSpecified.OptionValue
8900          Select Case .chkAdministrator
              Case True
                ' ** qryRpt_AccountReview_20_02_11 (xx), specified, w/ admin, For Export.
8910            strQry = "qryRpt_AccountReview_20_02_12"
8920          Case False
                ' ** qryRpt_AccountReview_20_01_08, specified, For Export.
8930            strQry = "qryRpt_AccountReview_20_01_09"
8940          End Select
8950          strMon = FormatProperCase(.cmbMonth.Column(CBX_MON_SHORT))  ' ** Module Function: modStringFuncs.
8960          strRptCap = "AccountReview_" & strMon & "_" & Format$(Date, "yyyymmdd")
8970        Case .opgAccountReview_optAll.OptionValue
8980          Select Case .chkAdministrator
              Case True
                ' ** qryRpt_AccountReview_20_04_12 (xx), all, w/ admin, For Export.
8990            strQry = "qryRpt_AccountReview_20_04_13"
9000          Case False
                ' ** qryRpt_AccountReview_20_03_10, all (xx), For Export.
9010            strQry = "qryRpt_AccountReview_20_03_11"
9020          End Select
9030          strRptCap = "AccountReview_All_" & Format$(Date, "yyyymmdd")
9040        End Select

9050        strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

9060        If strRptPathFile <> vbNullString Then

9070          gstrFormQuerySpec = THIS_NAME

9080          lngFiles = 0&
9090          ReDim arr_varFile(F_ELEMS, 0)

9100          Select Case .chkAccountProfiles
              Case True
9110            lngFiles = lngFiles + 1&
9120            lngE = lngFiles - 1&
9130            ReDim Preserve arr_varFile(F_ELEMS, lngE)
9140            arr_varFile(F_FNAM, lngE) = strRptPathFile
9150            arr_varFile(F_TYP1, lngE) = "Excel"
9160            arr_varFile(F_TYP2, lngE) = "Review"
9170            DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
9180          Case False
9190            DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
9200          End Select

9210          If Excel_NameOnly(strRptPathFile, "Account Reviews") = True Then  ' ** Module Function: modExcelFuncs.

9220            If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
9230              EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
9240            End If
9250            If .chkAccountProfiles = False Then
9260              OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
9270            End If
9280            .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

9290            If .chkAccountProfiles = True Then
9300              DoReport True  ' ** Function: Below.
9310              DoEvents
9320              For lngX = 0& To (lngFiles - 1&)
                    ' ** Can I feed it multiple files in one string?
9330                OpenExe arr_varFile(F_FNAM, lngX)  ' ** Module Function: modShellFuncs.
9340              Next
9350            End If

9360          End If  ' ** Excel_NameOnly().

9370        End If  ' ** vbNullString.

9380      End If  ' ** DoReport().
9390      blnIsExcel = True
9400    End With

EXITP:
9410    Exit Sub

ERRH:
9420    Select Case ERR.Number
        Case Else
9430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9440    End Select
9450    Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

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
          Case vbKeyTab
9570        With Me
9580          intRetVal = 0
9590          .cmdClose.SetFocus
9600        End With
9610      End Select
9620    End If

        ' ** Shift keys.
9630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9640      Select Case intRetVal
          Case vbKeyTab
9650        With Me
9660          intRetVal = 0
9670          .cmdWord.SetFocus
9680        End With
9690      End Select
9700    End If

EXITP:
9710    KeyCode = intRetVal
9720    Exit Sub

ERRH:
9730    intRetVal = 0
9740    THAT_PROC = THIS_PROC
9750    That_Erl = Erl
9760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9770    Resume EXITP

End Sub

Private Function DoReport(Optional varProfiles As Variant) As Boolean
' ** Run first for validation, then second for printing profiles, if chosen.
' ** cmbMonth:
' **   Col 0 : CBX_MON_ID    : Id/Month number.
' **   Col 1 : CBX_MON_NAME  : Month name.
' **   Col 2 : CBX_MON_SHORT : Month 3-letter abbreviation.
' **   Col 3 : CBX_MON_CNT   : Count.

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRpt As String, strRptCap As String, strRptPath As String, strRptPathFile As String
        Dim strQry As String
        Dim blnProfiles As Boolean
        Dim lngRecs As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

9810    With Me

9820      blnRetVal = True

9830      If IsMissing(varProfiles) = True Then
9840        blnProfiles = False
9850      Else
9860        blnProfiles = CBool(varProfiles)
9870      End If

9880      Select Case blnProfiles
          Case True
            ' ** Print profiles.

9890        gstrFormQuerySpec = THIS_NAME

9900        Set dbs = CurrentDb

9910        Set qdf = dbs.QueryDefs("qryRpt_AccountReview_01")
9920        Set rst = qdf.OpenRecordset
9930        rst.MoveLast
9940        lngRecs = rst.RecordCount
9950        rst.MoveFirst
9960        For lngX = 1& To lngRecs

9970          Select Case gblnSeparateCheckingAccounts
              Case True
                ' ** Account, with add'l fields, by specified FormRef('AccountNo').
                ' ** qryRpt_AccountProfile_01.
9980            strRpt = "rptAccountProfile_SepChkAcct"
9990          Case False
                ' ** Account, with add'l fields, by specified FormRef('AccountNo').
                ' ** qryRpt_AccountProfile_02.
10000           strRpt = "rptAccountProfile_SngChkAcct"
10010         End Select

              ' ** For the FormRef('AccountNo') in the report queries.
10020         .accountno = rst![accountno]
10030         gstrAccountNo = rst![accountno]
10040         DoEvents

10050         If blnIsWord = True Then
10060           strRptPath = .UserReportPath
10070           strRptCap = "AccountProfile_" & gstrAccountNo & "_" & Format$(Date, "yyyymmdd")
10080           strRptPathFile = strRptPath & LNK_SEP & strRptCap & ".rtf"
10090           lngFiles = lngFiles + 1&
10100           lngE = lngFiles - 1&
10110           ReDim Preserve arr_varFile(F_ELEMS, lngE)
10120           arr_varFile(F_FNAM, lngE) = strRptPathFile
10130           arr_varFile(F_TYP1, lngE) = "Word"
10140           arr_varFile(F_TYP2, lngE) = "Profile"
10150           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, False
10160         ElseIf blnIsExcel = True Then
10170           strRptPath = .UserReportPath
10180           strRptCap = "AccountProfile_" & gstrAccountNo & "_" & Format$(Date, "yyyymmdd")
10190           strRptPathFile = strRptPath & LNK_SEP & strRptCap & ".xls"
10200           lngFiles = lngFiles + 1&
10210           lngE = lngFiles - 1&
10220           ReDim Preserve arr_varFile(F_ELEMS, lngE)
10230           arr_varFile(F_FNAM, lngE) = strRptPathFile
10240           arr_varFile(F_TYP1, lngE) = "Excel"
10250           arr_varFile(F_TYP2, lngE) = "Profile"
10260           Select Case gblnSeparateCheckingAccounts
                Case True
                  ' ** qryRpt_AccountProfile_11_02_08 (xx), For Export.
10270             strQry = "qryRpt_AccountProfile_11_01_09"
10280           Case False
                  ' ** qryRpt_AccountProfile_11_02_08 (xx), For Export.
10290             strQry = "qryRpt_AccountProfile_11_02_09"
10300           End Select
10310           DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
10320           If Excel_NameOnly(strRptPathFile, "Account Profile") = True Then  ' ** Module Function: modExcelFuncs.
10330             If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
10340               EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
10350             End If
10360           End If
10370         Else
                '##GTR_Ref: rptAccountProfile_SepChkAcct
                '##GTR_Ref: rptAccountProfile_SngChkAcct
10380           DoCmd.OpenReport strRpt, acViewNormal
10390         End If

10400         If lngX < lngRecs Then rst.MoveNext
10410       Next

10420       rst.Close
10430       dbs.Close

10440     Case False
            ' ** Validation.

10450       Select Case .opgAccountReview
            Case .opgAccountReview_optSpecified.OptionValue
10460         If IsNull(.cmbMonth) = True Then
10470           blnRetVal = False
10480           MsgBox "You must select a valid month.", vbInformation + vbOKOnly, "Entry Required"
10490           .cmbMonth.SetFocus
10500         Else
                ' ** qryRpt_AccountReview_01_jan:
                ' ** tblMonth, with Account table, for revjan = True; Cartesian
10510           strQry = "qryRpt_AccountReview_" & Right("00" & CStr(.cmbMonth.Column(CBX_MON_ID)), 2) & "_" & .cmbMonth.Column(CBX_MON_SHORT)
10520         End If
10530       Case .opgAccountReview_optAll.OptionValue
              ' ** Union of qryRpt_AccountReview_01_jan - qryRpt_AccountReview_12_dec (tblMonth, with Account table, for revjan = True; Cartesian, etc.).
10540         strQry = "qryRpt_AccountReview_13"
10550       End Select
10560       If blnRetVal = True Then
10570         Set dbs = CurrentDb
10580         With dbs
10590           Set qdf = dbs.QueryDefs(strQry)
10600           Set rst = qdf.OpenRecordset
10610           With rst
10620             If .BOF = True And .EOF = True Then
10630               blnRetVal = False
10640               MsgBox "There are no accounts scheduled for the month of " & Me.cmbMonth.Column(CBX_MON_NAME) & ".", _
                      vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
10650             End If
10660             .Close
10670           End With
10680           .Close
10690         End With
10700       End If

10710     End Select

10720   End With

EXITP:
10730   Set rst = Nothing
10740   Set dbs = Nothing
10750   DoReport = blnRetVal
10760   Exit Function

ERRH:
10770   blnRetVal = False
10780   THAT_PROC = THIS_PROC
10790   That_Erl = Erl
10800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10810   Resume EXITP

End Function
