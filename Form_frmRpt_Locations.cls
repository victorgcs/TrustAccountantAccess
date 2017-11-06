VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Locations"

'VGC 07/16/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, sets options, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Location:         A {cmbLocations}
' **   Close:            C {cmdClose}
' **   Detail:           D {chkDetail}
' **   Excel:            E {cmdExcel}
' **   All:              L {opgLocation_optAll}
' **   Specified:        N {opgLocation_optSpecified}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Focus Holder:     F {FocusHolder}

Private blnGoingToReport As Boolean, blnWasGTR As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnWasGTR = gblnGoToReport

160           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
170           .cmdExcel.Enabled = False
      #Else
180           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberLocation!
              ' ** NONE!

190           opgLocation_AfterUpdate  ' ** Procedure: Below.
200           chkDetail_AfterUpdate  ' ** Procedure: Below.
210           chkRememberLocation_AfterUpdate  ' ** Procedure: Below.
220           chkPageOf_AfterUpdate  ' ** Procedure: Below.

230         Else
240           Cancel = -1
250         End If
260       Else
270         Cancel = -1
280       End If
290     End With

300     DoCmd.Hourglass False

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     DoCmd.Hourglass False
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

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

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Focus Holder:     F {FocusHolder}

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
1000    Select Case ERR.Number
        Case Else
1010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1020    End Select
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
1250          If garr_varGoToReport(GTR_RNAM) = "rptLocations_Detailed" Then
1260            .GoToReport_arw_det_img.Visible = False
1270            If .cmdPrint.Enabled = True Then
1280              .cmdPrint.SetFocus
1290            Else
1300              .cmdClose.SetFocus
1310            End If
1320            DoEvents
1330            blnGoingToReport = False
1340            gblnGoToReport = False
1350            Beep
1360            DoCmd.Hourglass False
1370            DoEvents
1380          End If
1390        Case False
1400          DoCmd.Hourglass True  ' ** Make sure it's still running.
1410          DoEvents
1420          blnGoingToReport = True
1430          If garr_varGoToReport(GTR_RNAM) = "rptLocations_Detailed" Then
1440            .GoToReport_arw_det_img.Visible = True
1450            DoEvents
1460            If .chkDetail = False Then
1470              .chkDetail = True
1480              chkDetail_AfterUpdate  ' ** Procedure: Below
1490            End If
1500            DoEvents
1510            .TimerInterval = GTR_WAIT
1520          Else
1530            If .chkDetail = True Then
1540              .chkDetail = False
1550              chkDetail_AfterUpdate  ' ** Procedure: Below
1560              DoEvents
1570            End If
1580            If .cmdPrint.Enabled = True Then
1590              .cmdPrint.SetFocus
1600            Else
1610              .cmdClose.SetFocus
1620            End If
1630            DoEvents
1640            blnGoingToReport = False
1650            gblnGoToReport = False
1660            Beep
1670            DoCmd.Hourglass False
1680            DoEvents
1690          End If
1700        End Select
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

Private Sub Form_Unload(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1810    With Me

1820      Select Case blnWasGTR
          Case True
1830        blnWasGTR = False
1840      Case False
            ' ** Don't save settings imposed by GoToReport!
1850        If .chkRememberLocation = False Then
1860          .cmbLocations = Null
1870        End If
1880      End Select

1890      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1900      gstrFormQuerySpec = vbNullString
1910      gblnSetFocus = True
1920      gstrReturningForm = THIS_NAME
1930      Select Case strCallingForm
          Case "frmMenu_Asset"
1940        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
1950      End Select

1960    End With

EXITP:
1970    Exit Sub

ERRH:
1980    Select Case ERR.Number
        Case Else
1990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000    End Select
2010    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2110    DoCmd.Close acForm, THIS_NAME

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyTab
2270        With Me
2280          intRetVal = 0
2290          .opgLocation.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab
2350        With Me
2360          intRetVal = 0
2370          If .cmdExcel.Enabled = True Then
2380            .cmdExcel.SetFocus
2390          Else
2400            .cmdWord.SetFocus
2410          End If
2420        End With
2430      End Select
2440    End If

EXITP:
2450    KeyCode = intRetVal
2460    Exit Sub

ERRH:
2470    intRetVal = 0
2480    Select Case ERR.Number
        Case Else
2490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2500    End Select
2510    Resume EXITP

End Sub

Private Sub opgLocation_AfterUpdate()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocation_AfterUpdate"

2610    With Me
2620      Select Case .opgLocation
          Case .opgLocation_optSpecified.OptionValue
2630        .opgLocation_optSpecified_lbl.FontBold = True
2640        .opgLocation_optAll_lbl.FontBold = False
2650        .opgLocation_optSpecified_lbl_box.Visible = True
2660        .opgLocation_optAll_lbl_box.Visible = False
2670        .cmbLocations.Enabled = True
2680        .cmbLocations.BorderColor = CLR_LTBLU2
2690        .cmbLocations.BackStyle = acBackStyleNormal
2700        .chkRememberLocation.Enabled = True
2710        .chkRememberLocation_lbl.Visible = True
2720        .chkRememberLocation_lbl2_dim.Visible = False
2730        .chkRememberLocation_lbl2_dim_hi.Visible = False
2740      Case .opgLocation_optAll.OptionValue
2750        .opgLocation_optSpecified_lbl.FontBold = False
2760        .opgLocation_optAll_lbl.FontBold = True
2770        .opgLocation_optSpecified_lbl_box.Visible = False
2780        .opgLocation_optAll_lbl_box.Visible = True
2790        .cmbLocations.Enabled = False
2800        .cmbLocations.BorderColor = WIN_CLR_DISR
2810        .cmbLocations.BackStyle = acBackStyleTransparent
2820        .chkRememberLocation.Enabled = False
2830        .chkRememberLocation_lbl.Visible = False
2840        .chkRememberLocation_lbl2_dim.Visible = True
2850        .chkRememberLocation_lbl2_dim_hi.Visible = True
2860      End Select
2870    End With

EXITP:
2880    Exit Sub

ERRH:
2890    Select Case ERR.Number
        Case Else
2900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2910    End Select
2920    Resume EXITP

End Sub

Private Sub opgLocation_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocation_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3020    intShiftDown = (Shift And acShiftMask) > 0
3030    intAltDown = (Shift And acAltMask) > 0
3040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3070        With Me
3080          intRetVal = 0
3090          If .cmbLocations.Enabled = True Then
3100            .cmbLocations.SetFocus
3110          Else
3120            .cmdPreview.SetFocus
3130          End If
3140        End With
3150      End Select
3160    End If

        ' ** Shift keys.
3170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3180      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3190        With Me
3200          intRetVal = 0
3210          .cmdClose.SetFocus
3220        End With
3230      End Select
3240    End If

        ' ** Ctrl keys.
3250    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3270        With Me
3280          intRetVal = 0
3290          .chkDetail.SetFocus
3300        End With
3310      End Select
3320    End If

EXITP:
3330    KeyCode = intRetVal
3340    Exit Sub

ERRH:
3350    intRetVal = 0
3360    Select Case ERR.Number
        Case Else
3370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3380    End Select
3390    Resume EXITP

End Sub

Private Sub opgLocation_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocation_optAll_KeyDown"

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
3490          If .cmbLocations.Enabled = True Then
3500            .cmbLocations.SetFocus
3510          Else
3520            .cmdPreview.SetFocus
3530          End If
3540        End With
3550      End Select
3560    End If

        ' ** Shift keys.
3570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3580      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3590        With Me
3600          intRetVal = 0
3610          .cmdClose.SetFocus
3620        End With
3630      End Select
3640    End If

        ' ** Ctrl keys.
3650    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3670        With Me
3680          intRetVal = 0
3690          .chkDetail.SetFocus
3700        End With
3710      End Select
3720    End If

EXITP:
3730    KeyCode = intRetVal
3740    Exit Sub

ERRH:
3750    intRetVal = 0
3760    Select Case ERR.Number
        Case Else
3770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3780    End Select
3790    Resume EXITP

End Sub

Public Sub chkDetail_AfterUpdate()
' ** Called by:
' **   frmMenu_Asset:
' **     cmdLocations_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkDetail_AfterUpdate"

3810    With Me
3820      Select Case .chkDetail
          Case True
3830        .chkDetail_lbl.FontBold = True
3840      Case False
3850        .chkDetail_lbl.FontBold = False
3860      End Select
3870    End With

EXITP:
3880    Exit Sub

ERRH:
3890    Select Case ERR.Number
        Case Else
3900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3910    End Select
3920    Resume EXITP

End Sub

Private Sub chkDetail_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkDetail_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4020    intShiftDown = (Shift And acShiftMask) > 0
4030    intAltDown = (Shift And acAltMask) > 0
4040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4070        With Me
4080          intRetVal = 0
4090          .chkRememberLocation.SetFocus
4100        End With
4110      End Select
4120    End If

        ' ** Shift keys.
4130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4150        With Me
4160          intRetVal = 0
4170          .opgLocation.SetFocus
4180        End With
4190      End Select
4200    End If

EXITP:
4210    KeyCode = intRetVal
4220    Exit Sub

ERRH:
4230    intRetVal = 0
4240    Select Case ERR.Number
        Case Else
4250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4260    End Select
4270    Resume EXITP

End Sub

Private Sub chkRememberLocation_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberLocation_AfterUpdate"

4310    With Me
4320      Select Case .chkRememberLocation
          Case True
4330        .chkRememberLocation_lbl.FontBold = True
4340        .chkRememberLocation_lbl2_dim.FontBold = True
4350        .chkRememberLocation_lbl2_dim_hi.FontBold = True
4360      Case False
4370        .chkRememberLocation_lbl.FontBold = False
4380        .chkRememberLocation_lbl2_dim.FontBold = False
4390        .chkRememberLocation_lbl2_dim_hi.FontBold = False
4400      End Select
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

Private Sub chkRememberLocation_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberLocation_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
4570        With Me
4580          intRetVal = 0
4590          If .cmbLocations.Enabled = True Then
4600            .cmbLocations.SetFocus
4610          Else
4620            .cmdPreview.SetFocus
4630          End If
4640        End With
4650      End Select
4660    End If

        ' ** Shift keys.
4670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4680      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4690        With Me
4700          intRetVal = 0
4710          .chkDetail.SetFocus
4720        End With
4730      End Select
4740    End If

EXITP:
4750    KeyCode = intRetVal
4760    Exit Sub

ERRH:
4770    intRetVal = 0
4780    Select Case ERR.Number
        Case Else
4790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4800    End Select
4810    Resume EXITP

End Sub

Private Sub cmbLocations_Enter()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_Enter"

4910    With Me
4920      If .cmbLocations.text = vbNullString Then
4930        ExpandCombo .cmbLocations  ' ** Module Procedure: modUtilities.
4940        .cmbLocations = .cmbLocations.ItemData(0)
4950      End If
4960    End With

EXITP:
4970    Exit Sub

ERRH:
4980    Select Case ERR.Number
        Case Else
4990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5000    End Select
5010    Resume EXITP

End Sub

Private Sub cmbLocations_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbLocations_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5170        With Me
5180          intRetVal = 0
5190          .cmdPreview.SetFocus
5200        End With
5210      End Select
5220    End If

        ' ** Shift keys.
5230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5250        With Me
5260          intRetVal = 0
5270          .opgLocation.SetFocus
5280        End With
5290      End Select
5300    End If

        ' ** Ctrl keys.
5310    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5330        With Me
5340          intRetVal = 0
5350          .chkPageOf.SetFocus
5360        End With
5370      End Select
5380    End If

        ' ** Ctrl-Shift keys.
5390    If intCtrlDown And (Not intAltDown) And intShiftDown Then
5400      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5410        With Me
5420          intRetVal = 0
5430          .chkRememberLocation.SetFocus
5440        End With
5450      End Select
5460    End If

EXITP:
5470    KeyCode = intRetVal
5480    Exit Sub

ERRH:
5490    intRetVal = 0
5500    Select Case ERR.Number
        Case Else
5510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5520    End Select
5530    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

5610    With Me
5620      Select Case .chkPageOf
          Case True
5630        .chkPageOf_lbl.FontBold = True
5640      Case False
5650        .chkPageOf_lbl.FontBold = False
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

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
5870        With Me
5880          intRetVal = 0
5890          .cmdPreview.SetFocus
5900        End With
5910      End Select
5920    End If

        ' ** Shift keys.
5930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5950        With Me
5960          intRetVal = 0
5970          .cmbLocations.SetFocus
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

Private Sub cmdPreview_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

6110    With Me
6120      If DoReport = True Then  ' ** Function: Below.
6130        Select Case .chkDetail
            Case True
6140          strDocName = "rptLocations_Detailed"
6150        Case False
6160          strDocName = "rptLocations"
6170        End Select
6180        Select Case .opgLocation
            Case .opgLocation_optSpecified.OptionValue
6190          DoCmd.OpenReport strDocName, acViewPreview, , "[Location_ID] = " & CStr(.cmbLocations)
6200        Case .opgLocation_optAll.OptionValue
6210          DoCmd.OpenReport strDocName, acViewPreview
6220        End Select
6230        DoCmd.Maximize
6240        DoCmd.RunCommand acCmdFitToWindow
6250      End If
6260    End With

EXITP:
6270    Exit Sub

ERRH:
6280    Select Case ERR.Number
        Case Else
6290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6300    End Select
6310    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
6490          .cmdPrint.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab
6550        With Me
6560          intRetVal = 0
6570          If .cmbLocations.Enabled = True Then
6580            .cmbLocations.SetFocus
6590          Else
6600            .opgLocation.SetFocus
6610          End If
6620        End With
6630      End Select
6640    End If

        ' ** Ctrl-Shift keys.
6650    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6660      Select Case intRetVal
          Case vbKeyTab
6670        With Me
6680          intRetVal = 0
6690          .chkPageOf.SetFocus
6700        End With
6710      End Select
6720    End If

EXITP:
6730    KeyCode = intRetVal
6740    Exit Sub

ERRH:
6750    intRetVal = 0
6760    Select Case ERR.Number
        Case Else
6770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6780    End Select
6790    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String, intMode As Integer

6810    With Me
6820      If DoReport = True Then  ' ** Function: Below.
6830        Select Case .chkDetail
            Case True
6840          strDocName = "rptLocations_Detailed"
6850        Case False
6860          strDocName = "rptLocations"
6870        End Select
6880        If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
6890          intMode = acViewPreview
6900        Else
6910          intMode = acViewNormal
6920        End If
6930        Select Case .opgLocation
            Case .opgLocation_optSpecified.OptionValue
              '##GTR_Ref: rptLocations
              '##GTR_Ref: rptLocations_Detailed
6940          DoCmd.OpenReport strDocName, acViewNormal, , "[Location_ID] = " & CStr(.cmbLocations)
6950        Case .opgLocation_optAll.OptionValue
6960          DoCmd.OpenReport strDocName, acViewNormal
6970        End Select
6980        If intMode = acViewPreview Then
6990          DoCmd.Maximize
7000          DoCmd.RunCommand acCmdFitToWindow
7010        End If
7020      End If
7030    End With

EXITP:
7040    Exit Sub

ERRH:
7050    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do Nothing.
7060    Case Else
7070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7080    End Select
7090    Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7120    intShiftDown = (Shift And acShiftMask) > 0
7130    intAltDown = (Shift And acAltMask) > 0
7140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7160      Select Case intRetVal
          Case vbKeyTab
7170        With Me
7180          intRetVal = 0
7190          .cmdWord.SetFocus
7200        End With
7210      End Select
7220    End If

        ' ** Shift keys.
7230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7240      Select Case intRetVal
          Case vbKeyTab
7250        With Me
7260          intRetVal = 0
7270          .cmdPreview.SetFocus
7280        End With
7290      End Select
7300    End If

EXITP:
7310    KeyCode = intRetVal
7320    Exit Sub

ERRH:
7330    intRetVal = 0
7340    Select Case ERR.Number
        Case Else
7350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7360    End Select
7370    Resume EXITP

End Sub

Private Sub cmdWord_Click()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim strLocName As String

7410    With Me
7420      If DoReport = True Then  ' ** Procedure: Below.

7430        If IsNull(.UserReportPath) = True Then
7440          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
7450        Else
7460          strRptPath = .UserReportPath
7470        End If

7480        Select Case .chkDetail
            Case True
7490          strRpt = "rptLocations_Detailed"
7500          strRptCap = "Locations_Detailed_"
7510        Case False
7520          strRpt = "rptLocations"
7530          strRptCap = "Locations_"
7540        End Select

7550        Select Case .opgLocation
            Case .opgLocation_optSpecified.OptionValue
7560          strLocName = .cmbLocations.Column(1)
7570          strLocName = Rem_Period(strLocName)  ' ** Module Function: modStringFuncs.
7580          strLocName = SpaceToUnder(strLocName)  ' ** Module Function: modStringFuncs.
7590          strLocName = Left(strLocName, 10)
7600          strRpt = strRpt & strLocName & "_"
7610          gstrPurchaseAsset = CStr(.cmbLocations.Column(0))
7620        Case .opgLocation_optAll.OptionValue
7630          strRptCap = strRptCap & "All_"
7640          gstrPurchaseAsset = vbNullString
7650        End Select

7660        strRpt = strRpt & Format$(Date, "yymmdd")

7670        strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

7680        If strRptPathFile <> vbNullString Then
7690          gblnMessage = True  ' ** Indicating is this going to Word.
7700          DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
7710          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
7720        End If

7730      End If
7740      gstrPurchaseAsset = vbNullString
7750      gblnMessage = False
7760    End With

EXITP:
7770    Exit Sub

ERRH:
7780    Select Case ERR.Number
        Case Else
7790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7800    End Select
7810    Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7920    intShiftDown = (Shift And acShiftMask) > 0
7930    intAltDown = (Shift And acAltMask) > 0
7940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7960      Select Case intRetVal
          Case vbKeyTab
7970        With Me
7980          intRetVal = 0
7990          If .cmdExcel.Enabled = True Then
8000            .cmdExcel.SetFocus
8010          Else
8020            .cmdClose.SetFocus
8030          End If
8040        End With
8050      End Select
8060    End If

        ' ** Shift keys.
8070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8080      Select Case intRetVal
          Case vbKeyTab
8090        With Me
8100          intRetVal = 0
8110          .cmdPrint.SetFocus
8120        End With
8130      End Select
8140    End If

EXITP:
8150    KeyCode = intRetVal
8160    Exit Sub

ERRH:
8170    intRetVal = 0
8180    Select Case ERR.Number
        Case Else
8190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8200    End Select
8210    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim strLocName As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

8310    With Me

8320      DoCmd.Hourglass True
8330      DoEvents

8340      blnContinue = True

8350      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
8360        DoCmd.Hourglass False
8370        msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
8380        If msgResponse <> vbRetry Then
8390          blnContinue = False
8400        End If
8410      End If

8420      If blnContinue = True Then

8430        DoCmd.Hourglass True
8440        DoEvents

8450        If DoReport = True Then  ' ** Procedure: Below.

8460          blnContinue = True

8470          If IsNull(.UserReportPath) = True Then
8480            strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
8490          Else
8500            strRptPath = .UserReportPath
8510          End If

8520          strRptCap = "Locations_"
8530          Select Case .opgLocation
              Case .opgLocation_optSpecified.OptionValue
8540            strLocName = .cmbLocations.Column(1)
8550            strLocName = Rem_Period(strLocName)  ' ** Module Function: modStringFuncs.
8560            strLocName = SpaceToUnder(strLocName)  ' ** Module Function: modStringFuncs.
8570            strLocName = Left(strLocName, 10)
8580            strRptCap = strRptCap & strLocName & "_"
8590          Case .opgLocation_optAll.OptionValue
8600            strRptCap = strRptCap & "All_"
8610          End Select
8620          If .chkDetail = True Then
8630            strRptCap = strRptCap & "Detailed_"
8640          End If
8650          strRptCap = strRptCap & Format$(Date, "yymmdd")

8660          Select Case .opgLocation
              Case .opgLocation_optSpecified.OptionValue
8670            Select Case .chkDetail
                Case True
                  ' ** qryRpt_Locations_Detailed_23 (xx), For Export, Specified Location, Detailed.
8680              strQry = "qryRpt_Locations_Detailed_24"
8690            Case False
                  ' ** qryRpt_Locations_21 (xx), For Export, Specified Location, Simple.
8700              strQry = "qryRpt_Locations_22"
8710            End Select
8720          Case .opgLocation_optAll.OptionValue
8730            Select Case .chkDetail
                Case True
                  ' ** qryRpt_Locations_Detailed_40 (xx), For Export, All Locations, Detailed.
8740              strQry = "qryRpt_Locations_Detailed_41"
8750            Case False
                  ' ** qryRpt_Locations_34 (xx), For Export, All Locations, Simple.
8760              strQry = "qryRpt_Locations_35"
8770            End Select
8780          End Select

8790          strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

8800          If strRptPathFile <> vbNullString Then
8810            If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
8820              EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
8830            End If
8840            DoEvents
8850            If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
8860              Kill strRptPathFile
8870            End If
8880            DoEvents
8890            gstrFormQuerySpec = THIS_NAME
8900            DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
8910            If Excel_NameOnly(strRptPathFile, "Locations") = True Then  ' ** Module Function: modExcelFuncs.
8920              If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
8930                EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
8940              End If
8950              OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
8960            End If
8970            .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
8980          End If

8990        End If  ' ** DoReport().

9000      End If  ' ** blnContinue.

9010      DoCmd.Hourglass False

9020    End With

EXITP:
9030    Exit Sub

ERRH:
9040    DoCmd.Hourglass False
9050    Select Case ERR.Number
        Case Else
9060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9070    End Select
9080    Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9120    intShiftDown = (Shift And acShiftMask) > 0
9130    intAltDown = (Shift And acAltMask) > 0
9140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9160      Select Case intRetVal
          Case vbKeyTab
9170        With Me
9180          intRetVal = 0
9190          .cmdClose.SetFocus
9200        End With
9210      End Select
9220    End If

        ' ** Shift keys.
9230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9240      Select Case intRetVal
          Case vbKeyTab
9250        With Me
9260          intRetVal = 0
9270          .cmdWord.SetFocus
9280        End With
9290      End Select
9300    End If

EXITP:
9310    KeyCode = intRetVal
9320    Exit Sub

ERRH:
9330    intRetVal = 0
9340    Select Case ERR.Number
        Case Else
9350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9360    End Select
9370    Resume EXITP

End Sub

Private Function DoReport() As Boolean

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

9410    blnRetVal = True

9420    With Me

9430      Select Case .opgLocation
          Case .opgLocation_optSpecified.OptionValue
9440        If .cmbLocations.Visible = True Then
9450          If IsNull(.cmbLocations) = True Then
9460            blnRetVal = False
9470            MsgBox "You must select a location to continue.", vbInformation + vbOKOnly, "Entry Required"
9480          Else
9490            If .cmbLocations = vbNullString Then
9500              blnRetVal = False
9510              MsgBox "You must select a location to continue.", vbInformation + vbOKOnly, "Entry Required"
9520            End If
9530          End If
9540        End If
9550      Case .opgLocation_optAll.OptionValue
            ' ** Nothing required.
9560      End Select

9570      If blnRetVal = True Then

9580        gstrFormQuerySpec = THIS_NAME

9590        Set dbs = CurrentDb

9600        Select Case .opgLocation
            Case .opgLocation_optSpecified.OptionValue
9610          Select Case .chkDetail
              Case True
                ' ** ActiveAssets, linked to Location, by specified FormRef('Location_ID').
9620            Set qdf = dbs.QueryDefs("qryRpt_Locations_04")
9630          Case False
                ' ** ActiveAssets, linked to Location, grouped by Location_ID, assetno, accountno, by specified FormRef('Location_ID').
9640            Set qdf = dbs.QueryDefs("qryRpt_Locations_02")
9650          End Select
9660        Case .opgLocation_optAll.OptionValue
9670          Select Case .chkDetail
              Case True
                ' ** ActiveAssets, linked to Location.
9680            Set qdf = dbs.QueryDefs("qryRpt_Locations_03")
9690          Case False
                ' ** ActiveAssets, linked to Location, grouped by Location_ID, assetno, accountno.
9700            Set qdf = dbs.QueryDefs("qryRpt_Locations_01")
9710          End Select
9720        End Select

9730        Set rst = qdf.OpenRecordset
9740        With rst
9750          If .BOF = True And .EOF = True Then
9760            blnRetVal = False
9770            MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
9780          End If
9790          .Close
9800        End With

9810        dbs.Close

9820      End If  ' ** blnRetVal.

9830    End With  ' ** Me.

EXITP:
9840    Set qdf = Nothing
9850    Set rst = Nothing
9860    Set dbs = Nothing
9870    DoReport = blnRetVal
9880    Exit Function

ERRH:
9890    blnRetVal = False
9900    Select Case ERR.Number
        Case Else
9910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9920    End Select
9930    Resume EXITP

End Function
