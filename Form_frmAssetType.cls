VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetType"

'VGC 09/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Show Assets:      H {chkShowAssets}
' **   Print Report:     P {cmdPrintReport}
' **   Exit:             X {cmdClose}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, lngPrintReportLbl_Left As Long, blnWasGTR As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim frm As Access.Form
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnWasGTR = gblnGoToReport

160           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

170           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

180           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200           Set frm = .frmAssetType_Sub.Form
210           Select Case gblnIncomeTaxCoding
              Case True
220             With frm
230               .taxcode_description.Enabled = True
240               .taxcode_description.Locked = True
250               .taxcode_description.ForeColor = CLR_DKGRY
260               .taxcode_description.BackStyle = acBackStyleNormal
270               .taxcode_description.BorderColor = CLR_LTBLU2
280               .TaxCode_Description_lbl.ForeColor = CLR_DKGRY2
290               .taxcode_description_lbl_dim_hi.Visible = False
300               .taxcode_description_lbl2.ForeColor = CLR_TEAL
310               .taxcode_description_lbl2_dim_hi.Visible = False
320               .TaxCode_Description_lbl_line.BorderColor = CLR_DKGRY
330               .taxcode_description_lbl_line_dim_hi.Visible = False
340               .taxcode_description_lbl3.Visible = False
350             End With
360             .taxcode_description_lbl3.ForeColor = CLR_TEAL
370             .taxcode_description_lbl3_dim_hi.Visible = False
380             .taxcode_description_lbl4.ForeColor = CLR_TEAL
390             .taxcode_description_lbl4_dim_hi.Visible = False
400           Case False
410             With frm
420               .taxcode_description.Enabled = False
430               .taxcode_description.Locked = True
440               .taxcode_description.ForeColor = CLR_LTGRY2
450               .taxcode_description.BackStyle = acBackStyleTransparent
460               .taxcode_description.BorderColor = WIN_CLR_DISR
470               .TaxCode_Description_lbl.ForeColor = WIN_CLR_DISF
480               .taxcode_description_lbl_dim_hi.Visible = True
490               .taxcode_description_lbl2.ForeColor = WIN_CLR_DISF
500               .taxcode_description_lbl2_dim_hi.Visible = True
510               .TaxCode_Description_lbl_line.BorderColor = WIN_CLR_DISR
520               .taxcode_description_lbl_line_dim_hi.Visible = True
530               .taxcode_description_lbl3.Visible = True
540             End With
550             .taxcode_description_lbl3.ForeColor = WIN_CLR_DISF
560             .taxcode_description_lbl3_dim_hi.Visible = True
570             .taxcode_description_lbl4.ForeColor = WIN_CLR_DISF
580             .taxcode_description_lbl4_dim_hi.Visible = True
590           End Select

600           chkShowAssets_AfterUpdate  ' ** Procedire: Below.

610         Else
620           Cancel = -1
630         End If
640       Else
650         Cancel = -1
660       End If
670     End With

680     If Cancel = -1 Then
690       Beep
700       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
710     End If

EXITP:
720     Set frm = Nothing
730     Exit Sub

ERRH:
740     Select Case ERR.Number
        Case Else
750       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
760     End Select
770     Resume EXITP

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

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       End Select
900     End If

        ' ** Alt keys.
910     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
920       Select Case intRetVal
          Case vbKeyX
930         intRetVal = 0
940         cmdClose_Click  ' ** Procedure: Below.
950       End Select
960     End If

        ' ** Ctrl-Shift keys.
970     If intCtrlDown And (Not intAltDown) And intShiftDown Then
980       Select Case intRetVal
          Case vbKeyF
990         With Me
1000          intRetVal = 0
1010          .FocusHolder.SetFocus
1020        End With
1030      End Select
1040    End If

EXITP:
1050    KeyCode = intRetVal
1060    Exit Sub

ERRH:
1070    intRetVal = 0
1080    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1090    Case Else
1100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1110    End Select
1120    Resume EXITP

End Sub

Private Sub Form_Timer()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1210    With Me
1220      .TimerInterval = 0&
1230      If gblnSetFocus = True Then
1240        gblnSetFocus = False
1250        strFormName = vbNullString
1260  On Error Resume Next
1270        Do Until strFormName = THIS_NAME
1280          DoCmd.SelectObject acForm, THIS_NAME, False
1290          strFormName = Screen.ActiveForm.Name
1300        Loop
1310  On Error GoTo ERRH
1320        .frmAssetType_Sub.SetFocus
1330      End If
1340      If gblnGoToReport = True Then
1350        Select Case blnGoingToReport
            Case True
1360          Select Case blnGoingToReport2
              Case True
1370            DoCmd.Hourglass False  ' ** It should already be off.
1380            gblnGoToReport = False
1390            blnGoingToReport = False
1400            blnGoingToReport2 = False
1410            .GoToReport_arw_pr_img.Visible = False
1420            .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1430            DoEvents
1440          Case False
1450            If garr_varGoToReport(GTR_RNAM) = "rptListOfAssetTypes_Detailed" Then
1460              blnGoingToReport2 = True
1470              .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1480              .GoToReport_arw_pr_img.Visible = True
1490              .GoToReport_arw_det_img.Visible = False
1500              .cmdPrintReport.SetFocus
1510              Beep
1520              DoCmd.Hourglass False
1530              DoEvents
1540              .TimerInterval = (GTR_WAIT * 2&)
1550            Else
1560              DoCmd.Hourglass False  ' ** It should already be off.
1570              gblnGoToReport = False
1580              blnGoingToReport = False
1590              .GoToReport_arw_pr_img.Visible = False
1600              .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1610              DoEvents
1620            End If
1630          End Select
1640        Case False
1650          DoCmd.Hourglass True  ' ** Make sure it's still running.
1660          DoEvents
1670          blnGoingToReport = True
1680          If garr_varGoToReport(GTR_RNAM) = "rptListOfAssetTypes_Detailed" Then
1690            .GoToReport_arw_det_img.Visible = True
1700            DoEvents
1710            If .chkShowAssets = False Then
1720              .chkShowAssets = True
1730              chkShowAssets_AfterUpdate  ' ** Procedure: Below.
1740            End If
1750            DoEvents
1760            .TimerInterval = GTR_WAIT
1770          Else
1780            .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1790            .GoToReport_arw_pr_img.Visible = True
1800            .cmdPrintReport.SetFocus
1810            DoEvents
1820            If .chkShowAssets = True Then
1830              .chkShowAssets = False
1840              chkShowAssets_AfterUpdate  ' ** Procedure: Below.
1850              DoEvents
1860            End If
1870            Beep
1880            DoCmd.Hourglass False
1890            DoEvents
1900            .TimerInterval = (GTR_WAIT * 2&)
1910          End If
1920        End Select
1930      Else
1940        DoCmd.Hourglass False
1950      End If
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

Private Sub Form_Unload(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean

2110    With Me

2120      DoCmd.Hourglass True
2130      DoEvents

2140      Select Case blnWasGTR
          Case True
2150        blnWasGTR = False
2160      Case False
            ' ** Don't save settings imposed by GoToReport!
2170        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
2180      End Select

2190      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

2200      gstrReturningForm = THIS_NAME
2210      gblnSetFocus = True
2220      strDocName = "frmMenu_SystemTable"
2230      DoCmd.OpenForm strDocName, , , , , , "frmMenu_Maintenance"

2240    End With

EXITP:
2250    Exit Sub

ERRH:
2260    DoCmd.Hourglass False
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
2420      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2430        Select Case blnPrintReport_Focus
            Case True
2440          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2450          .cmdPrintReport_raised_img.Visible = False
2460        Case False
2470          .cmdPrintReport_raised_img.Visible = True
2480          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2490        End Select
2500        .cmdPrintReport_raised_focus_img.Visible = False
2510        .cmdPrintReport_raised_focus_dots_img.Visible = False
2520        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2530        .cmdPrintReport_raised_img_dis.Visible = False
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

Public Sub cmdClose_Click()

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
2990          .frmAssetType_Sub.SetFocus
3000          .frmAssetType_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmassettype_Sub.
3010          .frmAssetType_Sub.Form.FocusHolder.SetFocus
3020        End With
3030      End Select
3040    End If

        ' ** Shift keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3060      Select Case intRetVal
          Case vbKeyTab
3070        With Me
3080          intRetVal = 0
3090          If .cmdPrintReport.Enabled = True Then
3100            .cmdPrintReport.SetFocus
3110          Else
3120            .chkShowAssets.SetFocus
3130          End If
3140        End With
3150      End Select
3160    End If

EXITP:
3170    KeyCode = intRetVal
3180    Exit Sub

ERRH:
3190    intRetVal = 0
3200    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3210    Case Else
3220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3230    End Select
3240    Resume EXITP

End Sub

Private Sub chkShowAssets_AfterUpdate()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssets_AfterUpdate"

3310    With Me
3320      Select Case .chkShowAssets
          Case True
3330        .chkShowAssets_lbl.FontBold = True
3340        If .frmAssetType_Sub.Form.RecordSource <> "qryAssetType_Sub_02" Then
              ' ** qryAssetType_Sub_01 (AssetType, linked to AssetTypeGrouping, TaxCode),
              ' ** linked to MasterAsset, with IsHidden, assetno, asset_description.
3350          .frmAssetType_Sub.Form.RecordSource = "qryAssetType_Sub_02"
3360        End If
3370        .frmAssetType_Sub.Form.ShowAsset True  ' ** Form Procedure: frmAssetType_Sub.
3380      Case False
3390        .chkShowAssets_lbl.FontBold = False
3400        If .frmAssetType_Sub.Form.RecordSource <> "qryAssetType_Sub_01" Then
              ' ** AssetType, linked to AssetTypeGrouping, TaxCode.
3410          .frmAssetType_Sub.Form.RecordSource = "qryAssetType_Sub_01"
3420        End If
3430        .frmAssetType_Sub.Form.ShowAsset False  ' ** Form Procedure: frmAssetType_Sub.
3440      End Select
3450      .frmAssetType_Sub.Form.Requery
3460      .frmAssetType_Sub.SetFocus
3470    End With

EXITP:
3480    Exit Sub

ERRH:
3490    Select Case ERR.Number
        Case Else
3500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3510    End Select
3520    Resume EXITP

End Sub

Private Sub chkShowAssets_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssets_KeyDown"

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
3690          If .cmdPrintReport.Enabled = True Then
3700            .cmdPrintReport.SetFocus
3710          Else
3720            .cmdClose.SetFocus
3730          End If
3740        End With
3750      End Select
3760    End If

        ' ** Shift keys.
3770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3780      Select Case intRetVal
          Case vbKeyTab
3790        With Me
3800          intRetVal = 0
3810          .frmAssetType_Sub.SetFocus
3820          .frmAssetType_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmassettype_Sub.
3830          .frmAssetType_Sub.Form.FocusHolder.SetFocus
3840        End With
3850      End Select
3860    End If

EXITP:
3870    KeyCode = intRetVal
3880    Exit Sub

ERRH:
3890    intRetVal = 0
3900    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3910    Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

4010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

4110    With Me
4120      blnPrintReport_Focus = True
4130      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
4140      .cmdPrintReport_raised_img.Visible = False
4150      .cmdPrintReport_raised_focus_img.Visible = False
4160      .cmdPrintReport_raised_focus_dots_img.Visible = False
4170      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4180      .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

4310    With Me
4320      blnPrintReport_MouseDown = True
4330      .cmdPrintReport_sunken_focus_dots_img.Visible = True
4340      .cmdPrintReport_raised_img.Visible = False
4350      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4360      .cmdPrintReport_raised_focus_img.Visible = False
4370      .cmdPrintReport_raised_focus_dots_img.Visible = False
4380      .cmdPrintReport_raised_img_dis.Visible = False
4390    End With

EXITP:
4400    Exit Sub

ERRH:
4410    Select Case ERR.Number
        Case Else
4420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4430    End Select
4440    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

4510    With Me

4520      Select Case .chkShowAssets
          Case True
4530        strDocName = "rptListOfAssetTypes_Detailed"
4540      Case False
4550        strDocName = "rptListOfAssetTypes"
4560      End Select

4570      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4580        DoCmd.OpenReport strDocName, acViewPreview
4590        DoCmd.Maximize
4600        DoCmd.RunCommand acCmdFitToWindow
4610        .Visible = False
4620      Else
            '##GTR_Ref: rptListOfAssetTypes_Detailed
            '##GTR_Ref: rptListOfAssetTypes
4630        DoCmd.OpenReport strDocName, acViewNormal
4640      End If

4650    End With

EXITP:
4660    Exit Sub

ERRH:
4670    Me.Visible = True
4680    DoCmd.Restore
4690    Select Case ERR.Number
        Case Else
4700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4710    End Select
4720    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
4890          .cmdClose.SetFocus
4900        End With
4910      End Select
4920    End If

        ' ** Shift keys.
4930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4940      Select Case intRetVal
          Case vbKeyTab
4950        With Me
4960          intRetVal = 0
4970          .chkShowAssets.SetFocus
4980        End With
4990      End Select
5000    End If

EXITP:
5010    KeyCode = intRetVal
5020    Exit Sub

ERRH:
5030    intRetVal = 0
5040    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5050    Case Else
5060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5070    End Select
5080    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

5110    With Me
5120      If blnPrintReport_MouseDown = False Then
5130        Select Case blnPrintReport_Focus
            Case True
5140          .cmdPrintReport_raised_focus_dots_img.Visible = True
5150          .cmdPrintReport_raised_focus_img.Visible = False
5160        Case False
5170          .cmdPrintReport_raised_focus_img.Visible = True
5180          .cmdPrintReport_raised_focus_dots_img.Visible = False
5190        End Select
5200        .cmdPrintReport_raised_img.Visible = False
5210        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5220        .cmdPrintReport_sunken_focus_dots_img.Visible = False
5230        .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

5410    With Me
5420      .cmdPrintReport_raised_focus_dots_img.Visible = True
5430      .cmdPrintReport_raised_img.Visible = False
5440      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5450      .cmdPrintReport_raised_focus_img.Visible = False
5460      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5470      .cmdPrintReport_raised_img_dis.Visible = False
5480      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

5610    With Me
5620      .cmdPrintReport_raised_img.Visible = True
5630      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5640      .cmdPrintReport_raised_focus_img.Visible = False
5650      .cmdPrintReport_raised_focus_dots_img.Visible = False
5660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5670      .cmdPrintReport_raised_img_dis.Visible = False
5680      blnPrintReport_Focus = False
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

Private Sub cmdPrintReport_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_lbl2_MouseMove"

5810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5820    Exit Sub

ERRH:
5830    Select Case ERR.Number
        Case Else
5840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5850    End Select
5860    Resume EXITP

End Sub
