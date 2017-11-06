VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStateCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmStateCode"

'VGC 09/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Military:         M {chkIncludeMilitary}
' **   Canada:           N {chkIncludeCanada}
' **   Print Report:     P {cmdPrintReport}
' **   Territories:      T {chkIncludeTerritories}
' **   Exit:             X {cmdClose}

' ** Array: arr_varStateQry().
Private lngStateQrys As Long, arr_varStateQry As Variant
'Private Const Q_QID  As Integer = 0
Private Const Q_QNAM As Integer = 1
Private Const Q_BAS  As Integer = 2
Private Const Q_TER  As Integer = 3
Private Const Q_CAN  As Integer = 4
Private Const Q_MIL  As Integer = 5

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long, lngTpp As Long
Private blnAdmin As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim grp As DAO.Group, usr As DAO.User, frm As Access.Form
        Dim strFilter00 As String, strFilter01 As String, strFilter02 As String, strFilter03 As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

180           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200           Set frm = .frmStateCode_Sub.Form

210           strFilter00 = vbNullString: strFilter01 = vbNullString: strFilter02 = vbNullString: strFilter03 = vbNullString
220           If .chkIncludeTerritories = False Then
230             strFilter01 = "[state_territory] = False"
240             frm.state_territory_lbl.ForeColor = WIN_CLR_DISF
250             frm.state_territory_lbl_dim_hi.Visible = True
260             frm.state_territory_lbl2.ForeColor = WIN_CLR_DISF
270             frm.state_territory_lbl2_dim_hi.Visible = True
280             frm.state_territory_lbl_line.BorderColor = WIN_CLR_DISR
290             frm.state_territory_lbl_line_dim_hi.Visible = True
300             frm.state_territory.Locked = False  ' ** It's already disabled.
310             frm.state_territory_box.BorderColor = WIN_CLR_DISR
320             frm.state_territory_box.BackStyle = acBackStyleTransparent
330           Else
340             .chkIncludeTerritories_lbl.FontBold = True
350           End If
360           If .chkIncludeCanada = False Then
370             strFilter02 = "[state_canada] = False"
380             frm.state_canada_lbl.ForeColor = WIN_CLR_DISF
390             frm.state_canada_lbl_dim_hi.Visible = True
400             frm.state_canada_lbl2.ForeColor = WIN_CLR_DISF
410             frm.state_canada_lbl2_dim_hi.Visible = True
420             frm.state_canada_lbl_line.BorderColor = WIN_CLR_DISR
430             frm.state_canada_lbl_line_dim_hi.Visible = True
440             frm.state_canada.Locked = False  ' ** It's already disabled.
450             frm.state_canada_box.BorderColor = WIN_CLR_DISR
460             frm.state_canada_box.BackStyle = acBackStyleTransparent
470           Else
480             .chkIncludeCanada_lbl.FontBold = True
490           End If
500           If .chkIncludeMilitary = False Then
510             strFilter03 = "[state_military] = False"
520             frm.state_military_lbl.ForeColor = WIN_CLR_DISF
530             frm.state_military_lbl_dim_hi.Visible = True
540             frm.state_military_lbl2.ForeColor = WIN_CLR_DISF
550             frm.state_military_lbl2_dim_hi.Visible = True
560             frm.state_military_lbl_line.BorderColor = WIN_CLR_DISR
570             frm.state_military_lbl_line_dim_hi.Visible = True
580             frm.state_military.Locked = False  ' ** It's already disabled.
590             frm.state_military_box.BorderColor = WIN_CLR_DISR
600             frm.state_military_box.BackStyle = acBackStyleTransparent
610           Else
620             .chkIncludeMilitary_lbl.FontBold = True
630           End If
640           If strFilter01 <> vbNullString Then
650             strFilter00 = strFilter01
660           End If
670           If strFilter02 <> vbNullString Then
680             If strFilter00 = vbNullString Then
690               strFilter00 = strFilter02
700             Else
710               strFilter00 = strFilter00 & " And " & strFilter02
720             End If
730           End If
740           If strFilter03 <> vbNullString Then
750             If strFilter00 = vbNullString Then
760               strFilter00 = strFilter03
770             Else
780               strFilter00 = strFilter00 & " And " & strFilter03
790             End If
800           End If
810           If strFilter00 = vbNullString Then
820             frm.Filter = vbNullString
830             frm.FilterOn = False
840           Else
850             frm.Filter = strFilter00
860             frm.FilterOn = True
870           End If

880           blnAdmin = False
890           For Each grp In DBEngine.Workspaces(0).Groups
900             If grp.Name = "Admins" Then
910               For Each usr In grp.Users
920                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
930                   blnAdmin = True
940                   Exit For
950                 End If
960               Next
970             End If
980           Next

990         Else
1000          Cancel = -1
1010        End If
1020      Else
1030        Cancel = -1
1040      End If
1050    End With

1060    If Cancel = -1 Then
1070      Beep
1080      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1090    End If

EXITP:
1100    Set frm = Nothing
1110    Set usr = Nothing
1120    Set grp = Nothing
1130    Exit Sub

ERRH:
1140    Select Case ERR.Number
        Case Else
1150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1160    End Select
1170    Resume EXITP

End Sub

Private Sub Form_Load()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

1210    With Me

1220      Set dbs = CurrentDb
1230      With dbs
            ' ** qryState_04 (tblQuery, just state combo box queries), with Boolean code types.
1240        Set qdf = .QueryDefs("qryState_05")
1250        Set rst = qdf.OpenRecordset
1260        With rst
1270          If .BOF = True And .EOF = True Then
1280            Beep
1290            MsgBox "Data needed for this form was not found.", vbCritical + vbOKOnly, "Data Not Found"
1300          Else
1310            .MoveLast
1320            lngStateQrys = .RecordCount
1330            .MoveFirst
1340            arr_varStateQry = .GetRows(lngStateQrys)
                ' ****************************************************
                ' ** Array: arr_varStateQry()
                ' **
                ' **   Field  Element  Name               Constant
                ' **   =====  =======  =================  ==========
                ' **     1       0     qry_id             Q_QID
                ' **     2       1     qry_name           Q_QNAM
                ' **     3       2     state_basic        Q_BAS
                ' **     4       3     state_territory    Q_TER
                ' **     5       4     state_canada       Q_CAN
                ' **     6       5     state_military     Q_MIL
                ' **
                ' ****************************************************
1350          End If
1360          .Close
1370        End With
1380        .Close
1390      End With

1400      chkIncludeTerritories_AfterUpdate  ' ** Procedure: Below.
1410      chkIncludeCanada_AfterUpdate  ' ** Procedure: Below.
1420      chkIncludeMilitary_AfterUpdate  ' ** Procedure: Below.

1430      .frmStateCode_Sub.Form.SortNow "Form_Load"  ' ** Form Procedure: frmStateCode_Sub.

1440    End With

EXITP:
1450    Set rst = Nothing
1460    Set qdf = Nothing
1470    Set dbs = Nothing
1480    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1620    intShiftDown = (Shift And acShiftMask) > 0
1630    intAltDown = (Shift And acAltMask) > 0
1640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyEscape
1670        intRetVal = 0
1680        cmdClose_Click  ' ** Procedure: Below.
1690      End Select
1700    End If

        ' ** Alt keys.
1710    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1720      Select Case intRetVal
          Case vbKeyX
1730        intRetVal = 0
1740        cmdClose_Click  ' ** Procedure: Below.
1750      End Select
1760    End If

        ' ** Ctrl-Shift keys.
1770    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1780      Select Case intRetVal
          Case vbKeyF
1790        With Me
1800          intRetVal = 0
1810          .FocusHolder.SetFocus
1820        End With
1830      End Select
1840    End If

EXITP:
1850    KeyCode = intRetVal
1860    Exit Sub

ERRH:
1870    intRetVal = 0
1880    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1890    Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2010    With Me
2020      .TimerInterval = 0&
2030      If gblnSetFocus = True Then
2040        gblnSetFocus = False
2050        strFormName = vbNullString
2060  On Error Resume Next
2070        Do Until strFormName = THIS_NAME
2080          DoCmd.SelectObject acForm, THIS_NAME, False
2090          strFormName = Screen.ActiveForm.Name
2100        Loop
2110  On Error GoTo ERRH
2120        .frmStateCode_Sub.SetFocus
2130      End If
2140      If gblnGoToReport = True Then
2150        Select Case blnGoingToReport
            Case True
2160          DoCmd.Hourglass False  ' ** It should already be off.
2170          gblnGoToReport = False
2180          blnGoingToReport = False
2190          .GoToReport_arw_pr_img.Visible = False
2200          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
2210          DoEvents
2220        Case False
2230          DoCmd.Hourglass True  ' ** Make sure it's still running.
2240          DoEvents
2250          blnGoingToReport = True
2260          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + GoToReport_arw_pr_img.Width)
2270          .GoToReport_arw_pr_img.Visible = True
2280          .cmdPrintReport.SetFocus
2290          Beep
2300          DoCmd.Hourglass False
2310          DoEvents
2320          .TimerInterval = (GTR_WAIT * 2&)
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

        Dim strDocName As String
        Dim blnFound As Boolean, blnSkip As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

2510    With Me

2520      DoCmd.Hourglass True
2530      DoEvents

2540      If blnAdmin = False Then
            ' ** Don't save this setting unless it's an administrator.
2550        .state_query = Null
2560      Else
            ' ** Save the state query in the preference.
2570        blnFound = False
2580        For lngX = 0& To (lngStateQrys - 1&)
2590          If arr_varStateQry(Q_BAS, lngX) = True Then
2600            blnSkip = False
2610            If ((.chkIncludeTerritories = True And arr_varStateQry(Q_TER, lngX) = False) Or _
                    (.chkIncludeTerritories = False And arr_varStateQry(Q_TER, lngX) = True)) Then
2620              blnSkip = True
2630            End If
2640            If blnSkip = False Then
2650              If ((.chkIncludeCanada = True And arr_varStateQry(Q_CAN, lngX) = False) Or _
                      (.chkIncludeCanada = False And arr_varStateQry(Q_CAN, lngX) = True)) Then
2660                blnSkip = True
2670              End If
2680            End If
2690            If blnSkip = False Then
2700              If ((.chkIncludeMilitary = True And arr_varStateQry(Q_MIL, lngX) = False) Or _
                      (.chkIncludeMilitary = False And arr_varStateQry(Q_MIL, lngX) = True)) Then
2710                blnSkip = True
2720              End If
2730            End If
2740            If blnSkip = False Then
2750              blnFound = True
2760              .state_query = arr_varStateQry(Q_QNAM, lngX)
2770              Exit For
2780            End If
2790          End If
2800        Next
2810        If blnFound = False Then
2820          Beep
2830          DoCmd.Hourglass False
2840          MsgBox "Didn't find one!", vbInformation + vbOKOnly, "Unexpected Result"
2850          DoCmd.Hourglass True
2860          DoEvents
2870        End If
2880      End If

2890      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2900      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

2910      gstrReturningForm = THIS_NAME
2920      gblnSetFocus = True
2930      strDocName = "frmMenu_SystemTable"
2940      DoCmd.OpenForm strDocName, , , , , , "frmMenu_Maintenance"

2950    End With

EXITP:
2960    Exit Sub

ERRH:
2970    DoCmd.Hourglass False
2980    Select Case ERR.Number
        Case Else
2990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3000    End Select
3010    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3110    With Me
3120      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3130        Select Case blnPrintReport_Focus
            Case True
3140          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3150          .cmdPrintReport_raised_img.Visible = False
3160        Case False
3170          .cmdPrintReport_raised_img.Visible = True
3180          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3190        End Select
3200        .cmdPrintReport_raised_focus_img.Visible = False
3210        .cmdPrintReport_raised_focus_dots_img.Visible = False
3220        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3230        .cmdPrintReport_raised_img_dis.Visible = False
3240      End If
3250    End With

EXITP:
3260    Exit Sub

ERRH:
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

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

Public Sub cmdClose_Click()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3510    DoCmd.Close acForm, THIS_NAME

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

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
3690          .frmStateCode_Sub.SetFocus
3700          .frmStateCode_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmStateCode_Sub.
3710          .frmStateCode_Sub.Form.FocusHolder.SetFocus
3720        End With
3730      End Select
3740    End If

        ' ** Shift keys.
3750    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3760      Select Case intRetVal
          Case vbKeyTab
3770        With Me
3780          intRetVal = 0
3790          .chkIncludeMilitary.SetFocus
3800        End With
3810      End Select
3820    End If

EXITP:
3830    KeyCode = intRetVal
3840    Exit Sub

ERRH:
3850    intRetVal = 0
3860    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3870    Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

4010    With Me
4020      blnPrintReport_Focus = True
4030      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
4040      .cmdPrintReport_raised_img.Visible = False
4050      .cmdPrintReport_raised_focus_img.Visible = False
4060      .cmdPrintReport_raised_focus_dots_img.Visible = False
4070      .cmdPrintReport_sunken_focus_dots_img.Visible = False
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

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

4210    With Me
4220      blnPrintReport_MouseDown = True
4230      .cmdPrintReport_sunken_focus_dots_img.Visible = True
4240      .cmdPrintReport_raised_img.Visible = False
4250      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4260      .cmdPrintReport_raised_focus_img.Visible = False
4270      .cmdPrintReport_raised_focus_dots_img.Visible = False
4280      .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strFilter As String

4410    With Me
4420      Select Case .frmStateCode_Sub.Form.FilterOn
          Case True
4430        strFilter = .frmStateCode_Sub.Form.Filter
4440      Case False
4450        strFilter = vbNullString
4460      End Select
4470      strDocName = "rptListOfStateCodes"
4480      If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4490        DoCmd.OpenReport strDocName, acViewPreview, , strFilter
4500        DoCmd.Maximize
4510        DoCmd.RunCommand acCmdFitToWindow
4520        .Visible = False
4530      Else
            '##GTR_Ref: rptListOfStateCodes
4540        DoCmd.OpenReport strDocName, acViewNormal, , strFilter
4550      End If
4560    End With

EXITP:
4570    Exit Sub

ERRH:
4580    Me.Visible = True
4590    DoCmd.Restore
4600    Select Case ERR.Number
        Case Else
4610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4620    End Select
4630    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4720    intShiftDown = (Shift And acShiftMask) > 0
4730    intAltDown = (Shift And acAltMask) > 0
4740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4760      Select Case intRetVal
          Case vbKeyTab
4770        With Me
4780          intRetVal = 0
4790          .chkIncludeTerritories.SetFocus
4800        End With
4810      End Select
4820    End If

        ' ** Shift keys.
4830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4840      Select Case intRetVal
          Case vbKeyTab
4850        With Me
4860          intRetVal = 0
4870          .frmStateCode_Sub.SetFocus
4880          .frmStateCode_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStateCode_Sub.
4890          .frmStateCode_Sub.Form.FocusHolder.SetFocus
4900        End With
4910      End Select
4920    End If

EXITP:
4930    KeyCode = intRetVal
4940    Exit Sub

ERRH:
4950    intRetVal = 0
4960    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4970    Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

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

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

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

Private Sub chkIncludeTerritories_AfterUpdate()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeTerritories_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String, blnNoFilter As Boolean
        Dim strStateCode As String
        Dim blnSortHere As Boolean, blnResort As Boolean

5910    blnNoFilter = False

5920    With Me
5930      Set frm = .frmStateCode_Sub.Form
5940      blnSortHere = False: blnResort = False
5950      Select Case .chkIncludeTerritories
          Case True
5960        .chkIncludeTerritories_lbl.FontBold = True
5970        frm.state_territory_lbl.ForeColor = CLR_DKGRY2
5980        frm.state_territory_lbl_dim_hi.Visible = False
5990        frm.state_territory_lbl2.ForeColor = CLR_DKGRY2
6000        frm.state_territory_lbl2_dim_hi.Visible = False
6010        frm.state_territory_lbl_line.BorderColor = CLR_DKGRY
6020        frm.state_territory_lbl_line_dim_hi.Visible = False
6030        frm.state_territory.Locked = True
6040        frm.state_territory_box.BorderColor = CLR_GRY
6050        frm.state_territory_box.BackStyle = acBackStyleNormal
6060        If .chkIncludeCanada = True And .chkIncludeMilitary = True Then
6070          strFilter = vbNullString
6080          blnNoFilter = True
6090        ElseIf .chkIncludeCanada = True And .chkIncludeMilitary = False Then
6100          strFilter = "[state_military] = False"
6110        ElseIf .chkIncludeCanada = False And .chkIncludeMilitary = True Then
6120          strFilter = "[state_canada] = False"
6130        ElseIf .chkIncludeCanada = False And .chkIncludeMilitary = False Then
6140          strFilter = "[state_canada] = False And [state_military] = False"
6150        End If
6160      Case False
6170        .chkIncludeTerritories_lbl.FontBold = False
6180        If frm.Sort_line2a.Left = frm.state_territory_lbl.Left And frm.Sort_line2a.Visible = True Then blnSortHere = True
6190        frm.state_territory_lbl.ForeColor = WIN_CLR_DISF
6200        frm.state_territory_lbl_dim_hi.Visible = True
6210        frm.state_territory_lbl2.ForeColor = WIN_CLR_DISF
6220        frm.state_territory_lbl2_dim_hi.Visible = True
6230        frm.state_territory_lbl_line.BorderColor = WIN_CLR_DISR
6240        frm.state_territory_lbl_line_dim_hi.Visible = True
6250        frm.state_territory.Locked = False  ' ** It's already disabled.
6260        frm.state_territory_box.BorderColor = WIN_CLR_DISR
6270        frm.state_territory_box.BackStyle = acBackStyleTransparent
6280        If blnSortHere = True Then
6290          frm.Sort_line2a.Visible = False
6300          frm.Sort_line2b.Visible = False
6310          blnResort = True
6320        End If
6330        blnSortHere = False
6340        If .chkIncludeCanada = True And .chkIncludeMilitary = True Then
6350          strFilter = "[state_territory] = False"
6360        ElseIf .chkIncludeCanada = True And .chkIncludeMilitary = False Then
6370          strFilter = "[state_territory] = False And [state_military] = False"
6380        ElseIf .chkIncludeCanada = False And .chkIncludeMilitary = True Then
6390          strFilter = "[state_territory] = False And [state_canada] = False"
6400        ElseIf .chkIncludeCanada = False And .chkIncludeMilitary = False Then
6410          strFilter = "[state_territory] = False And [state_canada] = False And [state_military] = False"
6420        End If
6430      End Select
6440      frm.Filter = strFilter
6450      Select Case blnNoFilter
          Case True
6460        frm.FilterOn = False
6470      Case False
6480        frm.FilterOn = True
6490      End Select
6500      frm.Requery
6510      If blnResort = True Then
6520        frm.SortNow "Form_Load"  ' ** Form Procedure: frmStateCode_Sub.
6530      End If
6540      DoEvents
6550      If .chkIncludeTerritories = True Then
6560        strStateCode = frm.FindCode("TR")  ' ** Form Function: frmStateCode_Sub.
6570        If strStateCode <> vbNullString Then
6580          frm.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStateCode_Sub.
6590          frm.MoveRec 0, strStateCode  ' ** Form Procedure: frmStateCode_Sub.
6600        End If
6610      End If
6620      .frmStateCode_Sub.SetFocus
6630    End With

EXITP:
6640    Set frm = Nothing
6650    Exit Sub

ERRH:
6660    Select Case ERR.Number
        Case Else
6670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6680    End Select
6690    Resume EXITP

End Sub

Private Sub chkIncludeTerritories_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeTerritories_KeyDown"

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
6790          .chkIncludeCanada.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6850        With Me
6860          intRetVal = 0
6870          If .cmdPrintReport.Enabled = True Then
6880            .cmdPrintReport.SetFocus
6890          Else
6900            .frmStateCode_Sub.SetFocus
6910            .frmStateCode_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStateCode_Sub.
6920            .frmStateCode_Sub.Form.FocusHolder.SetFocus
6930          End If
6940        End With
6950      End Select
6960    End If

EXITP:
6970    KeyCode = intRetVal
6980    Exit Sub

ERRH:
6990    intRetVal = 0
7000    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7010    Case Else
7020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7030    End Select
7040    Resume EXITP

End Sub

Private Sub chkIncludeTerritories_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeTerritories_MouseMove"

7110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7120    Exit Sub

ERRH:
7130    Select Case ERR.Number
        Case Else
7140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7150    End Select
7160    Resume EXITP

End Sub

Private Sub chkIncludeCanada_AfterUpdate()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCanada_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String, blnNoFilter As Boolean
        Dim strStateCode As String
        Dim blnSortHere As Boolean, blnResort As Boolean

7210    blnNoFilter = False

7220    With Me
7230      Set frm = .frmStateCode_Sub.Form
7240      blnSortHere = False: blnResort = False
7250      Select Case .chkIncludeCanada
          Case True
7260        .chkIncludeCanada_lbl.FontBold = True
7270        frm.state_canada_lbl.ForeColor = CLR_DKGRY2
7280        frm.state_canada_lbl_dim_hi.Visible = False
7290        frm.state_canada_lbl2.ForeColor = CLR_DKGRY2
7300        frm.state_canada_lbl2_dim_hi.Visible = False
7310        frm.state_canada_lbl_line.BorderColor = CLR_DKGRY
7320        frm.state_canada_lbl_line_dim_hi.Visible = False
7330        frm.state_canada.Locked = True
7340        frm.state_canada_box.BorderColor = CLR_GRY
7350        frm.state_canada_box.BackStyle = acBackStyleNormal
7360        If .chkIncludeTerritories = True And .chkIncludeMilitary = True Then
7370          strFilter = vbNullString
7380          blnNoFilter = True
7390        ElseIf .chkIncludeTerritories = True And .chkIncludeMilitary = False Then
7400          strFilter = "[state_military] = False"
7410        ElseIf .chkIncludeTerritories = False And .chkIncludeMilitary = True Then
7420          strFilter = "[state_territory] = False"
7430        ElseIf .chkIncludeTerritories = False And .chkIncludeMilitary = False Then
7440          strFilter = "[state_territory] = False And [state_military] = False"
7450        End If
7460      Case False
7470        .chkIncludeCanada_lbl.FontBold = False
7480        If frm.Sort_line2a.Left = frm.state_canada_lbl.Left And frm.Sort_line2a.Visible = True Then blnSortHere = True
7490        frm.state_canada_lbl.ForeColor = WIN_CLR_DISF
7500        frm.state_canada_lbl_dim_hi.Visible = True
7510        frm.state_canada_lbl2.ForeColor = WIN_CLR_DISF
7520        frm.state_canada_lbl2_dim_hi.Visible = True
7530        frm.state_canada_lbl_line.BorderColor = WIN_CLR_DISR
7540        frm.state_canada_lbl_line_dim_hi.Visible = True
7550        frm.state_canada.Locked = False  ' ** It's already disabled.
7560        frm.state_canada_box.BorderColor = WIN_CLR_DISR
7570        frm.state_canada_box.BackStyle = acBackStyleTransparent
7580        If blnSortHere = True Then
7590          frm.Sort_line2a.Visible = False
7600          frm.Sort_line2b.Visible = False
7610          blnResort = True
7620        End If
7630        blnSortHere = False
7640        If .chkIncludeTerritories = True And .chkIncludeMilitary = True Then
7650          strFilter = "[state_canada] = False"
7660        ElseIf .chkIncludeTerritories = True And .chkIncludeMilitary = False Then
7670          strFilter = "[state_canada] = False And [state_military] = False"
7680        ElseIf .chkIncludeTerritories = False And .chkIncludeMilitary = True Then
7690          strFilter = "[state_territory] = False And [state_canada] = False"
7700        ElseIf .chkIncludeTerritories = False And .chkIncludeMilitary = False Then
7710          strFilter = "[state_territory] = False And [state_canada] = False And [state_military] = False"
7720        End If
7730      End Select
7740      frm.Filter = strFilter
7750      Select Case blnNoFilter
          Case True
7760        frm.FilterOn = False
7770      Case False
7780        frm.FilterOn = True
7790      End Select
7800      frm.Requery
7810      If blnResort = True Then
7820        frm.SortNow "Form_Load"  ' ** Form Procedure: frmStateCode_Sub.
7830      End If
7840      DoEvents
7850      If .chkIncludeCanada = True Then
7860        strStateCode = frm.FindCode("CN")  ' ** Form Function: frmStateCode_Sub.
7870        If strStateCode <> vbNullString Then
7880          frm.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStateCode_Sub.
7890          frm.MoveRec 0, strStateCode  ' ** Form Procedure: frmStateCode_Sub.
7900        End If
7910      End If
7920      .frmStateCode_Sub.SetFocus
7930    End With

EXITP:
7940    Set frm = Nothing
7950    Exit Sub

ERRH:
7960    Select Case ERR.Number
        Case Else
7970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7980    End Select
7990    Resume EXITP

End Sub

Private Sub chkIncludeCanada_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCanada_MouseMove"

8010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8020    Exit Sub

ERRH:
8030    Select Case ERR.Number
        Case Else
8040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8050    End Select
8060    Resume EXITP

End Sub

Private Sub chkIncludeMilitary_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeMilitary_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String, blnNoFilter As Boolean
        Dim strStateCode As String
        Dim blnSortHere As Boolean, blnResort As Boolean

8110    blnNoFilter = False

8120    With Me
8130      Set frm = .frmStateCode_Sub.Form
8140      blnSortHere = False: blnResort = False
8150      Select Case .chkIncludeMilitary
          Case True
8160        .chkIncludeMilitary_lbl.FontBold = True
8170        frm.state_military_lbl.ForeColor = CLR_DKGRY2
8180        frm.state_military_lbl_dim_hi.Visible = False
8190        frm.state_military_lbl2.ForeColor = CLR_DKGRY2
8200        frm.state_military_lbl2_dim_hi.Visible = False
8210        frm.state_military_lbl_line.BorderColor = CLR_DKGRY
8220        frm.state_military_lbl_line_dim_hi.Visible = False
8230        frm.state_military.Locked = True
8240        frm.state_military_box.BorderColor = CLR_GRY
8250        frm.state_military_box.BackStyle = acBackStyleNormal
8260        If .chkIncludeTerritories = True And .chkIncludeCanada = True Then
8270          strFilter = vbNullString
8280          blnNoFilter = True
8290        ElseIf .chkIncludeTerritories = True And .chkIncludeCanada = False Then
8300          strFilter = "[state_canada] = False"
8310        ElseIf .chkIncludeTerritories = False And .chkIncludeCanada = True Then
8320          strFilter = "[state_territory] = False"
8330        ElseIf .chkIncludeTerritories = False And .chkIncludeCanada = False Then
8340          strFilter = "[state_territory] = False And [state_canada] = False"
8350        End If
8360      Case False
8370        .chkIncludeMilitary_lbl.FontBold = False
8380        If frm.Sort_line2a.Left = frm.state_military_lbl.Left And frm.Sort_line2a.Visible = True Then blnSortHere = True
8390        frm.state_military_lbl.ForeColor = WIN_CLR_DISF
8400        frm.state_military_lbl_dim_hi.Visible = True
8410        frm.state_military_lbl2.ForeColor = WIN_CLR_DISF
8420        frm.state_military_lbl2_dim_hi.Visible = True
8430        frm.state_military_lbl_line.BorderColor = WIN_CLR_DISR
8440        frm.state_military_lbl_line_dim_hi.Visible = True
8450        frm.state_military.Locked = False  ' ** It's already disabled.
8460        frm.state_military_box.BorderColor = WIN_CLR_DISR
8470        frm.state_military_box.BackStyle = acBackStyleTransparent
8480        If blnSortHere = True Then
8490          frm.Sort_line2a.Visible = False
8500          frm.Sort_line2b.Visible = False
8510          blnResort = True
8520        End If
8530        blnSortHere = False
8540        If .chkIncludeTerritories = True And .chkIncludeCanada = True Then
8550          strFilter = "[state_military] = False"
8560        ElseIf .chkIncludeTerritories = True And .chkIncludeCanada = False Then
8570          strFilter = "[state_canada] = False And [state_military] = False"
8580        ElseIf .chkIncludeTerritories = False And .chkIncludeCanada = True Then
8590          strFilter = "[state_territory] = False And [state_military] = False"
8600        ElseIf .chkIncludeTerritories = False And .chkIncludeCanada = False Then
8610          strFilter = "[state_territory] = False And [state_canada] = False And [state_military] = False"
8620        End If
8630      End Select
8640      frm.Filter = strFilter
8650      Select Case blnNoFilter
          Case True
8660        frm.FilterOn = False
8670      Case False
8680        frm.FilterOn = True
8690      End Select
8700      frm.Requery
8710      If blnResort = True Then
8720        frm.SortNow "Form_Load"  ' ** Form Procedure: frmStateCode_Sub.
8730      End If
8740      DoEvents
8750      If .chkIncludeMilitary = True Then
8760        strStateCode = frm.FindCode("ML")  ' ** Form Function: frmStateCode_Sub.
8770        If strStateCode <> vbNullString Then
8780          frm.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStateCode_Sub.
8790          frm.MoveRec 0, strStateCode  ' ** Form Procedure: frmStateCode_Sub.
8800        End If
8810      End If
8820      .frmStateCode_Sub.SetFocus
8830    End With

EXITP:
8840    Set frm = Nothing
8850    Exit Sub

ERRH:
8860    Select Case ERR.Number
        Case Else
8870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8880    End Select
8890    Resume EXITP

End Sub

Private Sub chkIncludeMilitary_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeMilitary_MouseMove"

8910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8920    Exit Sub

ERRH:
8930    Select Case ERR.Number
        Case Else
8940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8950    End Select
8960    Resume EXITP

End Sub
