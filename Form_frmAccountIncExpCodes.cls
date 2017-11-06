VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountIncExpCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountIncExpCodes"

'VGC 04/15/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   ShowCost()
' ##   IncludeCurrency()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Block Assign:     B {cmdBlockAssign}
' **   Close:            C {cmdClose}
' **   Include Archive:  H {chkIncludeArchive}
' **   Include Currency: L {chkIncludeCurrency}
' **   Show Cost:        O {chkShowCost}
' **   Shorten:          S {cmdShorten on frmAccountIncExpCodes_Sub}
' **   Update:           U {cmdUpdate}
' **   Widen:            W {cmdWiden on frmAccountIncExpCodes_Sub}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAccountIncExpCodes_Sub}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngCost_Width As Long, lngCurrID_Width As Long, lngSub_Width As Long
Private lngClose_Left As Long, lngUpdate_Left As Long, lngBlock_Left As Long

Private blnHasForEx_Chk As Boolean, lngOutOfSync As Long
Private blnAskedToAbandon As Boolean, blnViewOnly As Boolean, blnHasArchive As Boolean, blnThisArchive As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim usr As DAO.User, grp As DAO.Group, frm As Access.Form
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           lngOutOfSync = 0&
160           strCallingForm = Trim(.OpenArgs)
170           lngOutOfSync = CLng(Mid(strCallingForm, (InStr(strCallingForm, "^") + 1)))
              ' ** OutOfSync: CLng(IIf(IsNull([taxcode])=True Or IsNull([revcode_TYPE])=True,1,
              ' **   IIf(([taxcode]>0 And [taxcode]<10) And [revcode_TYPE]=2,1,
              ' **   IIf([taxcode]>=10 And [revcode_TYPE]=1,1,0))))
180           strCallingForm = Left(strCallingForm, (InStr(strCallingForm, "^") - 1))

190           blnAskedToAbandon = False: blnHasArchive = False: blnThisArchive = False
200           gblnHasForExThis = False: blnHasForEx_Chk = False

              ' ** Variables are fed empty, then populated ByRef.
210           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

220           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

230           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
240           If lngMonitorCnt > 1& Then
250             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
260             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
270             If lngMonitorNum > 1& Then
280               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
290             End If
300           End If

310           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
320           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
              'SystemColor_Move Me  ' ** Module Procedure: modSystemColorFuncs.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
330           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

340           DoEvents

350           Set frm = .frmAccountIncExpCodes_Sub.Form

360           .form_width_line.Left = 0&  ' ** Just get it out of the way.

370           lngForm_Width = .Width
380           lngCurrID_Width = ((frm.curr_id.Left + frm.curr_id.Width) - (frm.cost_str.Left + frm.cost_str.Width))
390           lngCost_Width = ((frm.cost_str.Left + frm.cost_str.Width) - (frm.pcash_str.Left + frm.pcash_str.Width))
400           lngClose_Left = .cmdClose.Left
410           lngUpdate_Left = .cmdUpdate.Left
420           lngBlock_Left = .cmdBlockAssign.Left
430           lngSub_Width = .frmAccountIncExpCodes_Sub.Width

440           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

450           .accountno = gstrAccountNo
460           .shortname = gstrAccountName

470           SetArchiveOption  ' ** Procedure: Below.

480           gblnHasForExThis = HasForEx_Acct(gstrAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.
490           Select Case gblnHasForEx
              Case True
500             Select Case gblnHasForExThis
                Case True
510               If Pref_HasPref(THIS_NAME, "chkIncludeCurrency") = False Then  ' ** Module Procedure: modPreferenceFuncs.
520                 .chkIncludeCurrency = True
530               End If
540               .chkIncludeCurrency.Enabled = True
550             Case False
560               .chkIncludeCurrency.Enabled = True
570             End Select
580           Case False
590             .chkIncludeCurrency = False
600             .chkIncludeCurrency.Enabled = False
610           End Select
620           chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
630           chkShowCost_AfterUpdate  ' ** Procedure: Below.
640           DoEvents

650           chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
660           DoEvents

670           .frmAccountIncExpCodes_Sub.SetFocus
680           frm.FocusHolder.SetFocus
690           .cmdUpdate.Enabled = False
700           DoEvents

              ' ** Check user status, and lock all fields if ViewOnly.
710           blnViewOnly = False
720           For Each grp In DBEngine.Workspaces(0).Groups
730             If grp.Name = "ViewOnly" Then
740               For Each usr In grp.Users
750                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
760                   blnViewOnly = True
770                   Exit For
780                 End If
790               Next
800             End If
810           Next
820           DoEvents

830           If blnViewOnly = True Then
840             frm.revcode_ID.Locked = True
850           End If

860           DoCmd.Hourglass False

870           blnIsOpen = False

880         Else
890           Cancel = -1
900         End If
910       Else
920         Cancel = -1
930       End If
940     End With

950     If Cancel = -1 Then
960       Beep
970       DoCmd.Hourglass False
980       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
990     End If

EXITP:
1000    Set frm = Nothing
1010    Set usr = Nothing
1020    Set grp = Nothing
1030    Exit Sub

ERRH:
1040    DoCmd.Hourglass False
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Sub

Private Sub Form_Load()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1110    With Me

1120      CLR_DISABLED_FG = CLR_DKGRY
1130      CLR_DISABLED_BG = CLR_LTTEAL

1140      .accountno.ForeColor = CLR_DISABLED_FG
1150      .accountno.BackColor = CLR_DISABLED_BG
1160      .shortname.ForeColor = CLR_DISABLED_FG
1170      .shortname.BackColor = CLR_DISABLED_BG
1180      DoEvents

1190    End With

EXITP:
1200    Exit Sub

ERRH:
1210    DoCmd.Hourglass False
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAccountIncExpCodes_Sub}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyEscape
1370        intRetVal = 0
1380        cmdClose_Click  ' ** Procedure: Below.
1390      End Select
1400    End If

        ' ** Alt keys.
1410    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1420      Select Case intRetVal
          Case vbKeyX
1430        intRetVal = 0
1440        cmdClose_Click  ' ** Procedure: Below.
1450      End Select
1460    End If

        ' ** Ctrl keys.
1470    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1480      Select Case intRetVal
          Case vbKeyS
1490        With Me
1500          intRetVal = 0
1510          lngRecsCur = .frmAccountIncExpCodes_Sub.Form.RecCnt  ' ** Form Function: frmAccountIncExpCodes_Sub.
1520          If lngRecsCur > 0& Then
1530            .frmAccountIncExpCodes_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
1540          Else
1550            Beep
1560          End If
1570        End With
1580      End Select
1590    End If

        ' ** Ctrl-Shift keys.
1600    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1610      Select Case intRetVal
          Case vbKeyF
1620        With Me
1630          intRetVal = 0
1640          .FocusHolder.SetFocus
1650        End With
1660      End Select
1670    End If

EXITP:
1680    KeyCode = intRetVal
1690    Exit Sub

ERRH:
1700    intRetVal = 0
1710    Select Case ERR.Number
        Case Else
1720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1730    End Select
1740    Resume EXITP

End Sub

Private Sub Form_Timer()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1810    With Me
1820      .TimerInterval = 0&
1830      If gblnSetFocus = True Then
1840        gblnSetFocus = False
1850        strFormName = vbNullString
1860  On Error Resume Next
1870        Do Until strFormName = THIS_NAME
1880          DoCmd.SelectObject acForm, THIS_NAME, False
1890          strFormName = Screen.ActiveForm.Name
1900        Loop
1910  On Error GoTo ERRH
1920        DoCmd.Hourglass False
1930      End If
1940      If gstrReturningForm <> vbNullString Then
1950        Select Case gstrReturningForm
            Case "frmAccountIncExpCodes_BlockAssign"
1960          .cmdBlockAssign.SetFocus
1970        Case Else
              ' ** Nothing else.
1980        End Select
1990        gstrReturningForm = vbNullString
2000      End If
2010      DoCmd.Hourglass False
2020    End With

EXITP:
2030    Exit Sub

ERRH:
2040    DoCmd.Hourglass False
2050    Select Case ERR.Number
        Case Else
2060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2070    End Select
2080    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean
        Dim blnContinue As Boolean

2110    With Me

2120      DoCmd.Hourglass True
2130      DoEvents

2140      blnContinue = True

2150      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2160      If .chkUpdated = True And blnAskedToAbandon = False Then
2170        Beep
2180        DoCmd.Hourglass False
2190        If MsgBox("You have made changes, but have not confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbYes Then
              ' ** Abandon changes.
2200        Else
2210          blnContinue = False
2220        End If
2230      End If

2240      If blnContinue = True Then

2250        DoCmd.Hourglass True
2260        DoEvents

2270        blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

            ' ** Make sure all reports are closed.
2280        If Reports.Count > 0 Then
2290          Do While Reports.Count > 0
2300            DoCmd.Close acReport, Reports(0).Name
2310          Loop
2320        End If

2330        gblnSetFocus = True
2340        gstrReturningForm = THIS_NAME
2350        strDocName = "frmMenu_Account"
2360        If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
2370          DoCmd.SelectObject acForm, strDocName, False
2380          Forms(strDocName).TimerInterval = 250&
2390        Else
2400          DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
2410        End If

2420      End If
2430    End With

EXITP:
2440    Exit Sub

ERRH:
2450    DoCmd.Hourglass False
2460    Select Case ERR.Number
        Case Else
2470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2480    End Select
2490    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnContinue As Boolean

2510    blnContinue = True

2520    With Me

2530      .TimerInterval = 0&

2540      If .chkUpdated = True And blnViewOnly = False Then
2550        blnAskedToAbandon = False
2560        Beep
2570        If MsgBox("You have made changes, but have not confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbYes Then
2580          blnAskedToAbandon = True
2590        Else
2600          blnContinue = False
2610        End If
2620      ElseIf blnViewOnly = True Then
2630        blnAskedToAbandon = True
2640      End If

2650      If blnContinue = True Then

2660        Set dbs = CurrentDb
            ' ** Empty tblRevCode_BlockAssign_Staging.
2670        Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_Block_06")
2680        qdf.Execute
2690        DoEvents

2700        DoCmd.Close acForm, THIS_NAME

2710      End If

2720    End With

EXITP:
2730    Set qdf = Nothing
2740    Set dbs = Nothing
2750    Exit Sub

ERRH:
2760    Select Case ERR.Number
        Case Else
2770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2780    End Select
2790    Resume EXITP

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
2890          .frmAccountIncExpCodes_Sub.SetFocus
2900          .frmAccountIncExpCodes_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
2910  On Error Resume Next
2920          .frmAccountIncExpCodes_Sub.Form.revcode_DESC.SetFocus
2930          If ERR.Number <> 0 Then
2940  On Error GoTo ERRH
2950            .frmAccountIncExpCodes_Sub.Form.revcode_ID.SetFocus
2960          Else
2970  On Error GoTo ERRH
2980          End If
2990        End With
3000      End Select
3010    End If

        ' ** Shift keys.
3020    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3030      Select Case intRetVal
          Case vbKeyTab
3040        With Me
3050          intRetVal = 0
3060          If .cmdUpdate.Enabled = True Then
3070            .cmdUpdate.SetFocus
3080          Else
3090            .cmdBlockAssign.SetFocus
3100          End If
3110        End With
3120      End Select
3130    End If

EXITP:
3140    KeyCode = intRetVal
3150    Exit Sub

ERRH:
3160    intRetVal = 0
3170    Select Case ERR.Number
        Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

        Dim strFilter As String

3310    With Me
3320      If blnThisArchive = True Then
3330        Select Case .chkIncludeArchive
            Case True
3340          .chkIncludeArchive_lbl.FontBold = True
3350          .archive_trans_lbl.Visible = True
3360          .archive_trans_lbl_box.Visible = True
3370          strFilter = vbNullString
3380          .frmAccountIncExpCodes_Sub.Form.Filter = strFilter
3390          .frmAccountIncExpCodes_Sub.Form.FilterOn = False
3400        Case False
3410          .chkIncludeArchive_lbl.FontBold = False
3420          .archive_trans_lbl.Visible = False
3430          .archive_trans_lbl_box.Visible = False
3440          strFilter = "[src] = 'L'"
3450          .frmAccountIncExpCodes_Sub.Form.Filter = strFilter
3460          .frmAccountIncExpCodes_Sub.Form.FilterOn = True
3470        End Select
3480        .frmAccountIncExpCodes_Sub.Form.Requery
3490        .frmAccountIncExpCodes_Sub.SetFocus
3500        .frmAccountIncExpCodes_Sub.Form.FocusHolder.SetFocus
3510      End If
3520    End With

EXITP:
3530    Exit Sub

ERRH:
3540    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
3550    Case Else
3560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3570    End Select
3580    Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

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
3690          If .chkIncludeCurrency.Enabled = True Then
3700            .chkIncludeCurrency.SetFocus
3710          Else
3720            .chkShowCost.SetFocus
3730          End If
3740        End With
3750      End Select
3760    End If

        ' ** Shift keys.
3770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3780      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3790        With Me
3800          intRetVal = 0
3810          .frmAccountIncExpCodes_Sub.SetFocus
3820          .frmAccountIncExpCodes_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
3830  On Error Resume Next
3840          .frmAccountIncExpCodes_Sub.Form.revcode_DESC.SetFocus
3850          If ERR.Number <> 0 Then
3860  On Error GoTo ERRH
3870            .frmAccountIncExpCodes_Sub.Form.revcode_ID.SetFocus
3880          Else
3890  On Error GoTo ERRH
3900          End If
3910        End With
3920      End Select
3930    End If

EXITP:
3940    KeyCode = intRetVal
3950    Exit Sub

ERRH:
3960    intRetVal = 0
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

4110    With Me
4120      Select Case .chkIncludeCurrency
          Case True
4130        .chkIncludeCurrency_lbl.FontBold = True
4140        .frmAccountIncExpCodes_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
4150        IncludeCurrency True  ' ** Procedure: Below.
4160      Case False
4170        .chkIncludeCurrency_lbl.FontBold = False
4180        .frmAccountIncExpCodes_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
4190        IncludeCurrency False  ' ** Procedure: Below.
4200      End Select
4210      .frmAccountIncExpCodes_Sub.SetFocus
4220    End With

EXITP:
4230    Exit Sub

ERRH:
4240    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
4250    Case Else
4260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4270    End Select
4280    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

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
4390          .chkShowCost.SetFocus
4400        End With
4410      End Select
4420    End If

        ' ** Shift keys.
4430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4450        With Me
4460          intRetVal = 0
4470          If .chkIncludeArchive.Enabled = True Then
4480            .chkIncludeArchive.SetFocus
4490          Else
4500            .frmAccountIncExpCodes_Sub.SetFocus
4510            .frmAccountIncExpCodes_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
4520  On Error Resume Next
4530            .frmAccountIncExpCodes_Sub.Form.revcode_DESC.SetFocus
4540            If ERR.Number <> 0 Then
4550  On Error GoTo ERRH
4560              .frmAccountIncExpCodes_Sub.Form.revcode_ID.SetFocus
4570            Else
4580  On Error GoTo ERRH
4590            End If
4600          End If
4610        End With
4620      End Select
4630    End If

EXITP:
4640    KeyCode = intRetVal
4650    Exit Sub

ERRH:
4660    intRetVal = 0
4670    Select Case ERR.Number
        Case Else
4680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4690    End Select
4700    Resume EXITP

End Sub

Public Sub chkShowCost_AfterUpdate()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowCost_AfterUpdate"

4810    With Me
4820      Select Case .chkShowCost
          Case True
4830        .chkShowCost_lbl.FontBold = True
4840        .frmAccountIncExpCodes_Sub.Form.ShowCost_Sub True  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
4850        ShowCost True  ' ** Procedure: Below.
4860      Case False
4870        .chkShowCost_lbl.FontBold = False
4880        .frmAccountIncExpCodes_Sub.Form.ShowCost_Sub False  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
4890        ShowCost False  ' ** Procedure: Below.
4900      End Select
4910      .frmAccountIncExpCodes_Sub.SetFocus
4920    End With

EXITP:
4930    Exit Sub

ERRH:
4940    Select Case ERR.Number
        Case Else
4950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4960    End Select
4970    Resume EXITP

End Sub

Private Sub chkShowCost_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowCost_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5020    intShiftDown = (Shift And acShiftMask) > 0
5030    intAltDown = (Shift And acAltMask) > 0
5040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5070        With Me
5080          intRetVal = 0
5090          .cmdBlockAssign.SetFocus
5100        End With
5110      End Select
5120    End If

        ' ** Shift keys.
5130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5150        With Me
5160          intRetVal = 0
5170          If .chkIncludeCurrency.Enabled = True Then
5180            .chkIncludeCurrency.SetFocus
5190          ElseIf .chkIncludeArchive.Enabled = True Then
5200            .chkIncludeArchive.SetFocus
5210          Else
5220            .frmAccountIncExpCodes_Sub.SetFocus
5230            .frmAccountIncExpCodes_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountIncExpCodes_Sub.
5240  On Error Resume Next
5250            .frmAccountIncExpCodes_Sub.Form.revcode_DESC.SetFocus
5260            If ERR.Number <> 0 Then
5270  On Error GoTo ERRH
5280              .frmAccountIncExpCodes_Sub.Form.revcode_ID.SetFocus
5290            Else
5300  On Error GoTo ERRH
5310            End If
5320          End If
5330        End With
5340      End Select
5350    End If

EXITP:
5360    KeyCode = intRetVal
5370    Exit Sub

ERRH:
5380    intRetVal = 0
5390    Select Case ERR.Number
        Case Else
5400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5410    End Select
5420    Resume EXITP

End Sub

Private Sub cmdBlockAssign_Click()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBlockAssign_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim blnArchive As Boolean

5510    With Me

5520      DoCmd.Hourglass True
5530      DoEvents

5540      Set dbs = CurrentDb

5550      blnArchive = False

          ' ** Empty tblRevCode_BlockAssign_Staging.
5560      Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_Block_06")
5570      qdf.Execute
5580      Set qdf = Nothing
5590      DoEvents

5600      blnArchive = .chkIncludeArchive

5610      If blnArchive = True Then
            ' ** Append qryAccountProfile_IncExpCodes_Block_03_03 (LedgerArchive,
            ' ** as tblRevCode_BlockAssign_Staging records, by specified [actno])
            ' ** to tblRevCode_BlockAssign_Staging.
5620        Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_Block_03_04")
5630        With qdf.Parameters
5640          ![actno] = gstrAccountNo
5650        End With
5660        qdf.Execute
5670        Set qdf = Nothing
5680      End If
5690      DoEvents

          ' ** Append qryAccountProfile_IncExpCodes_Block_03_01 (Ledger,
          ' ** as tblRevCode_BlockAssign_Staging records, by specified [actno])
          ' ** to tblRevCode_BlockAssign_Staging.
5700      Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_Block_03_02")
5710      With qdf.Parameters
5720        ![actno] = gstrAccountNo
5730      End With
5740      qdf.Execute
5750      Set qdf = Nothing
5760      DoEvents

5770      dbs.Close
5780      Set dbs = Nothing

5790      strDocName = "frmAccountIncExpCodes_BlockAssign"
5800      gblnSetFocus = True
5810      DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & gstrAccountNo & "~" & Format(blnArchive, "True/False")

5820    End With

EXITP:
5830    Set qdf = Nothing
5840    Set dbs = Nothing
5850    Exit Sub

ERRH:
5860    DoCmd.Hourglass False
5870    Select Case ERR.Number
        Case Else
5880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5890    End Select
5900    Resume EXITP

End Sub

Private Sub cmdBlockAssign_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBlockAssign_KeyDown"

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
          Case vbKeyTab
6070        With Me
6080          intRetVal = 0
6090          If .cmdUpdate.Enabled = True Then
6100            .cmdUpdate.SetFocus
6110          Else
6120            .cmdClose.SetFocus
6130          End If
6140        End With
6150      End Select
6160    End If

        ' ** Shift keys.
6170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6180      Select Case intRetVal
          Case vbKeyTab
6190        With Me
6200          intRetVal = 0
6210          .chkShowCost.SetFocus
6220        End With
6230      End Select
6240    End If

EXITP:
6250    KeyCode = intRetVal
6260    Exit Sub

ERRH:
6270    intRetVal = 0
6280    Select Case ERR.Number
        Case Else
6290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6300    End Select
6310    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, qdf3 As DAO.QueryDef
        Dim rstTmpEdit As DAO.Recordset, rstLedger As DAO.Recordset, rstArchive As DAO.Recordset
        Dim strSource As String
        Dim lngRecs1 As Long, lngRecs2 As Long, lngRecs3 As Long
        Dim blnInTrans As Boolean, blnContinue As Boolean
        Dim lngX As Long

6410    With Me
6420      If blnViewOnly = False Then

6430        blnInTrans = False: blnContinue = True

6440        .frmAccountIncExpCodes_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountIncExpCodes_Sub.

6450        Set wrk = DBEngine.Workspaces(0)
6460        Set dbs = wrk.Databases(0)

            ' ** Set out-of-sync transactions as changed.
            ' ** If any of these have already been fixed,
            ' ** they'll no longer show up in this query.
6470        If lngOutOfSync > 0& And gblnLinkRevTaxCodes = True Then
              ' ** Update qryAccountProfile_IncExpCodes_11 (tmpEdit06, for OutOfSync = 1,
              ' ** with taxcode_new, revcode_ID_new, revcode_TYPE_new, revcode_DESC_new).
6480          Set qdf1 = dbs.QueryDefs("qryAccountProfile_IncExpCodes_13")
6490          qdf1.Execute
6500          Set qdf1 = Nothing
6510        End If

            ' ** Table tmpEdit06, all fields, changed = True.
6520        Set qdf1 = dbs.QueryDefs("qryAccountProfile_IncExpCodes_02")
6530        Set rstTmpEdit = qdf1.OpenRecordset
6540        If rstTmpEdit.BOF = True And rstTmpEdit.EOF = True Then
              ' ** The button shouldn't have been available if there were no changes!
6550          Beep
6560          MsgBox "No changes were found.", vbInformation + vbOKOnly, "Nothing To Do"
6570        Else
6580          rstTmpEdit.MoveLast
6590          lngRecs1 = rstTmpEdit.RecordCount
6600          rstTmpEdit.MoveFirst
              ' ** Ledger, linked to tmpEdit06, by specified [actno].
6610          Set qdf2 = dbs.QueryDefs("qryAccountProfile_IncExpCodes_03")
6620          With qdf2.Parameters
6630            ![actno] = gstrAccountNo
6640          End With
6650          Set rstLedger = qdf2.OpenRecordset
              ' ** LedgerArchive, linked to tmpEdit6, by specified [actno].
6660          Set qdf3 = dbs.QueryDefs("qryAccountProfile_IncExpCodes_03a")
6670          With qdf3.Parameters
6680            ![actno] = gstrAccountNo
6690          End With
6700          Set rstArchive = qdf3.OpenRecordset
6710          If rstLedger.BOF = True And rstLedger.EOF = True Then
                ' ** I don't think it could get here.
6720            If rstArchive.BOF = True And rstArchive.EOF = True Then
6730              blnContinue = False
6740            End If
6750          End If
6760          If blnContinue = True Then
6770            rstLedger.MoveLast
6780            lngRecs2 = rstLedger.RecordCount
6790            rstLedger.MoveFirst
6800            If rstArchive.BOF = True And rstArchive.EOF = True Then
6810              lngRecs3 = 0&
6820            Else
6830              rstArchive.MoveLast
6840              lngRecs3 = rstArchive.RecordCount
6850              rstArchive.MoveFirst
6860            End If
6870            If lngRecs1 > 0& And (lngRecs2 > 0& Or lngRecs3 > 0&) Then
6880              wrk.BeginTrans
6890              blnInTrans = True
6900              rstTmpEdit.MoveFirst
6910              For lngX = 1& To lngRecs1
6920                blnContinue = True
6930                strSource = rstTmpEdit![src]
6940                Select Case strSource
                    Case "A"
6950                  rstArchive.FindFirst "journalno = " & CStr(rstTmpEdit![journalno])
6960                  If rstArchive.NoMatch = True Then
6970                    blnContinue = False
6980                    MsgBox "Transaction not found in Ledger Archive.", vbCritical + vbOKOnly, "Record Not Found"
6990                  End If
7000                Case "L"
7010                  rstLedger.FindFirst "journalno = " & CStr(rstTmpEdit![journalno])
7020                  If rstLedger.NoMatch = True Then
7030                    blnContinue = False
7040                    MsgBox "Transaction not found in ledger.", vbCritical + vbOKOnly, "Record Not Found"
7050                  End If
7060                End Select
7070                If blnContinue = True Then
7080                  If IsNull(rstTmpEdit![revcode_ID]) = False Then
7090                    Select Case strSource
                        Case "A"
7100                      If IsNull(rstArchive![revcode_ID]) = False Then
7110                        If rstArchive![revcode_ID] <> rstTmpEdit![revcode_ID] Then
7120                          rstArchive.Edit
7130                          rstArchive![revcode_ID] = rstTmpEdit![revcode_ID]
7140                          rstArchive.Update
7150                        End If
7160                      Else
7170                        rstArchive.Edit
7180                        rstArchive![revcode_ID] = rstTmpEdit![revcode_ID]
7190                        rstArchive.Update
7200                      End If
7210                    Case "L"
7220                      If IsNull(rstLedger![revcode_ID]) = False Then
7230                        If rstLedger![revcode_ID] <> rstTmpEdit![revcode_ID] Then
7240                          rstLedger.Edit
7250                          rstLedger![revcode_ID] = rstTmpEdit![revcode_ID]
7260                          rstLedger.Update
7270                        End If
7280                      Else
7290                        rstLedger.Edit
7300                        rstLedger![revcode_ID] = rstTmpEdit![revcode_ID]
7310                        rstLedger.Update
7320                      End If
7330                    End Select
7340                    If gblnLinkRevTaxCodes = True Then
7350                      Select Case strSource
                          Case "A"
7360                        If IsNull(rstTmpEdit![taxcode]) = False Then
7370                          If IsNull(rstArchive![taxcode]) = False Then
7380                            If rstArchive![taxcode] <> rstTmpEdit![taxcode] Then
7390                              rstArchive.Edit
7400                              rstArchive![taxcode] = rstTmpEdit![taxcode]
7410                              rstArchive.Update
7420                            End If
7430                          Else
7440                            rstArchive.Edit
7450                            rstArchive![taxcode] = rstTmpEdit![taxcode]
7460                            rstArchive.Update
7470                          End If
7480                        End If
7490                      Case "L"
7500                        If IsNull(rstTmpEdit![taxcode]) = False Then
7510                          If IsNull(rstLedger![taxcode]) = False Then
7520                            If rstLedger![taxcode] <> rstTmpEdit![taxcode] Then
7530                              rstLedger.Edit
7540                              rstLedger![taxcode] = rstTmpEdit![taxcode]
7550                              rstLedger.Update
7560                            End If
7570                          Else
7580                            rstLedger.Edit
7590                            rstLedger![taxcode] = rstTmpEdit![taxcode]
7600                            rstLedger.Update
7610                          End If
7620                        End If
7630                      End Select
7640                    End If
7650                  Else
                        ' ** If this is Null, don't make any change to the Ledger, though
                        ' ** I don't think they could ever succeed in Nulling out revcode_ID.
7660                  End If
7670                End If  ' ** blnContinue.
7680                If lngX < lngRecs1 Then rstTmpEdit.MoveNext
7690              Next
7700              wrk.CommitTrans
7710              blnInTrans = False
7720              .chkUpdated = False
7730              .cmdClose.SetFocus
7740              .cmdUpdate.Enabled = False
7750              MsgBox "Updates completed.", vbInformation + vbOKOnly, ("Update Successful" & Space(40))
7760            End If
7770          End If  ' ** blnContinue.
7780          rstLedger.Close
7790          rstArchive.Close
7800        End If
7810        rstTmpEdit.Close
7820        dbs.Close
7830        wrk.Close

7840      End If  ' ** blnViewOnly.
7850    End With

EXITP:
7860    Set rstLedger = Nothing
7870    Set rstArchive = Nothing
7880    Set rstTmpEdit = Nothing
7890    Set qdf1 = Nothing
7900    Set qdf2 = Nothing
7910    Set qdf3 = Nothing
7920    Set dbs = Nothing
7930    Set wrk = Nothing
7940    Exit Sub

ERRH:
7950    If blnInTrans = True Then
7960      wrk.Rollback
7970      blnInTrans = False
7980      rstTmpEdit.Close
7990      rstLedger.Close
8000      dbs.Close
8010      wrk.Close
8020    End If
8030    Select Case ERR.Number
        Case Else
8040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8050    End Select
8060    Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

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
8190          .cmdClose.SetFocus
8200        End With
8210      End Select
8220    End If

        ' ** Shift keys.
8230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8240      Select Case intRetVal
          Case vbKeyTab
8250        With Me
8260          intRetVal = 0
8270          .cmdBlockAssign.SetFocus
8280        End With
8290      End Select
8300    End If

EXITP:
8310    KeyCode = intRetVal
8320    Exit Sub

ERRH:
8330    intRetVal = 0
8340    Select Case ERR.Number
        Case Else
8350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8360    End Select
8370    Resume EXITP

End Sub

Private Sub SetArchiveOption()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "SetArchiveOption"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

8410    With Me

8420      blnHasArchive = False: blnThisArchive = False

8430      Set dbs = CurrentDb
8440      With dbs

8450        Set rst = .OpenRecordset("LedgerArchive", dbOpenDynaset, dbReadOnly)
8460        With rst
8470          If .BOF = True And .EOF = True Then
                ' ** No archive whatsoever!
8480          Else
8490            .MoveLast
8500            If .RecordCount > 1 Then
8510              blnHasArchive = True
8520            End If
8530          End If
8540          .Close
8550        End With
8560        Set rst = Nothing
8570        DoEvents

8580        If blnHasArchive = True Then
              ' ** LedgerArchive, with 'Liability (+)'/'Liability (-)' broken out,
              ' ** with totdescx, by specified GlobalVarGet('gstrAccountNo').
8590          Set qdf = .QueryDefs("qryAccountProfile_IncExpCodes_04")
8600          Set rst = qdf.OpenRecordset
8610          With rst
8620            If .BOF = True And .EOF = True Then
                  ' ** This account has no archive.
8630            Else
8640              .MoveLast
8650              If .RecordCount > 0 Then
8660                blnThisArchive = True
8670              End If
8680            End If
8690            .Close
8700          End With
8710          Set rst = Nothing
8720          Set qdf = Nothing
8730        End If  ' ** blnHasArchive.
8740        DoEvents

8750        .Close
8760      End With
8770      Set dbs = Nothing

8780      .archive_trans_lbl.Visible = False
8790      .archive_trans_lbl_box.Visible = False

8800      Select Case blnHasArchive
          Case True
8810        Select Case blnThisArchive
            Case True
8820          .chkIncludeArchive.Enabled = True
8830          .chkIncludeArchive_lbl.Visible = True
8840          .chkIncludeArchive_lbl4_dim.Visible = False
8850          .chkIncludeArchive_lbl4_dim_hi.Visible = False
8860          .chkIncludeArchive_lbl5_dim.Visible = False
8870          .chkIncludeArchive_lbl5_dim_hi.Visible = False
8880        Case False
8890          .chkIncludeArchive.Enabled = False
8900          .chkIncludeArchive_lbl.Visible = False
8910          .chkIncludeArchive_lbl4_dim.Visible = True
8920          .chkIncludeArchive_lbl4_dim_hi.Visible = True
8930          .chkIncludeArchive_lbl5_dim.Visible = True
8940          .chkIncludeArchive_lbl5_dim_hi.Visible = True
8950        End Select
8960        .chkIncludeArchive_lbl2_dim.Visible = False
8970        .chkIncludeArchive_lbl2_dim_hi.Visible = False
8980        .chkIncludeArchive_lbl3_dim.Visible = False
8990        .chkIncludeArchive_lbl3_dim_hi.Visible = False
9000      Case False
9010        .chkIncludeArchive.Enabled = False
9020        .chkIncludeArchive_lbl.Visible = False
9030        .chkIncludeArchive_lbl4_dim.Visible = False
9040        .chkIncludeArchive_lbl4_dim_hi.Visible = False
9050        .chkIncludeArchive_lbl5_dim.Visible = False
9060        .chkIncludeArchive_lbl5_dim_hi.Visible = False
9070        .chkIncludeArchive_lbl2_dim.Visible = True
9080        .chkIncludeArchive_lbl2_dim_hi.Visible = True
9090        .chkIncludeArchive_lbl3_dim.Visible = True
9100        .chkIncludeArchive_lbl3_dim_hi.Visible = True
9110      End Select
9120      DoEvents

9130    End With

EXITP:
9140    Set rst = Nothing
9150    Set qdf = Nothing
9160    Set dbs = Nothing
9170    Exit Sub

ERRH:
9180    Select Case ERR.Number
        Case Else
9190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9200    End Select
9210    Resume EXITP

End Sub

Private Sub ShowCost(blnShow As Boolean)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowCost"

        Dim lngNewFrmWidth As Long, lngNewFrmLeft As Long, lngNewFormWidth As Long
        Dim blnCurrID As Boolean
        Dim lngTmp01 As Long

9310    With Me

9320      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
9330        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9340      End If

9350      blnCurrID = .chkIncludeCurrency

9360      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
9370      lngMonitorNum = 1&: lngTmp01 = 0&
9380      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
9390      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

9400      Select Case blnShow
          Case True
9410        Select Case blnCurrID
            Case True
9420          lngNewFrmWidth = lngFrm_Width
9430          lngNewFrmLeft = lngFrm_Left
9440          lngNewFormWidth = lngForm_Width
9450        Case False
9460          lngNewFrmWidth = (lngFrm_Width - lngCurrID_Width)
9470          lngNewFrmLeft = (lngFrm_Left + (lngCurrID_Width / 2))
9480          lngNewFormWidth = (lngForm_Width - lngCurrID_Width)
9490        End Select
9500        .Width = lngNewFormWidth
9510        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
9520        DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
9530        If lngMonitorNum > 1& Then
9540          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
9550        End If
9560        Select Case blnCurrID
            Case True
9570          .cmdClose.Left = lngClose_Left
9580          .cmdUpdate.Left = lngUpdate_Left
9590          .cmdBlockAssign.Left = lngBlock_Left
9600          .frmAccountIncExpCodes_Sub.Width = lngSub_Width
9610        Case False
9620          .cmdClose.Left = (lngClose_Left - lngCurrID_Width)
9630          .cmdUpdate.Left = (lngUpdate_Left - lngCurrID_Width)
9640          .cmdBlockAssign.Left = (lngBlock_Left - lngCurrID_Width)
9650          .frmAccountIncExpCodes_Sub.Width = (lngSub_Width - lngCurrID_Width)
9660        End Select
9670        .frmAccountIncExpCodes_Sub_box.Width = (.frmAccountIncExpCodes_Sub.Width + (2& * lngTpp))
9680        .Nav_box01.Width = .frmAccountIncExpCodes_Sub.Width
9690        .Header_vline01.Left = lngNewFormWidth
9700        .Header_vline02.Left = lngNewFormWidth
9710        .Footer_vline01.Left = lngNewFormWidth
9720        .Footer_vline02.Left = lngNewFormWidth
9730        .Header_hline01.Width = lngNewFormWidth
9740        .Header_hline02.Width = lngNewFormWidth
9750        .Footer_hline01.Width = lngNewFormWidth
9760        .Footer_hline02.Width = lngNewFormWidth
9770        .Width = lngNewFormWidth
9780      Case False
9790        Select Case blnCurrID
            Case True
9800          lngNewFrmWidth = (lngFrm_Width - lngCost_Width)
9810          lngNewFrmLeft = (lngFrm_Left + (lngCost_Width / 2))
9820          lngNewFormWidth = (lngForm_Width - lngCost_Width)
9830        Case False
9840          lngNewFrmWidth = (lngFrm_Width - (lngCost_Width + lngCurrID_Width))
9850          lngNewFrmLeft = (lngFrm_Left + ((lngCost_Width + lngCurrID_Width) / 2))
9860          lngNewFormWidth = (lngForm_Width - (lngCost_Width + lngCurrID_Width))
9870        End Select
9880        Select Case blnCurrID
            Case True
9890          .cmdClose.Left = (lngClose_Left - lngCost_Width)
9900          .cmdUpdate.Left = (lngUpdate_Left - lngCost_Width)
9910          .cmdBlockAssign.Left = (lngBlock_Left - lngCost_Width)
9920          .frmAccountIncExpCodes_Sub.Width = (lngSub_Width - lngCost_Width)
9930        Case False
9940          .cmdClose.Left = (lngClose_Left - (lngCost_Width + lngCurrID_Width))
9950          .cmdUpdate.Left = (lngUpdate_Left - (lngCost_Width + lngCurrID_Width))
9960          .cmdBlockAssign.Left = (lngBlock_Left - (lngCost_Width + lngCurrID_Width))
9970          .frmAccountIncExpCodes_Sub.Width = (lngSub_Width - (lngCost_Width + lngCurrID_Width))
9980        End Select
9990        .frmAccountIncExpCodes_Sub_box.Width = (.frmAccountIncExpCodes_Sub.Width + (2& * lngTpp))
10000       .Nav_box01.Width = .frmAccountIncExpCodes_Sub.Width
10010       .Header_vline01.Left = lngNewFormWidth
10020       .Header_vline02.Left = lngNewFormWidth
10030       .Footer_vline01.Left = lngNewFormWidth
10040       .Footer_vline02.Left = lngNewFormWidth
10050       .Header_hline01.Width = lngNewFormWidth
10060       .Header_hline02.Width = lngNewFormWidth
10070       .Footer_hline01.Width = lngNewFormWidth
10080       .Footer_hline02.Width = lngNewFormWidth
10090       .Width = lngNewFormWidth
10100       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
10110       DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
10120       If lngMonitorNum > 1& Then
10130         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
10140       End If
10150     End Select
10160     DoEvents

10170   End With

EXITP:
10180   Exit Sub

ERRH:
10190   Select Case ERR.Number
        Case Else
10200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10210   End Select
10220   Resume EXITP

End Sub

Private Sub IncludeCurrency(blnShow As Boolean)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngNewFrmWidth As Long, lngNewFrmLeft As Long, lngNewFormWidth As Long
        Dim blnCost As Boolean
        Dim lngTmp01 As Long

10310   With Me

10320     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
10330       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
10340     End If
10350     blnCost = .chkShowCost

10360     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
10370     lngMonitorNum = 1&: lngTmp01 = 0&
10380     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
10390     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

10400     Select Case blnShow
          Case True
10410       Select Case blnCost
            Case True
10420         lngNewFrmWidth = lngFrm_Width
10430         lngNewFrmLeft = lngFrm_Left
10440         lngNewFormWidth = lngForm_Width
10450       Case False
10460         lngNewFrmWidth = (lngFrm_Width - lngCost_Width)
10470         lngNewFrmLeft = (lngFrm_Left + (lngCost_Width / 2))
10480         lngNewFormWidth = (lngForm_Width - lngCost_Width)
10490       End Select
10500       .Width = lngNewFormWidth
10510       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
10520       DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
10530       If lngMonitorNum > 1& Then
10540         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
10550       End If
10560       Select Case blnCost
            Case True
10570         .cmdClose.Left = lngClose_Left
10580         .cmdUpdate.Left = lngUpdate_Left
10590         .cmdBlockAssign.Left = lngBlock_Left
10600         .frmAccountIncExpCodes_Sub.Width = lngSub_Width
10610       Case False
10620         .cmdClose.Left = (lngClose_Left - lngCost_Width)
10630         .cmdUpdate.Left = (lngUpdate_Left - lngCost_Width)
10640         .cmdBlockAssign.Left = (lngBlock_Left - lngCost_Width)
10650         .frmAccountIncExpCodes_Sub.Width = (lngSub_Width - lngCost_Width)
10660       End Select
10670       .frmAccountIncExpCodes_Sub_box.Width = (.frmAccountIncExpCodes_Sub.Width + (2& * lngTpp))
10680       .Nav_box01.Width = .frmAccountIncExpCodes_Sub.Width
10690       .Header_vline01.Left = lngNewFormWidth
10700       .Header_vline02.Left = lngNewFormWidth
10710       .Footer_vline01.Left = lngNewFormWidth
10720       .Footer_vline02.Left = lngNewFormWidth
10730       .Header_hline01.Width = lngNewFormWidth
10740       .Header_hline02.Width = lngNewFormWidth
10750       .Footer_hline01.Width = lngNewFormWidth
10760       .Footer_hline02.Width = lngNewFormWidth
10770       .Width = lngNewFormWidth
10780     Case False
10790       Select Case blnCost
            Case True
10800         lngNewFrmWidth = (lngFrm_Width - lngCurrID_Width)
10810         lngNewFrmLeft = (lngFrm_Left + (lngCurrID_Width / 2))
10820         lngNewFormWidth = (lngForm_Width - lngCurrID_Width)
10830       Case False
10840         lngNewFrmWidth = (lngFrm_Width - (lngCost_Width + lngCurrID_Width))
10850         lngNewFrmLeft = (lngFrm_Left + ((lngCost_Width + lngCurrID_Width) / 2))
10860         lngNewFormWidth = (lngForm_Width - (lngCost_Width + lngCurrID_Width))
10870       End Select
10880       Select Case blnCost
            Case True
10890         .cmdClose.Left = (lngClose_Left - lngCurrID_Width)
10900         .cmdUpdate.Left = (lngUpdate_Left - lngCurrID_Width)
10910         .cmdBlockAssign.Left = (lngBlock_Left - lngCurrID_Width)
10920         .frmAccountIncExpCodes_Sub.Width = (lngSub_Width - lngCurrID_Width)
10930       Case False
10940         .cmdClose.Left = (lngClose_Left - (lngCost_Width + lngCurrID_Width))
10950         .cmdUpdate.Left = (lngUpdate_Left - (lngCost_Width + lngCurrID_Width))
10960         .cmdBlockAssign.Left = (lngBlock_Left - (lngCost_Width + lngCurrID_Width))
10970         .frmAccountIncExpCodes_Sub.Width = (lngSub_Width - (lngCost_Width + lngCurrID_Width))
10980       End Select
10990       .frmAccountIncExpCodes_Sub_box.Width = (.frmAccountIncExpCodes_Sub.Width + (2& * lngTpp))
11000       .Nav_box01.Width = .frmAccountIncExpCodes_Sub.Width
11010       .Header_vline01.Left = lngNewFormWidth
11020       .Header_vline02.Left = lngNewFormWidth
11030       .Footer_vline01.Left = lngNewFormWidth
11040       .Footer_vline02.Left = lngNewFormWidth
11050       .Header_hline01.Width = lngNewFormWidth
11060       .Header_hline02.Width = lngNewFormWidth
11070       .Footer_hline01.Width = lngNewFormWidth
11080       .Footer_hline02.Width = lngNewFormWidth
11090       .Width = lngNewFormWidth
11100       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
11110       DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
11120       If lngMonitorNum > 1& Then
11130         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
11140       End If
11150     End Select
11160     DoEvents

11170   End With

EXITP:
11180   Exit Sub

ERRH:
11190   Select Case ERR.Number
        Case Else
11200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11210   End Select
11220   Resume EXITP

End Sub
