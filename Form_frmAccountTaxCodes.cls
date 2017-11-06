VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountTaxCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountTaxCodes"

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
' **   Shorten:          S {cmdShorten on frmAccountTaxCodes_Sub}
' **   Update:           U {cmdUpdate}
' **   Widen:            W {cmdWiden on frmAccountTaxCodes_Sub}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAccountTaxCodes_Sub}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngCost_Width As Long, lngCurrID_Width As Long, lngSub_Width As Long
Private lngClose_Left As Long, lngUpdate_Left As Long, lngBlock_Left As Long

Private lngOutOfSync As Long
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

190           blnAskedToAbandon = False
200           gblnHasForExThis = False

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

350           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

360           .accountno = gstrAccountNo
370           .shortname = gstrAccountName

380           SetArchiveOption  ' ** Procedure: Below.
390           DoEvents

400           Set frm = .frmAccountTaxCodes_Sub.Form

410           .form_width_line.Left = 0&  ' ** Just get it out of the way.
420           lngForm_Width = .Width
430           lngCurrID_Width = ((frm.curr_id.Left + frm.curr_id.Width) - (frm.Cost.Left + frm.Cost.Width))
440           lngCost_Width = ((frm.Cost.Left + frm.Cost.Width) - (frm.PCash.Left + frm.PCash.Width))
450           lngClose_Left = .cmdClose.Left
460           lngUpdate_Left = .cmdUpdate.Left
470           lngBlock_Left = .cmdBlockAssign.Left
480           lngSub_Width = .frmAccountTaxCodes_Sub.Width

490           chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
500           DoEvents
510           chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
520           DoEvents
530           chkShowCost_AfterUpdate  ' ** Procedure: Below.
540           DoEvents

550           .frmAccountTaxCodes_Sub.SetFocus
560           frm.FocusHolder.SetFocus
570           .cmdUpdate.Enabled = False
580           DoEvents

              ' ** Check user status, and lock all fields if ViewOnly.
590           blnViewOnly = False
600           For Each grp In DBEngine.Workspaces(0).Groups
610             If grp.Name = "ViewOnly" Then
620               For Each usr In grp.Users
630                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
640                   blnViewOnly = True
650                   Exit For
660                 End If
670               Next
680             End If
690           Next
700           DoEvents

710           If blnViewOnly = True Then
720             .frmAccountTaxCodes_Sub.Form.taxcode.Locked = True
730           End If

740           DoCmd.Hourglass False

750         Else
760           Cancel = -1
770         End If
780       Else
790         Cancel = -1
800       End If

810     End With

820     If Cancel = -1 Then
830       Beep
840       DoCmd.Hourglass False
850       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
860     End If

EXITP:
870     Set frm = Nothing
880     Set usr = Nothing
890     Set grp = Nothing
900     Exit Sub

ERRH:
910     DoCmd.Hourglass False
920     Select Case ERR.Number
        Case Else
930       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
940     End Select
950     Resume EXITP

End Sub

Private Sub Form_Load()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1010    With Me

1020      CLR_DISABLED_FG = CLR_DKGRY
1030      CLR_DISABLED_BG = CLR_LTTEAL

1040      .accountno.ForeColor = CLR_DISABLED_FG
1050      .accountno.BackColor = CLR_DISABLED_BG
1060      .shortname.ForeColor = CLR_DISABLED_FG
1070      .shortname.BackColor = CLR_DISABLED_BG
1080      DoEvents

1090    End With

EXITP:
1100    Exit Sub

ERRH:
1110    Select Case ERR.Number
        Case Else
1120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1130    End Select
1140    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAccountTaxCodes_Sub}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyEscape
1270        intRetVal = 0
1280        cmdClose_Click  ' ** Procedure: Below.
1290      End Select
1300    End If

        ' ** Alt keys.
1310    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1320      Select Case intRetVal
          Case vbKeyX
1330        intRetVal = 0
1340        cmdClose_Click  ' ** Procedure: Below.
1350      End Select
1360    End If

        ' ** Ctrl keys.
1370    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1380      Select Case intRetVal
          Case vbKeyS
1390        With Me
1400          intRetVal = 0
1410          lngRecsCur = .frmAccountTaxCodes_Sub.Form.RecCnt  ' ** Form Function: frmAccountTaxCodes_Sub.
1420          If lngRecsCur > 0& Then
1430            .frmAccountTaxCodes_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountTaxCodes_Sub.
1440          Else
1450            Beep
1460          End If
1470        End With
1480      End Select
1490    End If

        ' ** Ctrl-Shift keys.
1500    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1510      Select Case intRetVal
          Case vbKeyF
1520        With Me
1530          intRetVal = 0
1540          .FocusHolder.SetFocus
1550        End With
1560      End Select
1570    End If

EXITP:
1580    KeyCode = intRetVal
1590    Exit Sub

ERRH:
1600    intRetVal = 0
1610    Select Case ERR.Number
        Case Else
1620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1630    End Select
1640    Resume EXITP

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
1820        DoCmd.Hourglass False
1830      End If
1840      If blnIsOpen = True Then
1850        blnIsOpen = False
1860        gblnHasForExThis = HasForEx_Acct(gstrAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.
1870        Select Case gblnHasForEx
            Case True
1880          Select Case gblnHasForExThis
              Case True
1890            If Pref_HasPref(THIS_NAME, "chkIncludeCurrency") = False Then  ' ** Module Procedure: modPreferenceFuncs.
1900              .chkIncludeCurrency = True
1910              chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1920            End If
1930            .chkIncludeCurrency.Enabled = True
1940          Case False
1950            .chkIncludeCurrency.Enabled = True
1960          End Select
1970        Case False
1980          .chkIncludeCurrency = False
1990          chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
2000          .chkIncludeCurrency.Enabled = False
2010        End Select
2020      End If  ' ** blnIsOpen.
2030      If gstrReturningForm <> vbNullString Then
2040        Select Case gstrReturningForm
            Case "frmAccountTaxCodes_BlockAssign"
2050          .cmdBlockAssign.SetFocus
2060        Case Else
              ' ** Nothing else.
2070        End Select
2080        gstrReturningForm = vbNullString
2090        DoCmd.Hourglass False
2100      End If
2110    End With

EXITP:
2120    Exit Sub

ERRH:
2130    DoCmd.Hourglass False
2140    Select Case ERR.Number
        Case Else
2150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2160    End Select
2170    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean
        Dim blnContinue As Boolean

2210    With Me

2220      DoCmd.Hourglass True
2230      DoEvents

2240      blnContinue = True

2250      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2260      If .chkUpdated = True And blnAskedToAbandon = False Then
2270        Beep
2280        DoCmd.Hourglass False
2290        If MsgBox("You have made changes, but have not confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbYes Then
              ' ** Abandon changes.
2300        Else
2310          blnContinue = False
2320          Cancel = -1
2330        End If
2340      End If

2350      If blnContinue = True Then

2360        DoCmd.Hourglass True
2370        DoEvents

2380        blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

            ' ** Make sure all reports are closed.
2390        If Reports.Count > 0 Then
2400          Do While Reports.Count > 0
2410            DoCmd.Close acReport, Reports(0).Name
2420          Loop
2430        End If

2440        gblnSetFocus = True
2450        gstrReturningForm = THIS_NAME
2460        strDocName = "frmMenu_Account"
2470        If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
2480          DoCmd.SelectObject acForm, strDocName, False
2490          Forms(strDocName).TimerInterval = 250&
2500        Else
2510          DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
2520        End If

2530      End If
2540    End With

EXITP:
2550    Exit Sub

ERRH:
2560    DoCmd.Hourglass False
2570    Select Case ERR.Number
        Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnContinue As Boolean

2710    blnContinue = True

2720    With Me

2730      .TimerInterval = 0&

2740      If .chkUpdated = True And blnViewOnly = False Then
2750        blnAskedToAbandon = False
2760        Beep
2770        If MsgBox("You have made changes, but have not confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbYes Then
2780          blnAskedToAbandon = True
2790        Else
2800          blnContinue = False
2810        End If
2820      ElseIf blnViewOnly = True Then
2830        blnAskedToAbandon = True
2840      End If

2850      If blnContinue = True Then

2860        Set dbs = CurrentDb
            ' ** Empty tblTaxCode_BlockAssign_Staging.
2870        Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_Block_06")
2880        qdf.Execute
2890        DoEvents

2900        DoCmd.Close acForm, THIS_NAME

2910      End If

2920    End With

EXITP:
2930    Exit Sub

ERRH:
2940    Set qdf = Nothing
2950    Set dbs = Nothing
2960    Select Case ERR.Number
        Case Else
2970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2980    End Select
2990    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
3070        With Me
3080          intRetVal = 0
3090          .frmAccountTaxCodes_Sub.SetFocus
3100          .frmAccountTaxCodes_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccountTaxCodes_Sub.
3110  On Error Resume Next
3120          .frmAccountTaxCodes_Sub.Form.taxcode_DESC.SetFocus
3130          If ERR.Number <> 0 Then
3140  On Error GoTo ERRH
3150            .frmAccountTaxCodes_Sub.Form.taxcode.SetFocus
3160          Else
3170  On Error GoTo ERRH
3180          End If
3190        End With
3200      End Select
3210    End If

        ' ** Shift keys.
3220    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3230      Select Case intRetVal
          Case vbKeyTab
3240        With Me
3250          intRetVal = 0
3260          If .cmdUpdate.Enabled = True Then
3270            .cmdUpdate.SetFocus
3280          Else
3290            .cmdBlockAssign.SetFocus
3300          End If
3310        End With
3320      End Select
3330    End If

EXITP:
3340    KeyCode = intRetVal
3350    Exit Sub

ERRH:
3360    intRetVal = 0
3370    Select Case ERR.Number
        Case Else
3380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3390    End Select
3400    Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

        Dim strFilter As String

3510    With Me
3520      If blnThisArchive = True Then
3530        Select Case .chkIncludeArchive
            Case True
3540          .chkIncludeArchive_lbl.FontBold = True
3550          .archive_trans_lbl.Visible = True
3560          .archive_trans_lbl_box.Visible = True
3570          strFilter = vbNullString
3580          .frmAccountTaxCodes_Sub.Form.Filter = strFilter
3590          .frmAccountTaxCodes_Sub.Form.FilterOn = False
3600        Case False
3610          .chkIncludeArchive_lbl.FontBold = False
3620          .archive_trans_lbl.Visible = False
3630          .archive_trans_lbl_box.Visible = False
3640          strFilter = "[src] = 'L'"
3650          .frmAccountTaxCodes_Sub.Form.Filter = strFilter
3660          .frmAccountTaxCodes_Sub.Form.FilterOn = True
3670        End Select
3680        .frmAccountTaxCodes_Sub.Form.Requery
3690        .frmAccountTaxCodes_Sub.SetFocus
3700        .frmAccountTaxCodes_Sub.Form.FocusHolder.SetFocus
3710      End If
3720    End With

EXITP:
3730    Exit Sub

ERRH:
3740    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
3750    Case Else
3760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3770    End Select
3780    Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3820    intShiftDown = (Shift And acShiftMask) > 0
3830    intAltDown = (Shift And acAltMask) > 0
3840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3870        With Me
3880          intRetVal = 0
3890          If .chkIncludeCurrency.Enabled = True Then
3900            .chkIncludeCurrency.SetFocus
3910          Else
3920            .chkShowCost.SetFocus
3930          End If
3940        End With
3950      End Select
3960    End If

        ' ** Shift keys.
3970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3980      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3990        With Me
4000          intRetVal = 0
4010          .frmAccountTaxCodes_Sub.SetFocus
4020          .frmAccountTaxCodes_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountTaxCodes_Sub.
4030  On Error Resume Next
4040          .frmAccountTaxCodes_Sub.Form.revcode_DESC.SetFocus
4050          If ERR.Number <> 0 Then
4060  On Error GoTo ERRH
4070            .frmAccountTaxCodes_Sub.Form.revcode_ID.SetFocus
4080          Else
4090  On Error GoTo ERRH
4100          End If
4110        End With
4120      End Select
4130    End If

EXITP:
4140    KeyCode = intRetVal
4150    Exit Sub

ERRH:
4160    intRetVal = 0
4170    Select Case ERR.Number
        Case Else
4180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4190    End Select
4200    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

4310    With Me
4320      Select Case .chkIncludeCurrency
          Case True
4330        .chkIncludeCurrency_lbl.FontBold = True
4340        .frmAccountTaxCodes_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountTaxCodes_Sub.
4350        IncludeCurrency True  ' ** Procedure: Below.
4360      Case False
4370        .chkIncludeCurrency_lbl.FontBold = False
4380        .frmAccountTaxCodes_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountTaxCodes_Sub.
4390        IncludeCurrency False  ' ** Procedure: Below.
4400      End Select
4410      .frmAccountTaxCodes_Sub.SetFocus
4420    End With

EXITP:
4430    Exit Sub

ERRH:
4440    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
4450    Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

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
4590          .chkShowCost.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4650        With Me
4660          intRetVal = 0
4670          If .chkIncludeArchive.Enabled = True Then
4680            .chkIncludeArchive.SetFocus
4690          Else
4700            .frmAccountTaxCodes_Sub.SetFocus
4710            .frmAccountTaxCodes_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountTaxCodes_Sub.
4720  On Error Resume Next
4730            .frmAccountTaxCodes_Sub.Form.revcode_DESC.SetFocus
4740            If ERR.Number <> 0 Then
4750  On Error GoTo ERRH
4760              .frmAccountTaxCodes_Sub.Form.revcode_ID.SetFocus
4770            Else
4780  On Error GoTo ERRH
4790            End If
4800          End If
4810        End With
4820      End Select
4830    End If

EXITP:
4840    KeyCode = intRetVal
4850    Exit Sub

ERRH:
4860    intRetVal = 0
4870    Select Case ERR.Number
        Case Else
4880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4890    End Select
4900    Resume EXITP

End Sub

Public Sub chkShowCost_AfterUpdate()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowCost_AfterUpdate"

5010    With Me
5020      Select Case .chkShowCost
          Case True
5030        .chkShowCost_lbl.FontBold = True
5040        .frmAccountTaxCodes_Sub.Form.ShowCost_Sub True  ' ** Form Procedure: frmAccountTaxCodes_Sub.
5050        ShowCost True  ' ** Procedure: Below.
5060      Case False
5070        .chkShowCost_lbl.FontBold = False
5080        .frmAccountTaxCodes_Sub.Form.ShowCost_Sub False  ' ** Form Procedure: frmAccountTaxCodes_Sub.
5090        ShowCost False  ' ** Procedure: Below.
5100      End Select
5110      .frmAccountTaxCodes_Sub.SetFocus
5120    End With

EXITP:
5130    Exit Sub

ERRH:
5140    Select Case ERR.Number
        Case Else
5150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5160    End Select
5170    Resume EXITP

End Sub

Private Sub chkShowCost_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowCost_KeyDown"

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
5290          .cmdBlockAssign.SetFocus
5300        End With
5310      End Select
5320    End If

        ' ** Shift keys.
5330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5350        With Me
5360          intRetVal = 0
5370          If .chkIncludeCurrency.Enabled = True Then
5380            .chkIncludeCurrency.SetFocus
5390          ElseIf .chkIncludeArchive.Enabled = True Then
5400            .chkIncludeArchive.SetFocus
5410          Else
5420            .frmAccountTaxCodes_Sub.SetFocus
5430            .frmAccountTaxCodes_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountTaxCodes_Sub.
5440  On Error Resume Next
5450            .frmAccountTaxCodes_Sub.Form.taxcode_DESC.SetFocus
5460            If ERR.Number <> 0 Then
5470  On Error GoTo ERRH
5480              .frmAccountTaxCodes_Sub.Form.taxcode.SetFocus
5490            Else
5500  On Error GoTo ERRH
5510            End If
5520          End If
5530        End With
5540      End Select
5550    End If

EXITP:
5560    KeyCode = intRetVal
5570    Exit Sub

ERRH:
5580    intRetVal = 0
5590    Select Case ERR.Number
        Case Else
5600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5610    End Select
5620    Resume EXITP

End Sub

Private Sub cmdBlockAssign_Click()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBlockAssign_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim blnArchive As Boolean

5710    With Me

5720      DoCmd.Hourglass True
5730      DoEvents

5740      Set dbs = CurrentDb

5750      blnArchive = False

          ' ** Empty tblTaxCode_BlockAssign_Staging.
5760      Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_Block_06")
5770      qdf.Execute
5780      Set qdf = Nothing
5790      DoEvents

5800      blnArchive = .chkIncludeArchive

5810      If blnArchive = True Then
            ' ** Append qryAccountProfile_TaxCodes_Block_03_03 (LedgerArchive,
            ' ** as tblTaxCode_BlockAssign_Staging records, by specified [actno])
            ' ** to tblTaxCode_BlockAssign_Staging.
5820        Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_Block_03_04")
5830        With qdf.Parameters
5840          ![actno] = gstrAccountNo
5850        End With
5860        qdf.Execute
5870        Set qdf = Nothing
5880      End If
5890      DoEvents

          ' ** Append qryAccountProfile_TaxCodes_Block_03_01 (Ledger,
          ' ** as tblTaxCode_BlockAssign_Staging records, by specified [actno])
          ' ** to tblTaxCode_BlockAssign_Staging.
5900      Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_Block_03_02")
5910      With qdf.Parameters
5920        ![actno] = gstrAccountNo
5930      End With
5940      qdf.Execute
5950      Set qdf = Nothing
5960      DoEvents

5970      dbs.Close
5980      Set dbs = Nothing

5990      strDocName = "frmAccountTaxCodes_BlockAssign"
6000      gblnSetFocus = True
6010      DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & gstrAccountNo & "~" & Format(blnArchive, "True/False")

6020    End With

EXITP:
6030    Set qdf = Nothing
6040    Set dbs = Nothing
6050    Exit Sub

ERRH:
6060    DoCmd.Hourglass False
6070    Select Case ERR.Number
        Case Else
6080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6090    End Select
6100    Resume EXITP

End Sub

Private Sub cmdBlockAssign_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBlockAssign_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6220    intShiftDown = (Shift And acShiftMask) > 0
6230    intAltDown = (Shift And acAltMask) > 0
6240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6260      Select Case intRetVal
          Case vbKeyTab
6270        With Me
6280          intRetVal = 0
6290          If .cmdUpdate.Enabled = True Then
6300            .cmdUpdate.SetFocus
6310          Else
6320            .cmdClose.SetFocus
6330          End If
6340        End With
6350      End Select
6360    End If

        ' ** Shift keys.
6370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6380      Select Case intRetVal
          Case vbKeyTab
6390        With Me
6400          intRetVal = 0
6410          .chkShowCost.SetFocus
6420        End With
6430      End Select
6440    End If

EXITP:
6450    KeyCode = intRetVal
6460    Exit Sub

ERRH:
6470    intRetVal = 0
6480    Select Case ERR.Number
        Case Else
6490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6500    End Select
6510    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, qdf3 As DAO.QueryDef
        Dim rstTmpEdit As DAO.Recordset, rstLedger As DAO.Recordset, rstArchive As DAO.Recordset
        Dim strSource As String
        Dim lngRecs1 As Long, lngRecs2 As Long, lngRecs3 As Long
        Dim blnInTrans As Boolean, blnContinue As Boolean
        Dim lngX As Long

6610    With Me
6620      If blnViewOnly = False Then

6630        blnInTrans = False: blnContinue = True

6640        .frmAccountTaxCodes_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountTaxCodes_Sub.

6650        Set wrk = DBEngine.Workspaces(0)
6660        Set dbs = wrk.Databases(0)

            ' ** Set out-of-sync transactions as changed.
            ' ** If any of these have already been fixed,
            ' ** they'll no longer show up in this query.
6670        If lngOutOfSync > 0& And gblnLinkRevTaxCodes = True Then
              ' ** Update qryAccountProfile_TaxCodes_11 (tmpEdit06, for OutOfSync = 1,
              ' ** with taxcode_new, revcode_ID_new, revcode_TYPE_new, revcode_DESC_new).
6680          Set qdf1 = dbs.QueryDefs("qryAccountProfile_TaxCodes_13")
6690          qdf1.Execute
6700          Set qdf1 = Nothing
6710        End If

            ' ** Table tmpEdit06, all fields, changed = True.
6720        Set qdf1 = dbs.QueryDefs("qryAccountProfile_TaxCodes_02")
6730        Set rstTmpEdit = qdf1.OpenRecordset
6740        If rstTmpEdit.BOF = True And rstTmpEdit.EOF = True Then
              ' ** The button shouldn't have been available if there were no changes!
6750          Beep
6760          MsgBox "No changes were found.", vbInformation + vbOKOnly, "Nothing To Do"
6770        Else
6780          rstTmpEdit.MoveLast
6790          lngRecs1 = rstTmpEdit.RecordCount
6800          rstTmpEdit.MoveFirst
              ' ** Ledger, linked to tmpEdit06, by specified [actno].
6810          Set qdf2 = dbs.QueryDefs("qryAccountProfile_TaxCodes_03")
6820          With qdf2.Parameters
6830            ![actno] = gstrAccountNo
6840          End With
6850          Set rstLedger = qdf2.OpenRecordset
              ' ** LedgerArchive, linked to tmpEdit6, by specified [actno].
6860          Set qdf3 = dbs.QueryDefs("qryAccountProfile_TaxCodes_03a")
6870          With qdf3.Parameters
6880            ![actno] = gstrAccountNo
6890          End With
6900          Set rstArchive = qdf3.OpenRecordset
6910          If rstLedger.BOF = True And rstLedger.EOF = True Then
                ' ** I don't think it could get here.
6920            If rstArchive.BOF = True And rstArchive.EOF = True Then
6930              blnContinue = False
6940            End If
6950          End If
6960          If blnContinue = True Then
6970            rstLedger.MoveLast
6980            lngRecs2 = rstLedger.RecordCount
6990            rstLedger.MoveFirst
7000            If rstArchive.BOF = True And rstArchive.EOF = True Then
7010              lngRecs3 = 0&
7020            Else
7030              rstArchive.MoveLast
7040              lngRecs3 = rstArchive.RecordCount
7050              rstArchive.MoveFirst
7060            End If
7070            If lngRecs1 > 0& And (lngRecs2 > 0& Or lngRecs3 > 0&) Then
7080              wrk.BeginTrans
7090              blnInTrans = True
7100              rstTmpEdit.MoveFirst
7110              For lngX = 1& To lngRecs1
7120                blnContinue = True
7130                strSource = rstTmpEdit![src]
7140                Select Case strSource
                    Case "A"
7150                  rstArchive.FindFirst "journalno = " & CStr(rstTmpEdit![journalno])
7160                  If rstArchive.NoMatch = True Then
7170                    blnContinue = False
7180                    MsgBox "Transaction not found in Ledger Archive.", vbCritical + vbOKOnly, "Record Not Found"
7190                  End If
7200                Case "L"
7210                  rstLedger.FindFirst "journalno = " & str(rstTmpEdit![journalno])
7220                  If rstLedger.NoMatch = True Then
7230                    blnContinue = False
7240                    MsgBox "Transaction not found in ledger.", vbCritical + vbOKOnly, "Record Not Found"
7250                  End If
7260                End Select
7270                If blnContinue = True Then
7280                  If IsNull(rstTmpEdit![taxcode]) = False Then
7290                    Select Case strSource
                        Case "A"
7300                      If IsNull(rstArchive![taxcode]) = False Then
7310                        If rstArchive![taxcode] <> rstTmpEdit![taxcode] Then
7320                          rstArchive.Edit
7330                          rstArchive![taxcode] = rstTmpEdit![taxcode]
7340                          rstArchive.Update
7350                        End If
7360                      Else
7370                        rstArchive.Edit
7380                        rstArchive![taxcode] = Nz(rstTmpEdit![taxcode], 0)
7390                        rstArchive.Update
7400                      End If
7410                    Case "L"
7420                      If IsNull(rstLedger![taxcode]) = False Then
7430                        If rstLedger![taxcode] <> rstTmpEdit![taxcode] Then
7440                          rstLedger.Edit
7450                          rstLedger![taxcode] = rstTmpEdit![taxcode]
7460                          rstLedger.Update
7470                        End If
7480                      Else
7490                        rstLedger.Edit
7500                        rstLedger![taxcode] = Nz(rstTmpEdit![taxcode], 0)
7510                        rstLedger.Update
7520                      End If
7530                    End Select
7540                    If gblnLinkRevTaxCodes = True Then
7550                      Select Case strSource
                          Case "A"
7560                        If IsNull(rstArchive![revcode_ID]) = False Then
7570                          If rstArchive![revcode_ID] <> rstTmpEdit![revcode_ID] Then
7580                            rstArchive.Edit
7590                            rstArchive![revcode_ID] = rstTmpEdit![revcode_ID]
7600                            rstArchive.Update
7610                          End If
7620                        Else
7630                          rstArchive.Edit
7640                          rstArchive![revcode_ID] = Nz(rstTmpEdit![revcode_ID], 1&)
7650                          rstArchive.Update
7660                        End If
7670                      Case "L"
7680                        If IsNull(rstLedger![revcode_ID]) = False Then
7690                          If rstLedger![revcode_ID] <> rstTmpEdit![revcode_ID] Then
7700                            rstLedger.Edit
7710                            rstLedger![revcode_ID] = rstTmpEdit![revcode_ID]
7720                            rstLedger.Update
7730                          End If
7740                        Else
7750                          rstLedger.Edit
7760                          rstLedger![revcode_ID] = Nz(rstTmpEdit![revcode_ID], 1&)
7770                          rstLedger.Update
7780                        End If
7790                      End Select
7800                    End If
7810                  Else
                        ' ** If this is Null, don't make any change to the Ledger, though
                        ' ** I don't think they could ever succeed in Nulling out taxcode.
7820                  End If
7830                End If  ' ** blnContinue.
7840                If lngX < lngRecs1 Then rstTmpEdit.MoveNext
7850              Next
7860              wrk.CommitTrans
7870              blnInTrans = False
7880              .chkUpdated = False
7890              .cmdClose.SetFocus
7900              .cmdUpdate.Enabled = False
7910              MsgBox "Updates completed.", vbInformation + vbOKOnly, ("Update Successful" & Space(40))
7920            End If
7930          End If  ' ** blnContinue.
7940          rstLedger.Close
7950          rstArchive.Close
7960        End If
7970        rstTmpEdit.Close
7980        dbs.Close
7990        wrk.Close

8000      End If  ' ** blnViewOnly.
8010    End With

EXITP:
8020    Set rstTmpEdit = Nothing
8030    Set rstLedger = Nothing
8040    Set rstArchive = Nothing
8050    Set qdf1 = Nothing
8060    Set qdf2 = Nothing
8070    Set qdf3 = Nothing
8080    Set dbs = Nothing
8090    Set wrk = Nothing
8100    Exit Sub

ERRH:
8110    If blnInTrans = True Then
8120      wrk.Rollback
8130      blnInTrans = False
8140      rstTmpEdit.Close
8150      rstLedger.Close
8160      dbs.Close
8170      wrk.Close
8180    End If
8190    Select Case ERR.Number
        Case Else
8200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8210    End Select
8220    Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8320    intShiftDown = (Shift And acShiftMask) > 0
8330    intAltDown = (Shift And acAltMask) > 0
8340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8360      Select Case intRetVal
          Case vbKeyTab
8370        With Me
8380          intRetVal = 0
8390          .cmdClose.SetFocus
8400        End With
8410      End Select
8420    End If

        ' ** Shift keys.
8430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8440      Select Case intRetVal
          Case vbKeyTab
8450        With Me
8460          intRetVal = 0
8470          .cmdBlockAssign.SetFocus
8480        End With
8490      End Select
8500    End If

EXITP:
8510    KeyCode = intRetVal
8520    Exit Sub

ERRH:
8530    intRetVal = 0
8540    Select Case ERR.Number
        Case Else
8550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8560    End Select
8570    Resume EXITP

End Sub

Private Sub SetArchiveOption()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "SetArchiveOption"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

8610    With Me

8620      blnHasArchive = False: blnThisArchive = False

8630      Set dbs = CurrentDb
8640      With dbs

8650        Set rst = .OpenRecordset("LedgerArchive", dbOpenDynaset, dbReadOnly)
8660        With rst
8670          If .BOF = True And .EOF = True Then
                ' ** No archive whatsoever!
8680          Else
8690            .MoveLast
8700            If .RecordCount > 1 Then
8710              blnHasArchive = True
8720            End If
8730          End If
8740          .Close
8750        End With
8760        Set rst = Nothing
8770        DoEvents

8780        If blnHasArchive = True Then
              ' ** LedgerArchive, with 'Liability (+)'/'Liability (-)' broken out,
              ' ** with totdescx, by specified GlobalVarGet('gstrAccountNo').
8790          Set qdf = .QueryDefs("qryAccountProfile_TaxCodes_04")
8800          Set rst = qdf.OpenRecordset
8810          With rst
8820            If .BOF = True And .EOF = True Then
                  ' ** This account has no archive.
8830            Else
8840              .MoveLast
8850              If .RecordCount > 0 Then
8860                blnThisArchive = True
8870              End If
8880            End If
8890            .Close
8900          End With
8910          Set rst = Nothing
8920          Set qdf = Nothing
8930        End If  ' ** blnHasArchive.
8940        DoEvents

8950        .Close
8960      End With
8970      Set dbs = Nothing

8980      .archive_trans_lbl.Visible = False
8990      .archive_trans_lbl_box.Visible = False

9000      Select Case blnHasArchive
          Case True
9010        Select Case blnThisArchive
            Case True
9020          .chkIncludeArchive.Enabled = True
9030          .chkIncludeArchive_lbl.Visible = True
9040          .chkIncludeArchive_lbl4_dim.Visible = False
9050          .chkIncludeArchive_lbl4_dim_hi.Visible = False
9060          .chkIncludeArchive_lbl5_dim.Visible = False
9070          .chkIncludeArchive_lbl5_dim_hi.Visible = False
9080        Case False
9090          .chkIncludeArchive.Enabled = False
9100          .chkIncludeArchive_lbl.Visible = False
9110          .chkIncludeArchive_lbl4_dim.Visible = True
9120          .chkIncludeArchive_lbl4_dim_hi.Visible = True
9130          .chkIncludeArchive_lbl5_dim.Visible = True
9140          .chkIncludeArchive_lbl5_dim_hi.Visible = True
9150        End Select
9160        .chkIncludeArchive_lbl2_dim.Visible = False
9170        .chkIncludeArchive_lbl2_dim_hi.Visible = False
9180        .chkIncludeArchive_lbl3_dim.Visible = False
9190        .chkIncludeArchive_lbl3_dim_hi.Visible = False
9200      Case False
9210        .chkIncludeArchive.Enabled = False
9220        .chkIncludeArchive_lbl.Visible = False
9230        .chkIncludeArchive_lbl4_dim.Visible = False
9240        .chkIncludeArchive_lbl4_dim_hi.Visible = False
9250        .chkIncludeArchive_lbl5_dim.Visible = False
9260        .chkIncludeArchive_lbl5_dim_hi.Visible = False
9270        .chkIncludeArchive_lbl2_dim.Visible = True
9280        .chkIncludeArchive_lbl2_dim_hi.Visible = True
9290        .chkIncludeArchive_lbl3_dim.Visible = True
9300        .chkIncludeArchive_lbl3_dim_hi.Visible = True
9310      End Select
9320      DoEvents

9330    End With

EXITP:
9340    Set rst = Nothing
9350    Set qdf = Nothing
9360    Set dbs = Nothing
9370    Exit Sub

ERRH:
9380    Select Case ERR.Number
        Case Else
9390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9400    End Select
9410    Resume EXITP

End Sub

Private Sub ShowCost(blnShow As Boolean)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowCost"

        Dim lngNewFrmWidth As Long, lngNewFrmLeft As Long, lngNewFormWidth As Long
        Dim blnCurrID As Boolean
        Dim lngTmp01 As Long

9510    With Me

9520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
9530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9540      End If
9550      blnCurrID = .chkIncludeCurrency

9560      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
9570      lngMonitorNum = 1&: lngTmp01 = 0&
9580      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
9590      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

9600      Select Case blnShow
          Case True
9610        Select Case blnCurrID
            Case True
9620          lngNewFrmWidth = lngFrm_Width
9630          lngNewFrmLeft = lngFrm_Left
9640          lngNewFormWidth = lngForm_Width
9650        Case False
9660          lngNewFrmWidth = (lngFrm_Width - lngCurrID_Width)
9670          lngNewFrmLeft = (lngFrm_Left + (lngCurrID_Width / 2))
9680          lngNewFormWidth = (lngForm_Width - lngCurrID_Width)
9690        End Select
9700        .Width = lngNewFormWidth
9710        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
9720        DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
9730        If lngMonitorNum > 1& Then
9740          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
9750        End If
9760        Select Case blnCurrID
            Case True
9770          .cmdClose.Left = lngClose_Left
9780          .cmdUpdate.Left = lngUpdate_Left
9790          .cmdBlockAssign.Left = lngBlock_Left
9800          .frmAccountTaxCodes_Sub.Width = lngSub_Width
9810        Case False
9820          .cmdClose.Left = (lngClose_Left - lngCurrID_Width)
9830          .cmdUpdate.Left = (lngUpdate_Left - lngCurrID_Width)
9840          .cmdBlockAssign.Left = (lngBlock_Left - lngCurrID_Width)
9850          .frmAccountTaxCodes_Sub.Width = (lngSub_Width - lngCurrID_Width)
9860        End Select
9870        .frmAccountTaxCodes_Sub_box.Width = (.frmAccountTaxCodes_Sub.Width + (2& * lngTpp))
9880        .Nav_box01.Width = .frmAccountTaxCodes_Sub.Width
9890        .Header_vline01.Left = lngNewFormWidth
9900        .Header_vline02.Left = lngNewFormWidth
9910        .Detail_vline01.Left = lngNewFormWidth
9920        .Detail_vline02.Left = lngNewFormWidth
9930        .Footer_vline01.Left = lngNewFormWidth
9940        .Footer_vline02.Left = lngNewFormWidth
9950        .Header_hline01.Width = lngNewFormWidth
9960        .Header_hline02.Width = lngNewFormWidth
9970        .Detail_hline01.Width = lngNewFormWidth
9980        .Detail_hline02.Width = lngNewFormWidth
9990        .Footer_hline01.Width = lngNewFormWidth
10000       .Footer_hline02.Width = lngNewFormWidth
10010     Case False
10020       Select Case blnCurrID
            Case True
10030         lngNewFrmWidth = (lngFrm_Width - lngCost_Width)
10040         lngNewFrmLeft = (lngFrm_Left + (lngCost_Width / 2))
10050         lngNewFormWidth = (lngForm_Width - lngCost_Width)
10060       Case False
10070         lngNewFrmWidth = (lngFrm_Width - (lngCost_Width + lngCurrID_Width))
10080         lngNewFrmLeft = (lngFrm_Left + ((lngCost_Width + lngCurrID_Width) / 2))
10090         lngNewFormWidth = (lngForm_Width - (lngCost_Width + lngCurrID_Width))
10100       End Select
10110       Select Case blnCurrID
            Case True
10120         .cmdClose.Left = (lngClose_Left - lngCost_Width)
10130         .cmdUpdate.Left = (lngUpdate_Left - lngCost_Width)
10140         .cmdBlockAssign.Left = (lngBlock_Left - lngCost_Width)
10150         .frmAccountTaxCodes_Sub.Width = (lngSub_Width - lngCost_Width)
10160       Case False
10170         .cmdClose.Left = (lngClose_Left - (lngCost_Width + lngCurrID_Width))
10180         .cmdUpdate.Left = (lngUpdate_Left - (lngCost_Width + lngCurrID_Width))
10190         .cmdBlockAssign.Left = (lngBlock_Left - (lngCost_Width + lngCurrID_Width))
10200         .frmAccountTaxCodes_Sub.Width = (lngSub_Width - (lngCost_Width + lngCurrID_Width))
10210       End Select
10220       .frmAccountTaxCodes_Sub_box.Width = (.frmAccountTaxCodes_Sub.Width + (2& * lngTpp))
10230       .Nav_box01.Width = .frmAccountTaxCodes_Sub.Width
10240       .Header_vline01.Left = lngNewFormWidth
10250       .Header_vline02.Left = lngNewFormWidth
10260       .Detail_vline01.Left = lngNewFormWidth
10270       .Detail_vline02.Left = lngNewFormWidth
10280       .Footer_vline01.Left = lngNewFormWidth
10290       .Footer_vline02.Left = lngNewFormWidth
10300       .Header_hline01.Width = lngNewFormWidth
10310       .Header_hline02.Width = lngNewFormWidth
10320       .Detail_hline01.Width = lngNewFormWidth
10330       .Detail_hline02.Width = lngNewFormWidth
10340       .Footer_hline01.Width = lngNewFormWidth
10350       .Footer_hline02.Width = lngNewFormWidth
10360       .Width = lngNewFormWidth
10370       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
10380       DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
10390       If lngMonitorNum > 1& Then
10400         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
10410       End If
10420     End Select
10430     DoEvents

10440   End With

EXITP:
10450   Exit Sub

ERRH:
10460   Select Case ERR.Number
        Case Else
10470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10480   End Select
10490   Resume EXITP

End Sub

Private Sub IncludeCurrency(blnShow As Boolean)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngNewFrmWidth As Long, lngNewFrmLeft As Long, lngNewFormWidth As Long
        Dim blnCost As Boolean
        Dim lngTmp01 As Long

10510   With Me

10520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
10530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
10540     End If
10550     blnCost = .chkShowCost

10560     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
10570     lngMonitorNum = 1&: lngTmp01 = 0&
10580     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
10590     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

10600     Select Case blnShow
          Case True
10610       Select Case blnCost
            Case True
10620         lngNewFrmWidth = lngFrm_Width
10630         lngNewFrmLeft = lngFrm_Left
10640         lngNewFormWidth = lngForm_Width
10650       Case False
10660         lngNewFrmWidth = (lngFrm_Width - lngCost_Width)
10670         lngNewFrmLeft = (lngFrm_Left + (lngCost_Width / 2))
10680         lngNewFormWidth = (lngForm_Width - lngCost_Width)
10690       End Select
10700       .Width = lngNewFormWidth
10710       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
10720       DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
10730       If lngMonitorNum > 1& Then
10740         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
10750       End If
10760       Select Case blnCost
            Case True
10770         .cmdClose.Left = lngClose_Left
10780         .cmdUpdate.Left = lngUpdate_Left
10790         .cmdBlockAssign.Left = lngBlock_Left
10800         .frmAccountTaxCodes_Sub.Width = lngSub_Width
10810       Case False
10820         .cmdClose.Left = (lngClose_Left - lngCost_Width)
10830         .cmdUpdate.Left = (lngUpdate_Left - lngCost_Width)
10840         .cmdBlockAssign.Left = (lngBlock_Left - lngCost_Width)
10850         .frmAccountTaxCodes_Sub.Width = (lngSub_Width - lngCost_Width)
10860       End Select
10870       .frmAccountTaxCodes_Sub_box.Width = (.frmAccountTaxCodes_Sub.Width + (2& * lngTpp))
10880       .Nav_box01.Width = .frmAccountTaxCodes_Sub.Width
10890       .Header_vline01.Left = lngNewFormWidth
10900       .Header_vline02.Left = lngNewFormWidth
10910       .Detail_vline01.Left = lngNewFormWidth
10920       .Detail_vline02.Left = lngNewFormWidth
10930       .Footer_vline01.Left = lngNewFormWidth
10940       .Footer_vline02.Left = lngNewFormWidth
10950       .Header_hline01.Width = lngNewFormWidth
10960       .Header_hline02.Width = lngNewFormWidth
10970       .Detail_hline01.Width = lngNewFormWidth
10980       .Detail_hline02.Width = lngNewFormWidth
10990       .Footer_hline01.Width = lngNewFormWidth
11000       .Footer_hline02.Width = lngNewFormWidth
11010     Case False
11020       Select Case blnCost
            Case True
11030         lngNewFrmWidth = (lngFrm_Width - lngCurrID_Width)
11040         lngNewFrmLeft = (lngFrm_Left + (lngCurrID_Width / 2))
11050         lngNewFormWidth = (lngForm_Width - lngCurrID_Width)
11060       Case False
11070         lngNewFrmWidth = (lngFrm_Width - (lngCost_Width + lngCurrID_Width))
11080         lngNewFrmLeft = (lngFrm_Left + ((lngCost_Width + lngCurrID_Width) / 2))
11090         lngNewFormWidth = (lngForm_Width - (lngCost_Width + lngCurrID_Width))
11100       End Select
11110       Select Case blnCost
            Case True
11120         .cmdClose.Left = (lngClose_Left - lngCurrID_Width)
11130         .cmdUpdate.Left = (lngUpdate_Left - lngCurrID_Width)
11140         .cmdBlockAssign.Left = (lngBlock_Left - lngCurrID_Width)
11150         .frmAccountTaxCodes_Sub.Width = (lngSub_Width - lngCurrID_Width)
11160       Case False
11170         .cmdClose.Left = (lngClose_Left - (lngCost_Width + lngCurrID_Width))
11180         .cmdUpdate.Left = (lngUpdate_Left - (lngCost_Width + lngCurrID_Width))
11190         .cmdBlockAssign.Left = (lngBlock_Left - (lngCost_Width + lngCurrID_Width))
11200         .frmAccountTaxCodes_Sub.Width = (lngSub_Width - (lngCost_Width + lngCurrID_Width))
11210       End Select
11220       .frmAccountTaxCodes_Sub_box.Width = (.frmAccountTaxCodes_Sub.Width + (2& * lngTpp))
11230       .Nav_box01.Width = .frmAccountTaxCodes_Sub.Width
11240       .Header_vline01.Left = lngNewFormWidth
11250       .Header_vline02.Left = lngNewFormWidth
11260       .Detail_vline01.Left = lngNewFormWidth
11270       .Detail_vline02.Left = lngNewFormWidth
11280       .Footer_vline01.Left = lngNewFormWidth
11290       .Footer_vline02.Left = lngNewFormWidth
11300       .Header_hline01.Width = lngNewFormWidth
11310       .Header_hline02.Width = lngNewFormWidth
11320       .Detail_hline01.Width = lngNewFormWidth
11330       .Detail_hline02.Width = lngNewFormWidth
11340       .Footer_hline01.Width = lngNewFormWidth
11350       .Footer_hline02.Width = lngNewFormWidth
11360       .Width = lngNewFormWidth
11370       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
11380       DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
11390       If lngMonitorNum > 1& Then
11400         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
11410       End If
11420     End Select
11430     DoEvents

11440   End With

EXITP:
11450   Exit Sub

ERRH:
11460   Select Case ERR.Number
        Case Else
11470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11480   End Select
11490   Resume EXITP

End Sub
