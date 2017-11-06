VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssets"

'VGC 10/02/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   IncludeCurrency()
' #######################################

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   Close:            C {cmdClose}
' **   Edit:             E {cmdEdit}
' **   Show Hidden:      H {chkShowHidden}
' **   CUSIP:            I {cusip on frmAssets_Sub}
' **   Include Currency: L {chkIncludeCurrency}  (cmdUpdate has to be Alt+U)
' **   Show AssetNo:     N {chkShowAssetNo}
' **   Only Assets Held: O {opgShow_optHeldOnly}
' **   Print:            P {cmdPrintReport}
' **   Shorten:          S {cmdShorten on frmAssets_Sub}
' **   All Assets:       T {opgShow_optAll}
' **   Update:           U {cmdUpdate}
' **   Widen:            W {cmdWiden on frmAssets_Sub}
' **   Exit:             X {cmdClose}
' **   Asset Type:       Y {cmbAssetType on frmAssets_Sub}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:          F5 {RecalcNewAssets, RecalcForEx}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAssets_Sub}

' ** MasterAsset table:
' **   Masterasset_TYPE field:
' **     RA = Regular Asset
' **     IA = Interest Asset
' **       CUSIP: 999999999
' **       This asset comes with the program, there can be only one, and it
' **       can't be deleted, though it can be renamed to whatever you want.
' **       Used with frmAccruedIncome.
' **   Description field prefixes:
' **     'HA-' = Hidden Asset
' **       Doesn't show up in drop-downs, and can only be hidden if 0 shares.
' **       Can have as many as you like.
' **     'SW-' = Sweep Asset
' **       Used with account_SWEEP field in Account table, and frmSweeper.
' **       Only 'SW-' assets show up in drop-down.
' **       Can have as many as you like.

' ** RecChanged = 0 : Unchanged.
' **              1 : Changed.
' **              2 : New asset.
' **              3 : Marked for deletion (currently not allowed).

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const HID_FILTER As String = "Left([description_masterasset],3)<>'HA-'"
Private Const HELD_FILTER As String = "[IsHeld] = True"

' ** Array: arr_varIA_Asset().
Private arr_varIA_Asset As Variant
Private Const IA_MTYPE As Integer = 0
Private Const IA_ID    As Integer = 1
Private Const IA_CUSIP As Integer = 2
Private Const IA_DESC  As Integer = 3
Private Const IA_ATYPE As Integer = 5
Private Const IA_RATE  As Integer = 6
Private Const IA_DUE   As Integer = 7
Private blnIAHasChanged As Boolean

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngSub_Width As Long, lngCurrID_Width As Long, lngPrintReportLbl_Left As Long
Private lngClose_Left As Long, lngUpdate_Left As Long, lngAdd_Left As Long, lngEdit_Left As Long

Private blnHasForEx_Chk As Boolean, blnRecalcTots As Boolean
Private blnGoingToReport As Boolean, blnJustClose As Boolean, blnFromLoad As Boolean
Private lngAssetNo As Long, lngIsHidSum As Long, lngIsHid_Width As Long
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim frm As Access.Form
        Dim intPos01 As Integer
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)
160           intPos01 = InStr(strCallingForm, "~")
170           If intPos01 > 0 Then
                ' ** Came from frmJournal_Columns.
180             lngAssetNo = Val(Mid(strCallingForm, (intPos01 + 1)))
190             strCallingForm = Left(strCallingForm, (intPos01 - 1))
200           Else
210             lngAssetNo = 0&
220           End If

230           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

240           If strCallingForm <> "frmJournal_Columns" Then
250             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
260           End If
270           DoEvents

              ' ** Just get these out of the way.
280           .form_width_line.Left = 0&
290           .chkUpdated.Left = 0&
300           .chkReloading.Left = 0&
310           .NewAssets.Visible = False

320           blnJustClose = False: blnRecalcTots = False
330           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
340           gblnHasForExThis = False: blnHasForEx_Chk = True  ' ** Yes, True.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
350           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Variables are fed empty, then populated ByRef.
360           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.
370           DoEvents

380           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

390           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
400           If lngMonitorCnt > 1& Then
410             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
420             lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
430             If lngMonitorNum > 1& Then
440               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
450             End If
460           End If

470           Set frm = .frmAssets_Sub.Form

480           lngForm_Width = .Width
490           lngClose_Left = .cmdClose.Left
500           lngUpdate_Left = .cmdUpdate.Left
510           lngAdd_Left = .cmdAdd.Left
520           lngEdit_Left = .cmdEdit.Left
530           lngSub_Width = .frmAssets_Sub.Width
540           lngCurrID_Width = ((frm.curr_id.Left + frm.curr_id.Width) - (frm.description_masterasset.Left + frm.description_masterasset.Width))

550           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
560           DoEvents

570           CLR_DISABLED_FG = CLR_DKGRY
580           CLR_DISABLED_BG = CLR_LTTEAL

              ' ** Load arr_varIA_Asset().
590           blnIAHasChanged = False
600           LoadIAData  ' ** Function: Below.
610           DoEvents

620           RecalcForEx  ' ** Procedure: Below.

630           blnIsOpen = False

640         Else
650           Cancel = -1
660         End If
670       Else
680         Cancel = -1
690       End If
700     End With

710     If Cancel = -1 Then
720       DoCmd.Hourglass False
730       Beep
740       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
750     End If

EXITP:
760     Set frm = Nothing
770     Exit Sub

ERRH:
780     DoCmd.Hourglass False
790     THAT_PROC = THIS_PROC
800     That_Erl = Erl
810     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
820     Resume EXITP

End Sub

Private Sub Form_Load()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form

910     With Me

920       Set frm = .frmAssets_Sub.Form
930       frm.FocusHolder.SetFocus
940       If frm.cusip.Visible = True Then  'WHAT DOES THIS MEAN? THE CUSIP IS ALWAYS VISIBLE!
950         frm.cusip.Enabled = False       'AAHH. WHEN DELETE WAS ALLOWED, ASSETS MARKED FOR DELETION DISABLED THE FIELDS.
960         frm.cusip.Locked = True
970         frm.description_masterasset.Enabled = False
980         frm.description_masterasset.Locked = True
990         frm.curr_id.Enabled = False
1000        frm.curr_id.Locked = True
1010        frm.rate_entry.Enabled = False
1020        frm.rate_entry.Locked = True
1030        frm.due.Enabled = False
1040        frm.due.Locked = True
1050        frm.cmbAssetType.Enabled = False
1060        frm.cmbAssetType.Locked = True
1070        .cmdAdd.Enabled = True
1080        .cmdEdit.Enabled = True
1090        .cmdUpdate.Enabled = False
1100        .cmdEdit.Tag = "Locked"
1110      Else
1120        frm.cusip.Enabled = False
1130        frm.cusip.Locked = True
1140        frm.description_masterasset.Enabled = False
1150        frm.description_masterasset.Locked = True
1160        frm.curr_id.Enabled = False
1170        frm.curr_id.Locked = True
1180        frm.rate_entry.Enabled = False
1190        frm.rate_entry.Locked = True
1200        frm.due.Enabled = False
1210        frm.due.Locked = True
1220        frm.cmbAssetType.Enabled = False
1230        frm.cmbAssetType.Locked = True
1240        .cmdAdd.Enabled = False
1250      End If
1260      DoEvents
          ' ** Controls with 'disabled' BackColor:
1270      frm.cusip.ForeColor = CLR_DISABLED_FG
1280      frm.cusip.BackColor = CLR_DISABLED_BG
1290      frm.description_masterasset.ForeColor = CLR_DISABLED_FG
1300      frm.description_masterasset.BackColor = CLR_DISABLED_BG
1310      frm.curr_id.ForeColor = CLR_DISABLED_FG
1320      frm.curr_id.BackColor = CLR_DISABLED_BG
1330      frm.curr_id_usd_box.ForeColor = CLR_DISABLED_BG
1340      frm.curr_id_forex_box.ForeColor = CLR_VLTGRN
1350      frm.rate_entry.ForeColor = CLR_DISABLED_FG
1360      frm.rate_entry.BackColor = CLR_DISABLED_BG
1370      frm.due.ForeColor = CLR_DISABLED_FG
1380      frm.due.BackColor = CLR_DISABLED_BG
1390      frm.shareface.ForeColor = CLR_DISABLED_FG
1400      frm.shareface.BackColor = CLR_DISABLED_BG
1410      frm.cmbAssetType.ForeColor = CLR_DISABLED_FG
1420      frm.cmbAssetType.BackColor = CLR_DISABLED_BG
1430      frm.Filter = HID_FILTER
1440      frm.FilterOn = True
1450      DoEvents

1460      Set dbs = CurrentDb
1470      With dbs
            ' ** tmpEdit09, summed by IsHid.
1480        Set qdf = .QueryDefs("qryAssets_07")
1490        Set rst = qdf.OpenRecordset
1500        With rst
1510          If .BOF = True And .EOF = True Then
1520            lngIsHidSum = 0&
1530          Else
1540            .MoveFirst
1550            lngIsHidSum = ![IsHidSum]
1560          End If
1570          .Close
1580        End With
1590        .Close
1600      End With
1610      DoEvents

1620      .IsHidSum = lngIsHidSum
1630      lngIsHid_Width = .chkShowHidden_lbl.Width
1640      If lngIsHidSum = 0& Then
1650        .chkShowHidden.Enabled = False
1660        .chkShowHidden = False
1670        .chkShowHidden_lbl.FontBold = False
1680        .chkShowHidden_lbl.Visible = False
1690        .chkShowHidden_lbl2_dim.Visible = True
1700        .chkShowHidden_lbl2_dim_hi.Visible = True
1710        .chkShowHidden_lbl3_dim.Visible = True
1720        .chkShowHidden_lbl3_dim_hi.Visible = True
1730      End If
1740      .frmAssets_Sub.Form.IsHid_Width_Set lngIsHid_Width  ' ** Form Procedure: frmAssets_Sub.
1750      DoEvents

          ' ** Pass some values to the subform.
1760      frm.IAHasChanged_Set blnIAHasChanged  ' ** Form Procedure: frmAssets_Sub.
1770      frm.ParForm_Width_Set lngForm_Width  ' ** Form Procedure: frmAssets_Sub.
1780      DoEvents

1790      blnFromLoad = True
1800      opgShow_AfterUpdate  ' ** Procedure: Below.
1810      DoEvents
1820      chkShowHidden_AfterUpdate  ' ** Procedure: Below.
1830      DoEvents
1840      chkShowAssetNo_AfterUpdate  ' ** Procedure: Below.
1850      DoEvents
1860      chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1870      DoEvents
1880      chkOrderBy_AfterUpdate  ' ** Procedure: Below.
1890      DoEvents

1900      If lngAssetNo > 0& Then
1910        .TimerInterval = 500&
1920      End If

1930      DoCmd.Hourglass False
1940      DoEvents
1950      .cmdEdit.SetFocus

1960    End With

EXITP:
1970    Set frm = Nothing
1980    Set rst = Nothing
1990    Set qdf = Nothing
2000    Set dbs = Nothing
2010    Exit Sub

ERRH:
2020    DoCmd.Hourglass False
2030    THAT_PROC = THIS_PROC
2040    That_Erl = Erl
2050    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2060    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc New:      F5 {RecalcNewAssets, RecalcForEx}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   CUSIP:            I {cusip on frmAssets_Sub}
        ' **   Asset Type:       Y {cmbAssetType on frmAssets_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAssets_Sub}

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyF5
2170        intRetVal = 0
2180        RecalcNewAssets  ' ** Procedure: Below.
2190        RecalcForEx  ' ** Procedure: Below.
2200      Case vbKeyEscape
2210        intRetVal = 0
2220        cmdClose_Click  ' ** Procedure: Below.
2230      End Select
2240    End If

        ' ** Alt keys.
2250    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyI
2270        With Me
2280          intRetVal = 0
2290          .frmAssets_Sub.SetFocus
2300          If .frmAssets_Sub.Form.cusip.Enabled = True Then
2310            .frmAssets_Sub.Form.cusip.SetFocus
2320          End If
2330        End With
2340      Case vbKeyX
2350        intRetVal = 0
2360        cmdClose_Click  ' ** Procedure: Below.
2370      Case vbKeyY
2380        With Me
2390          intRetVal = 0
2400          .frmAssets_Sub.SetFocus
2410          If .frmAssets_Sub.Form.cmbAssetType.Enabled = True Then
2420            .frmAssets_Sub.Form.cmbAssetType.SetFocus
2430          End If
2440        End With
2450      End Select
2460    End If

        ' ** Ctrl keys.
2470    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2480      Select Case intRetVal
          Case vbKeyS
2490        With Me
2500          intRetVal = 0
2510          lngRecsCur = .frmAssets_Sub.Form.RecCnt  ' ** Form Function: frmAssets_Sub.
2520          If lngRecsCur > 0& Then
2530            .frmAssets_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAssets_Sub.
2540          Else
2550            Beep
2560          End If
2570        End With
2580      End Select
2590    End If

        ' ** Ctrl-Shift keys.
2600    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2610      Select Case intRetVal
          Case vbKeyF
2620        With Me
2630          intRetVal = 0
2640          .FocusHolder.SetFocus
2650        End With
2660      End Select
2670    End If

EXITP:
2680    KeyCode = intRetVal
2690    Exit Sub

ERRH:
2700    intRetVal = 0
2710    THAT_PROC = THIS_PROC
2720    That_Erl = Erl
2730    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2740    Resume EXITP

End Sub

Private Sub Form_Timer()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2810    With Me
2820      .TimerInterval = 0&
2830      If gblnSetFocus = True Then
2840        gblnSetFocus = False
2850        strFormName = vbNullString
2860  On Error Resume Next
2870        Do Until Screen.ActiveForm.Name = strFormName
2880          DoCmd.SelectObject acForm, THIS_NAME, False
2890          strFormName = Screen.ActiveForm.Name
2900        Loop
2910  On Error GoTo ERRH
2920      End If  ' ** gblnSetFocus.
2930      If gblnGoToReport = True Then
2940        Select Case blnGoingToReport
            Case True
2950          DoCmd.Hourglass False  ' ** It should already be off.
2960          gblnGoToReport = False
2970          blnGoingToReport = False
2980          .GoToReport_arw_pr_img.Visible = False
2990          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3000          DoEvents
3010        Case False
3020          DoCmd.Hourglass True  ' ** Make sure it's still running.
3030          DoEvents
3040          blnGoingToReport = True
3050          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3060          .GoToReport_arw_pr_img.Visible = True
3070          If .cmdPrintReport.Enabled = True Then
3080            .cmdPrintReport.SetFocus
3090          Else
3100            .cmdClose.SetFocus
3110          End If
3120          DoEvents
3130          Beep
3140          DoCmd.Hourglass False
3150          DoEvents
3160          .TimerInterval = (GTR_WAIT * 2&)
3170        End Select
3180      End If
3190      If lngAssetNo > 0& Then
3200        .frmAssets_Sub.Form.MoveRec 0, lngAssetNo  ' ** Form Procedure: frmAssets_Sub.
3210      End If
3220      If blnHasForEx_Chk = True Then
3230        blnHasForEx_Chk = False
3240        gblnHasForExThis = gblnHasForEx
3250        Select Case gblnHasForEx
            Case True
3260          Select Case gblnHasForExThis
              Case True
3270            If Pref_HasPref(THIS_NAME, "chkIncludeCurrency") = False Then  ' ** Module Procedure: modPreferenceFuncs.
3280              .chkIncludeCurrency = True
3290              chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
3300            End If
3310            .chkIncludeCurrency.Enabled = True
3320          Case False
3330            .chkIncludeCurrency.Enabled = True
3340          End Select
3350        Case False
3360          .chkIncludeCurrency = False
3370          chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
3380          .chkIncludeCurrency.Enabled = False
3390        End Select
3400      End If
3410      If blnRecalcTots = True Then
3420        blnRecalcTots = False
3430        RecalcNewAssets  ' ** Procedure: Below.
3440        RecalcForEx  ' ** Procedure: Below.
3450      End If
3460    End With

EXITP:
3470    Exit Sub

ERRH:
3480    Select Case ERR.Number
        Case Else
3490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3500    End Select
3510    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3610    If DataErr <> 0 Then
3620      Select Case DataErr
          Case 2046, 2108, 2116, 2169, 2237, 2501, 3020, 3101, 3314, 7753, 8519
            ' ** 2046: The command or action '|' isn't available now.
            ' ** 2108: You must save the field before you execute the GoToControl action, the GoToControl method, or the SetFocus method.
            ' ** 2116: The value in the field or record violates the validation rule for the record or field.
            ' ** 2169: You can't save this record at this time.
            ' ** 2237: The text you entered isn't an item in the list.
            ' ** 2501: The '|' action was Canceled.
            ' ** 3020: Update or CancelUpdate without AddNew or Edit.
            ' ** 3101: The Microsoft Jet database engine cannot find a record in the table '|2' with key matching field(s) '|1'.
            ' ** 3314: The field '|' cannot contain a Null value because the Required property for this field is set to True.
            ' ** 7753: The value you entered doesn't meet the validation rule defined for the field or control.
            ' ** 8519: You are about to delete | record(s).
3630        Response = acDataErrContinue
3640      Case 2105  ' ** You can't go to the specified record.
3650        Response = acDataErrContinue
3660      Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
3670        Response = acDataErrContinue
3680      Case 2427  ' ** The expression you entered has no value.
3690        Response = acDataErrContinue
3700      Case Else
3710        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3720        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3730      End Select
3740    End If

EXITP:
3750    Exit Sub

ERRH:
3760    Select Case ERR.Number
        Case Else
3770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3780    End Select
3790    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean

3810    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3820    If strCallingForm <> "frmJournal_Columns" Then
3830      blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel
3840    End If

3850    strDocName = "frmAssets_Add"
3860    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
3870      DoCmd.Close acForm, strDocName, acSaveNo
3880    End If

3890    strDocName = "rptListOfAssets"
3900    If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
3910      DoCmd.Close acReport, strDocName, acSaveNo
3920    End If

3930    If blnJustClose = False Then
3940      gstrReturningForm = THIS_NAME
3950      gblnSetFocus = True
3960      Select Case strCallingForm
          Case "frmMenu_Asset"
3970        If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
3980          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
3990        Else
4000          DoCmd.SelectObject acForm, strCallingForm, False
4010        End If
4020      Case "frmJournal_Columns"
4030        DoCmd.SelectObject acForm, strCallingForm, False
4040        Forms(strCallingForm).cmdLocNew.Enabled = True
4050        Forms(strCallingForm).cmdRecurNew.Enabled = True
4060      End Select
4070    End If

EXITP:
4080    Exit Sub

ERRH:
4090    Select Case ERR.Number
        Case Else
4100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select
4120    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4210    With Me
4220      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
4230        Select Case blnPrintReport_Focus
            Case True
4240          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
4250          .cmdPrintReport_raised_img.Visible = False
4260        Case False
4270          .cmdPrintReport_raised_img.Visible = True
4280          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4290        End Select
4300        .cmdPrintReport_raised_focus_img.Visible = False
4310        .cmdPrintReport_raised_focus_dots_img.Visible = False
4320        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4330        .cmdPrintReport_raised_img_dis.Visible = False
4340      End If
4350    End With

EXITP:
4360    Exit Sub

ERRH:
4370    Select Case ERR.Number
        Case Else
4380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4390    End Select
4400    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Public Sub cmdClose_Click()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

4610    blnContinue = True

4620    With Me

4630      If .chkUpdated = True Then
4640        If .NewAssets > 0 Then
4650          strMsg = "You have added " & CStr(.NewAssets) & " new asset"
4660          If .NewAssets > 1 Then strMsg = strMsg & "s"
4670          strMsg = strMsg & ", but not saved "
4680          If .NewAssets > 1 Then strMsg = strMsg & "them." Else strMsg = strMsg & "it."
4690          strMsg = strMsg & vbCrLf & vbCrLf
4700          strMsg = strMsg & "Are you sure you want to abandon changes?"
4710        Else
4720          strMsg = "You have made changes, but have not saved them." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?"
4730        End If
4740        msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved")
4750        If msgResponse <> vbYes Then
4760          blnContinue = False
4770        End If
4780      End If

4790      If blnContinue = True Then

4800        .frmAssets_Sub.SetFocus
4810        .frmAssets_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAssets_Sub.

4820        Set dbs = CurrentDb
4830        With dbs
              ' ** MasterAsset, grouped, for 'HA-' prefix, with cnt.
4840          Set qdf = .QueryDefs("qryAssets_14")
4850          Set rst = qdf.OpenRecordset
4860          With rst
4870            .MoveFirst  ' ** There will always be one record.
4880            If ![cnt] > 0 Then
4890              blnContinue = False
4900              MsgBox "You have marked an asset as hidden which is currently held in accounts (non-zero share/face)." & vbCrLf & vbCrLf & _
                    "Please unhide all assets with an amount in the share/face column to continue.", vbInformation + vbOKOnly, "Invalid Entry"
4910            End If
4920            .Close
4930          End With
4940          If blnContinue = True Then
                ' ** MasterAsset, grouped, for cusip, description = Null, NullString, with cnt.
4950            Set qdf = .QueryDefs("qryAssets_15")
4960            Set rst = qdf.OpenRecordset
4970            With rst
4980              .MoveFirst  ' ** There will always be one record.
4990              If ![cnt] > 0 Then
5000                blnContinue = False
5010                MsgBox "You have at least one asset with no description and/or no CUSIP." & vbCrLf & vbCrLf & _
                      "Please make sure every asset has a description and a CUSIP to continue.", vbInformation + vbOKOnly, "Invalid Entry"
5020              End If
5030              .Close
5040            End With
5050          End If
5060          .Close
5070        End With  ' ** dbs.

5080        If blnContinue = True Then  ' ** frmJournal_Columns handled in Form_Close().

5090          DoCmd.Close acForm, THIS_NAME

5100        End If

5110      End If

5120    End With

EXITP:
5130    Set rst = Nothing
5140    Set qdf = Nothing
5150    Set dbs = Nothing
5160    Exit Sub

ERRH:
5170    THAT_PROC = THIS_PROC
5180    That_Erl = Erl
5190    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5200    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab
5370        intRetVal = 0
5380        With Me
5390          lngRecsCur = .frmAssets_Sub.Form.RecCnt  ' ** Form Function: frmAssets_Sub.
5400          If lngRecsCur > 0& Then
5410            .frmAssets_Sub.SetFocus
5420            .frmAssets_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAssets_Sub.
5430            If .frmAssets_Sub.Form.cusip.Enabled = True Then
5440              .frmAssets_Sub.Form.cusip.SetFocus
5450            Else
5460              .frmAssets_Sub.Form.FocusHolder.SetFocus
5470            End If
5480          Else
5490            .opgShow.SetFocus
5500          End If
5510        End With
5520      End Select
5530    End If

        ' ** Shift keys.
5540    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5550      Select Case intRetVal
          Case vbKeyTab
5560        intRetVal = 0
5570        With Me
5580          If .cmdUpdate.Enabled = True Then
5590            .cmdUpdate.SetFocus
5600          ElseIf .cmdAdd.Enabled = True Then
5610            .cmdAdd.SetFocus
5620          ElseIf .cmdEdit.Enabled = True Then
5630            .cmdEdit.SetFocus
5640          ElseIf .cmdPrintReport.Enabled = True Then
5650            .cmdPrintReport.SetFocus
5660          ElseIf .chkShowHidden.Enabled = True Then
5670            .chkShowHidden.SetFocus
5680          ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
5690            .chkShowAssetNo.SetFocus
5700          ElseIf .chkIncludeCurrency.Enabled = True Then
5710            .chkIncludeCurrency.SetFocus
5720          Else
5730            .opgShow.SetFocus
5740          End If
5750        End With
5760      End Select
5770    End If

EXITP:
5780    KeyCode = intRetVal
5790    Exit Sub

ERRH:
5800    intRetVal = 0
5810    THAT_PROC = THIS_PROC
5820    That_Erl = Erl
5830    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5840    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

5910    With Me
5920      blnPrintReport_Focus = True
5930      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
5940      .cmdPrintReport_raised_img.Visible = False
5950      .cmdPrintReport_raised_focus_img.Visible = False
5960      .cmdPrintReport_raised_focus_dots_img.Visible = False
5970      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5980      .cmdPrintReport_raised_img_dis.Visible = False
5990    End With

EXITP:
6000    Exit Sub

ERRH:
6010    Select Case ERR.Number
        Case Else
6020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6030    End Select
6040    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

6110    With Me
6120      blnPrintReport_MouseDown = True
6130      .cmdPrintReport_sunken_focus_dots_img.Visible = True
6140      .cmdPrintReport_raised_img.Visible = False
6150      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6160      .cmdPrintReport_raised_focus_img.Visible = False
6170      .cmdPrintReport_raised_focus_dots_img.Visible = False
6180      .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_Click()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String, strSortNow As String
        Dim blnContinue As Boolean, blnDesc As Boolean

6310    blnContinue = True
6320    blnDesc = False

6330    With Me

6340      Set dbs = CurrentDb

6350      Select Case .chkShowHidden
          Case True
            ' ** MasterAsset table, with hidden assets.
6360        Set qdf = dbs.QueryDefs("qryRpt_ListOfAssets_02")
6370      Case False
            ' ** MasterAsset table, without hidden assets.
6380        Set qdf = dbs.QueryDefs("qryRpt_ListOfAssets_01")
6390      End Select
6400      Set rst = qdf.OpenRecordset
6410      With rst
6420        If .BOF = True And .EOF = True Then
6430          blnContinue = False
6440          MsgBox "There are no assets to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
6450        End If
6460        .Close
6470      End With
6480      dbs.Close

6490      If blnContinue = True Then
            ' ** The sort order of the report will reflect the sort order of the subform.
6500        strDocName = "rptListOfAssets"
6510        strSortNow = .frmAssets_Sub.Form.SortNow_Get  ' ** Form Function: frmAssets_Sub.
6520        strSortNow = Trim(Left(strSortNow, (InStr(strSortNow, ",") - 1)))  ' ** All sorts have 3 fields.
6530        If Right(strSortNow, 4) = "DESC" Then
6540          blnDesc = True
6550          strSortNow = Trim(Left(strSortNow, InStr(strSortNow, " ")))
6560        End If
6570        strSortNow = Rem_Brackets(strSortNow)  ' ** Module Function: modStringFuncs.
6580        Select Case strSortNow
            Case "cusip"
6590          strDocName = strDocName & "_01a"
6600        Case "description_masterasset_sort"
6610          strDocName = strDocName & "_02a"
6620        Case "rate"
6630          strDocName = strDocName & "_03a"
6640        Case "due"
6650          strDocName = strDocName & "_04a"
6660        Case "shareface"
6670          strDocName = strDocName & "_05a"
6680        Case "assettype"
6690          strDocName = strDocName & "_06a"
6700        End Select
6710        If blnDesc = True Then strDocName = Left(strDocName, (Len(strDocName) - 1)) & "d"
6720        If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
6730          DoCmd.OpenReport strDocName, acViewPreview
6740          DoCmd.Maximize
6750          DoCmd.RunCommand acCmdFitToWindow
6760          .Visible = False
6770        Else
              '##GTR_Ref: rptListOfAssets
6780          DoCmd.OpenReport strDocName, acViewNormal
6790        End If
6800      End If

6810    End With

EXITP:
6820    Set rst = Nothing
6830    Set qdf = Nothing
6840    Set dbs = Nothing
6850    Exit Sub

ERRH:
6860    DoCmd.Restore
6870    Me.Visible = True
6880    THAT_PROC = THIS_PROC
6890    That_Erl = Erl
6900    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6910    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7020    intShiftDown = (Shift And acShiftMask) > 0
7030    intAltDown = (Shift And acAltMask) > 0
7040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7060      Select Case intRetVal
          Case vbKeyTab
7070        With Me
7080          intRetVal = 0
7090          If .cmdEdit.Enabled = True Then
7100            .cmdEdit.SetFocus
7110          ElseIf .cmdAdd.Enabled = True Then
7120            .cmdAdd.SetFocus
7130          ElseIf .cmdUpdate.Enabled = True Then
7140            .cmdUpdate.SetFocus
7150          Else
7160            .cmdClose.SetFocus
7170          End If
7180        End With
7190      End Select
7200    End If

        ' ** Shift keys.
7210    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7220      Select Case intRetVal
          Case vbKeyTab
7230        With Me
7240          intRetVal = 0
7250          If .chkShowHidden.Enabled = True Then
7260            .chkShowHidden.SetFocus
7270          ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
7280            .chkShowAssetNo.SetFocus
7290          ElseIf .chkIncludeCurrency.Enabled = True Then
7300            .chkIncludeCurrency.SetFocus
7310          Else
7320            .opgShow.SetFocus
7330          End If
7340        End With
7350      End Select
7360    End If

EXITP:
7370    KeyCode = intRetVal
7380    Exit Sub

ERRH:
7390    intRetVal = 0
7400    THAT_PROC = THIS_PROC
7410    That_Erl = Erl
7420    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7430    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

7510    With Me
7520      If blnPrintReport_MouseDown = False Then
7530        Select Case blnPrintReport_Focus
            Case True
7540          .cmdPrintReport_raised_focus_dots_img.Visible = True
7550          .cmdPrintReport_raised_focus_img.Visible = False
7560        Case False
7570          .cmdPrintReport_raised_focus_img.Visible = True
7580          .cmdPrintReport_raised_focus_dots_img.Visible = False
7590        End Select
7600        .cmdPrintReport_raised_img.Visible = False
7610        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7620        .cmdPrintReport_sunken_focus_dots_img.Visible = False
7630        .cmdPrintReport_raised_img_dis.Visible = False
7640      End If
7650    End With

EXITP:
7660    Exit Sub

ERRH:
7670    Select Case ERR.Number
        Case Else
7680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7690    End Select
7700    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

7810    With Me
7820      .cmdPrintReport_raised_focus_dots_img.Visible = True
7830      .cmdPrintReport_raised_img.Visible = False
7840      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7850      .cmdPrintReport_raised_focus_img.Visible = False
7860      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7870      .cmdPrintReport_raised_img_dis.Visible = False
7880      blnPrintReport_MouseDown = False
7890    End With

EXITP:
7900    Exit Sub

ERRH:
7910    Select Case ERR.Number
        Case Else
7920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7930    End Select
7940    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

8010    With Me
8020      .cmdPrintReport_raised_img.Visible = True
8030      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8040      .cmdPrintReport_raised_focus_img.Visible = False
8050      .cmdPrintReport_raised_focus_dots_img.Visible = False
8060      .cmdPrintReport_sunken_focus_dots_img.Visible = False
8070      .cmdPrintReport_raised_img_dis.Visible = False
8080      blnPrintReport_Focus = False
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

Private Sub opgShow_AfterUpdate()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim strTmp01 As String

8210    With Me
8220      Set frm = .frmAssets_Sub.Form
8230      Select Case .opgShow
          Case .opgShow_optAll.OptionValue
8240        .opgShow_optAll_lbl.FontBold = True
8250        .opgShow_optHeldOnly_lbl.FontBold = False
8260        Select Case .chkShowHidden
            Case True
8270          strFilter = vbNullString
8280          frm.Filter = strFilter
8290          frm.FilterOn = False
8300        Case False
8310          strFilter = HID_FILTER
8320          frm.Filter = strFilter
8330          frm.FilterOn = True
8340        End Select
8350      Case .opgShow_optHeldOnly.OptionValue
8360        .opgShow_optAll_lbl.FontBold = False
8370        .opgShow_optHeldOnly_lbl.FontBold = True
8380        Select Case .chkShowHidden
            Case True
8390          strFilter = HELD_FILTER
8400          frm.Filter = strFilter
8410          frm.FilterOn = True
8420        Case False
8430          strFilter = HID_FILTER & " And " & HELD_FILTER
8440          frm.Filter = strFilter
8450          frm.FilterOn = True
8460        End Select
8470      End Select
8480      frm.Refresh
8490      strTmp01 = frm.SortNow_Get  ' ** Form Function: frmAssets_Sub.
8500      frm.OrderBy = strTmp01
8510      frm.OrderByOn = True
8520      If blnFromLoad = False Then
8530        lngRecsCur = .frmAssets_Sub.Form.RecCnt  ' ** Form Function: frmAssets_Sub.
8540        If lngRecsCur > 0& Then
8550          .frmAssets_Sub.SetFocus
8560        End If
8570      Else
8580        blnFromLoad = False
8590      End If
8600    End With

EXITP:
8610    Set frm = Nothing
8620    Exit Sub

ERRH:
8630    THAT_PROC = THIS_PROC
8640    That_Erl = Erl
8650    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
8660    Resume EXITP

End Sub

Private Sub opgShow_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8720    intShiftDown = (Shift And acShiftMask) > 0
8730    intAltDown = (Shift And acAltMask) > 0
8740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8770        With Me
8780          intRetVal = 0
8790          If .chkIncludeCurrency.Enabled = True Then
8800            .chkIncludeCurrency.SetFocus
8810          ElseIf .chkShowAssetNo.Visible And .chkShowAssetNo.Enabled = True Then
8820            .chkShowAssetNo.SetFocus
8830          ElseIf .chkShowHidden.Enabled = True Then
8840            .chkShowHidden.SetFocus
8850          ElseIf .cmdPrintReport.Enabled = True Then
8860            .cmdPrintReport.SetFocus
8870          ElseIf .cmdEdit.Enabled = True Then
8880            .cmdEdit.SetFocus
8890          ElseIf .cmdAdd.Enabled = True Then
8900            .cmdAdd.SetFocus
8910          ElseIf .cmdUpdate.Enabled = True Then
8920            .cmdUpdate.SetFocus
8930          Else
8940            .cmdClose.SetFocus
8950          End If
8960        End With
8970      End Select
8980    End If

        ' ** Shift keys.
8990    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9000      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9010        With Me
9020          intRetVal = 0
9030          lngRecsCur = .frmAssets_Sub.Form.RecCnt  ' ** Form Function: frmAssets_Sub.
9040          If lngRecsCur > 0& Then
9050            .frmAssets_Sub.SetFocus
9060          Else
9070            .cmdClose.SetFocus
9080          End If
9090        End With
9100      End Select
9110    End If

EXITP:
9120    KeyCode = intRetVal
9130    Exit Sub

ERRH:
9140    intRetVal = 0
9150    THAT_PROC = THIS_PROC
9160    That_Erl = Erl
9170    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9180    Resume EXITP

End Sub

Private Sub opgShow_optHeldOnly_KeyDown(KeyCode As Integer, Shift As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_optHeldOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9220    intShiftDown = (Shift And acShiftMask) > 0
9230    intAltDown = (Shift And acAltMask) > 0
9240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9270        With Me
9280          intRetVal = 0
9290          If .chkIncludeCurrency.Enabled = True Then
9300            .chkIncludeCurrency.SetFocus
9310          ElseIf .chkShowAssetNo.Visible And .chkShowAssetNo.Enabled = True Then
9320            .chkShowAssetNo.SetFocus
9330          ElseIf .chkShowHidden.Enabled = True Then
9340            .chkShowHidden.SetFocus
9350          ElseIf .cmdPrintReport.Enabled = True Then
9360            .cmdPrintReport.SetFocus
9370          ElseIf .cmdEdit.Enabled = True Then
9380            .cmdEdit.SetFocus
9390          ElseIf .cmdAdd.Enabled = True Then
9400            .cmdAdd.SetFocus
9410          ElseIf .cmdUpdate.Enabled = True Then
9420            .cmdUpdate.SetFocus
9430          Else
9440            .cmdClose.SetFocus
9450          End If
9460        End With
9470      End Select
9480    End If

        ' ** Shift keys.
9490    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9500      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9510        With Me
9520          intRetVal = 0
9530          lngRecsCur = .frmAssets_Sub.Form.RecCnt  ' ** Form Function: frmAssets_Sub.
9540          If lngRecsCur > 0& Then
9550            .frmAssets_Sub.SetFocus
9560          Else
9570            .cmdClose.SetFocus
9580          End If
9590        End With
9600      End Select
9610    End If

EXITP:
9620    KeyCode = intRetVal
9630    Exit Sub

ERRH:
9640    intRetVal = 0
9650    THAT_PROC = THIS_PROC
9660    That_Erl = Erl
9670    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9680    Resume EXITP

End Sub

Public Sub chkIncludeCurrency_AfterUpdate()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

9710    With Me
9720      Select Case .chkIncludeCurrency
          Case True
9730        .chkIncludeCurrency_lbl.FontBold = True
9740        .frmAssets_Sub.Form.IncludeCurrency True  ' ** Form Procedure: frmAssets_Sub.
9750        DoEvents
9760        IncludeCurrency True  ' ** Procedure: Below.
9770      Case False
9780        .chkIncludeCurrency_lbl.FontBold = False
9790        .frmAssets_Sub.Form.IncludeCurrency False  ' ** Form Procedure: frmAssets_Sub.
9800        DoEvents
9810        IncludeCurrency False  ' ** Procedure: Below.
9820      End Select
9830      .frmAssets_Sub.SetFocus
9840    End With

EXITP:
9850    Exit Sub

ERRH:
9860    THAT_PROC = THIS_PROC
9870    That_Erl = Erl
9880    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
9890    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

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
9990          If .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
10000           .chkShowAssetNo.SetFocus
10010         ElseIf .chkShowHidden.Enabled = True Then
10020           .chkShowHidden.SetFocus
10030         ElseIf .cmdPrintReport.Enabled = True Then
10040           .cmdPrintReport.SetFocus
10050         ElseIf .cmdEdit.Enabled = True Then
10060           .cmdEdit.SetFocus
10070         ElseIf .cmdAdd.Enabled = True Then
10080           .cmdAdd.SetFocus
10090         ElseIf .cmdUpdate.Enabled = True Then
10100           .cmdUpdate.SetFocus
10110         Else
10120           .cmdClose.SetFocus
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
10210         .opgShow.SetFocus
10220       End With
10230     End Select
10240   End If

EXITP:
10250   KeyCode = intRetVal
10260   Exit Sub

ERRH:
10270   intRetVal = 0
10280   THAT_PROC = THIS_PROC
10290   That_Erl = Erl
10300   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10310   Resume EXITP

End Sub

Private Sub chkShowAssetNo_AfterUpdate()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetNo_AfterUpdate"

10410   With Me
10420     Select Case .chkShowAssetNo
          Case True
10430       .chkShowAssetNo_lbl.FontBold = True
10440       .frmAssets_Sub.Form.ShowAssetNo True  ' ** Form Procedure: frmAssets_Sub.
10450     Case False
10460       .chkShowAssetNo_lbl.FontBold = False
10470       .frmAssets_Sub.Form.ShowAssetNo False  ' ** Form Procedure: frmAssets_Sub.
10480     End Select
10490     .frmAssets_Sub.SetFocus
10500     .frmAssets_Sub.Form.FocusHolder.SetFocus
10510   End With

EXITP:
10520   Exit Sub

ERRH:
10530   THAT_PROC = THIS_PROC
10540   That_Erl = Erl
10550   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
10560   Resume EXITP

End Sub

Private Sub chkShowAssetNo_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10620   intShiftDown = (Shift And acShiftMask) > 0
10630   intAltDown = (Shift And acAltMask) > 0
10640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10670       With Me
10680         intRetVal = 0
10690         If .chkShowHidden.Enabled = True Then
10700           .chkShowHidden.SetFocus
10710         ElseIf .cmdPrintReport.Enabled = True Then
10720           .cmdPrintReport.SetFocus
10730         ElseIf .cmdEdit.Enabled = True Then
10740           .cmdEdit.SetFocus
10750         ElseIf .cmdAdd.Enabled = True Then
10760           .cmdAdd.SetFocus
10770         ElseIf .cmdUpdate.Enabled = True Then
10780           .cmdUpdate.SetFocus
10790         Else
10800           .cmdClose.SetFocus
10810         End If
10820       End With
10830     End Select
10840   End If

        ' ** Shift keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
10890         If .chkIncludeCurrency.Enabled = True Then
10900           .chkIncludeCurrency.SetFocus
10910         Else
10920           .opgShow.SetFocus
10930         End If
10940       End With
10950     End Select
10960   End If

EXITP:
10970   KeyCode = intRetVal
10980   Exit Sub

ERRH:
10990   intRetVal = 0
11000   THAT_PROC = THIS_PROC
11010   That_Erl = Erl
11020   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11030   Resume EXITP

End Sub

Private Sub chkShowHidden_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHidden_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim strTmp01 As String

11110   With Me
11120     Set frm = .frmAssets_Sub.Form
11130     Select Case .chkShowHidden
          Case True
11140       Select Case .opgShow
            Case .opgShow_optAll.OptionValue
11150         strFilter = vbNullString
11160         frm.Filter = strFilter
11170         frm.FilterOn = False
11180       Case .opgShow_optHeldOnly.OptionValue
11190         strFilter = HELD_FILTER
11200         frm.Filter = strFilter
11210         frm.FilterOn = True
11220       End Select
11230       .chkShowHidden_lbl.FontBold = True
11240     Case False
11250       Select Case .opgShow
            Case .opgShow_optAll.OptionValue
11260         strFilter = HID_FILTER
11270         frm.Filter = strFilter
11280         frm.FilterOn = True
11290       Case .opgShow_optHeldOnly.OptionValue
11300         strFilter = HID_FILTER & " And " & HELD_FILTER
11310         frm.Filter = strFilter
11320         frm.FilterOn = True
11330       End Select
11340       .chkShowHidden_lbl.FontBold = False
11350     End Select
11360     frm.Refresh
11370     strTmp01 = frm.SortNow_Get  ' ** Form Function: frmAssets_Sub.
11380     frm.OrderBy = strTmp01
11390     frm.OrderByOn = True
11400   End With

EXITP:
11410   Set frm = Nothing
11420   Exit Sub

ERRH:
11430   Select Case ERR.Number
        Case Else
11440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11450   End Select
11460   Resume EXITP

End Sub

Private Sub chkShowHidden_KeyDown(KeyCode As Integer, Shift As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHidden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11520   intShiftDown = (Shift And acShiftMask) > 0
11530   intAltDown = (Shift And acAltMask) > 0
11540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11570       With Me
11580         intRetVal = 0
11590         If .cmdPrintReport.Enabled = True Then
11600           .cmdPrintReport.SetFocus
11610         ElseIf .cmdEdit.Enabled = True Then
11620           .cmdEdit.SetFocus
11630         ElseIf .cmdAdd.Enabled = True Then
11640           .cmdAdd.SetFocus
11650         ElseIf .cmdUpdate.Enabled = True Then
11660           .cmdUpdate.SetFocus
11670         Else
11680           .cmdClose.SetFocus
11690         End If
11700       End With
11710     End Select
11720   End If

        ' ** Shift keys.
11730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11750       With Me
11760         intRetVal = 0
11770         If .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
11780           .chkShowAssetNo.SetFocus
11790         ElseIf .chkIncludeCurrency.Enabled = True Then
11800           .chkIncludeCurrency.SetFocus
11810         Else
11820           .opgShow.SetFocus
11830         End If
11840       End With
11850     End Select
11860   End If

EXITP:
11870   KeyCode = intRetVal
11880   Exit Sub

ERRH:
11890   intRetVal = 0
11900   THAT_PROC = THIS_PROC
11910   That_Erl = Erl
11920   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11930   Resume EXITP

End Sub

Public Sub cmdEdit_Click()

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_Click"

        Dim usr As DAO.User, grp As DAO.Group, frm As Access.Form
        Dim blnFound As Boolean

12010   blnFound = False

12020   For Each grp In DBEngine.Workspaces(0).Groups
12030     If grp.Name = "ViewOnly" Then
12040       For Each usr In grp.Users
12050         If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
12060           blnFound = True
12070           Exit For
12080         End If
12090       Next
12100     End If
12110   Next

12120   If blnFound Then
12130     MsgBox gstrNoPermission, vbInformation + vbOKOnly, "Access Denied"
12140   Else
12150     With Me
12160       Set frm = .frmAssets_Sub.Form
12170       If frm.cusip.Enabled = False And .cmdEdit.Tag = "Open" Then
              ' ** Somehow got out of sync.
12180         .cmdEdit.Tag = "Locked"
12190       End If
12200       Select Case .cmdEdit.Tag
            Case "Locked", vbNullString
12210         frm.cusip.Enabled = True
12220         frm.cusip.Locked = False
12230         frm.cusip.BackColor = CLR_WHT
12240         frm.cusip.ForeColor = CLR_BLK
12250         frm.cusip.BorderColor = CLR_LTBLU2
12260         frm.description_masterasset.Enabled = True
12270         frm.description_masterasset.Locked = False
12280         frm.description_masterasset.BackColor = CLR_WHT
12290         frm.description_masterasset.ForeColor = CLR_BLK
12300         frm.description_masterasset.BorderColor = CLR_LTBLU2
12310         frm.curr_id.Enabled = True
12320         frm.curr_id.Locked = False
12330         frm.curr_id.BackColor = CLR_WHT
12340         frm.curr_id.ForeColor = CLR_BLK
12350         frm.curr_id.BorderColor = CLR_LTBLU2
12360         frm.curr_id_usd_box.ForeColor = CLR_WHT
12370         frm.curr_id_forex_box.ForeColor = CLR_VLTGRN
12380         frm.rate_entry.Enabled = True
12390         frm.rate_entry.Locked = False
12400         frm.rate_entry.BackColor = CLR_WHT
12410         frm.rate_entry.ForeColor = CLR_BLK
12420         frm.rate_entry.BorderColor = CLR_LTBLU2
12430         frm.due.Enabled = True
12440         frm.due.Locked = False
12450         frm.due.BackColor = CLR_WHT
12460         frm.due.ForeColor = CLR_BLK
12470         frm.due.BorderColor = CLR_LTBLU2
12480         frm.cmbAssetType.Enabled = True
12490         frm.cmbAssetType.Locked = False
12500         frm.cmbAssetType.BackColor = CLR_WHT
12510         frm.cmbAssetType.ForeColor = CLR_BLK
12520         frm.cmbAssetType.BorderColor = CLR_LTBLU2
              ' ** shareface doesn't change.
              ' ** txtMAType doesn't change.
12530         .cmdAdd.Enabled = True
              '.cmdUpdate.Enabled = True  ' ** Not until they've changed something.
12540         .cmdEdit.Tag = "Open"
12550         .frmAssets_Sub.SetFocus
12560       Case "Open"
12570         If .chkUpdated = False Then
12580           frm.FocusHolder.SetFocus
12590           frm.cusip.Enabled = False
12600           frm.cusip.Locked = True
12610           frm.cusip.ForeColor = CLR_DISABLED_FG
12620           frm.cusip.BackColor = CLR_DISABLED_BG
12630           frm.cusip.BorderColor = WIN_CLR_DISR
12640           frm.description_masterasset.Enabled = False
12650           frm.description_masterasset.Locked = True
12660           frm.description_masterasset.ForeColor = CLR_DISABLED_FG
12670           frm.description_masterasset.BackColor = CLR_DISABLED_BG
12680           frm.description_masterasset.BorderColor = WIN_CLR_DISR
12690           frm.curr_id.Enabled = False
12700           frm.curr_id.Locked = True
12710           frm.curr_id.ForeColor = CLR_DISABLED_FG
12720           frm.curr_id.BackColor = CLR_DISABLED_BG
12730           frm.curr_id.BorderColor = WIN_CLR_DISR
12740           frm.curr_id_usd_box.ForeColor = CLR_DISABLED_BG
12750           frm.curr_id_forex_box.ForeColor = CLR_VLTGRN
12760           frm.rate_entry.Enabled = False
12770           frm.rate_entry.Locked = True
12780           frm.rate_entry.ForeColor = CLR_DISABLED_FG
12790           frm.rate_entry.BackColor = CLR_DISABLED_BG
12800           frm.rate_entry.BorderColor = WIN_CLR_DISR
12810           frm.due.Enabled = False
12820           frm.due.Locked = True
12830           frm.due.ForeColor = CLR_DISABLED_FG
12840           frm.due.BackColor = CLR_DISABLED_BG
12850           frm.due.BorderColor = WIN_CLR_DISR
12860           frm.cmbAssetType.Enabled = False
12870           frm.cmbAssetType.Locked = True
12880           frm.cmbAssetType.ForeColor = CLR_DISABLED_FG
12890           frm.cmbAssetType.BackColor = CLR_DISABLED_BG
12900           frm.cmbAssetType.BorderColor = WIN_CLR_DISR
                '.cmdAdd.Enabled = False
12910           .cmdUpdate.Enabled = False  ' ** They haven't changed anything.
12920           .cmdEdit.Tag = "Locked"
12930         Else
12940           Beep
12950           If .cmdUpdate.Enabled = True Then
12960             .cmdUpdate.SetFocus
12970             .cmdEdit.Enabled = False
12980           End If
12990         End If
13000       End Select
13010     End With
13020   End If

EXITP:
13030   Set frm = Nothing
13040   Set usr = Nothing
13050   Set grp = Nothing
13060   Exit Sub

ERRH:
13070   THAT_PROC = THIS_PROC
13080   That_Erl = Erl
13090   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13100   Resume EXITP

End Sub

Private Sub cmdEdit_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_KeyDown"

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
13290         If .cmdAdd.Enabled = True Then
13300           .cmdAdd.SetFocus
13310         ElseIf .cmdUpdate.Enabled = True Then
13320           .cmdUpdate.SetFocus
13330         Else
13340           .cmdClose.SetFocus
13350         End If
13360       End With
13370     End Select
13380   End If

        ' ** Shift keys.
13390   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13400     Select Case intRetVal
          Case vbKeyTab
13410       With Me
13420         intRetVal = 0
13430         If .cmdPrintReport.Enabled = True Then
13440           .cmdPrintReport.SetFocus
13450         ElseIf .chkShowHidden.Enabled = True Then
13460           .chkShowHidden.SetFocus
13470         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
13480           .chkShowAssetNo.SetFocus
13490         ElseIf .chkIncludeCurrency.Enabled = True Then
13500           .chkIncludeCurrency.SetFocus
13510         Else
13520           .opgShow.SetFocus
13530         End If
13540       End With
13550     End Select
13560   End If

EXITP:
13570   KeyCode = intRetVal
13580   Exit Sub

ERRH:
13590   intRetVal = 0
13600   THAT_PROC = THIS_PROC
13610   That_Erl = Erl
13620   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13630   Resume EXITP

End Sub

Private Sub cmdAdd_Click()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim usr As DAO.User, grp As DAO.Group
        Dim blnFound As Boolean

13710   blnFound = False

13720   For Each grp In DBEngine.Workspaces(0).Groups
13730     If grp.Name = "ViewOnly" Then
13740       For Each usr In grp.Users
13750         If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
13760           blnFound = True
13770           Exit For
13780         End If
13790       Next
13800     End If
13810   Next
13820   Set usr = Nothing
13830   Set grp = Nothing

13840   If blnFound = True Then
13850     MsgBox gstrNoPermission, vbInformation + vbOKOnly, "Access Denied"
13860   Else
13870     Select Case strCallingForm
          Case "frmJournal_Columns"
13880       DoCmd.OpenForm "frmAssets_Add", , , , acFormAdd, , THIS_NAME & "~" & strCallingForm
13890     Case Else
13900       DoCmd.OpenForm "frmAssets_Add", , , , acFormAdd, acDialog, THIS_NAME & "~" & strCallingForm
13910     End Select
13920   End If

EXITP:
13930   Set usr = Nothing
13940   Set grp = Nothing
13950   Exit Sub

ERRH:
13960   THAT_PROC = THIS_PROC
13970   That_Erl = Erl
13980   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13990   Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14020   intShiftDown = (Shift And acShiftMask) > 0
14030   intAltDown = (Shift And acAltMask) > 0
14040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14060     Select Case intRetVal
          Case vbKeyTab
14070       With Me
14080         intRetVal = 0
14090         If .cmdUpdate.Enabled = True Then
14100           .cmdUpdate.SetFocus
14110         Else
14120           .cmdClose.SetFocus
14130         End If
14140       End With
14150     End Select
14160   End If

        ' ** Shift keys.
14170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14180     Select Case intRetVal
          Case vbKeyTab
14190       With Me
14200         intRetVal = 0
14210         If .cmdAdd.Enabled = True Then
14220           .cmdAdd.SetFocus
14230         ElseIf .cmdPrintReport.Enabled = True Then
14240           .cmdPrintReport.SetFocus
14250         ElseIf .chkShowHidden.Enabled = True Then
14260           .chkShowHidden.SetFocus
14270         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
14280           .chkShowAssetNo.SetFocus
14290         ElseIf .chkIncludeCurrency.Enabled = True Then
14300           .chkIncludeCurrency.SetFocus
14310         Else
14320           .opgShow.SetFocus
14330         End If
14340       End With
14350     End Select
14360   End If

EXITP:
14370   KeyCode = intRetVal
14380   Exit Sub

ERRH:
14390   intRetVal = 0
14400   THAT_PROC = THIS_PROC
14410   That_Erl = Erl
14420   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14430   Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        'Dim cnxn As ADODB.Connection, rsx1 As ADODB.Recordset, rsx2 As ADODB.Recordset  ' ** Early binding.
        Dim cnxn As Object, rsx1 As Object, rsx2 As Object                               ' ** Late binding.
        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnInTrans As Boolean
        Dim varAssetNo As Variant
        Dim blnCurrChange As Boolean, lngCurrID_Old As Long, strCurrCode_Old As String, blnCurrConvert As Boolean
        Dim lngCurrID_New As Long, strCurrCode_New As String
        Dim lngCurrs As Long, arr_varCurr() As Variant
        Dim strDocName As String
        Dim lngRecs As Long
        Dim blnCloseForm As Boolean, blnContinue As Boolean
        Dim lngX As Long, lngE As Long

        ' ** Array: arr_varCurr().
        Const C_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const C_ASTNO As Integer = 0
        Const C_CIDO  As Integer = 1
        Const C_CIDN  As Integer = 2
        Const C_CODO  As Integer = 3
        Const C_CODN  As Integer = 4
        Const C_CONV  As Integer = 5

14510   blnContinue = True
14520   blnInTrans = False
14530   varAssetNo = 0
14540   blnCloseForm = False

14550   With Me

14560     If .chkUpdated = True Then

14570       RecalcNewAssets  ' ** Procedure: Below.
14580       RecalcForEx  ' ** Procedure: Below.

14590       DoCmd.Hourglass True
14600       DoEvents

14610       lngAssetNo = .frmAssets_Sub.Form.assetno

14620       lngCurrs = 0&
14630       ReDim arr_varCurr(C_ELEMS, 0)

14640       Set dbs = CurrentDb

            'Set rsx1 = New ADODB.Recordset             ' ** Early binding.
14650       Set rsx1 = CreateObject("ADODB.Recordset")  ' ** Late binding.
            ' ** tmpEdit09, for RecChanged <> 0.  #curr_id
14660       rsx1.Open "qryAssets_17", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTableDirect
14670       With rsx1  ' ** adOpenDynamic doesn't want to give a RecordCount!

14680         If .BOF = True And .EOF = True Then
                ' ** But really this should never happen.
                ' ** ADO: When you open a Recordset, the current record is positioned to the
                ' ** first record (if any) and the BOF and EOF properties are set to False.
                ' ** If there are no records, the BOF and EOF property settings are True.
14690           blnContinue = False
14700           DoCmd.Hourglass False
14710           MsgBox "No changes found!", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
14720         Else
14730           .MoveLast
14740           lngRecs = .RecordCount
14750           .MoveFirst
14760           DoEvents
14770           If lngRecs > 0 Then
14780             For lngX = 1& To lngRecs

14790               blnCurrChange = False

14800               If .Fields("RecChanged") = 1 Then  ' ** Changed.
                      ' ** Edit existing.
14810                 varAssetNo = .Fields("assetno")
14820               Else
                      ' ** New record, RecChanged = 2.
                      'Set rsx2 = New ADODB.Recordset             ' ** Early binding.
14830                 Set rsx2 = CreateObject("ADODB.Recordset")  ' ** Late binding.
14840                 Set cnxn = CurrentProject.Connection
14850                 rsx2.Open "masterasset", cnxn, adOpenDynamic, adLockOptimistic, adCmdTableDirect
14860                 cnxn.BeginTrans
14870                 blnInTrans = True
14880                 With rsx2
14890                   .AddNew
14900                   .Fields("cusip") = rsx1.Fields("cusip")
14910                   .Fields("assettype") = rsx1.Fields("assettype")
14920                   .Fields("description") = rsx1.Fields("description_masterasset")
14930                   .Fields("shareface") = 0#
14940                   .Fields("rate") = 0#
14950                   .Fields("marketvaluecurrent") = 0#
14960                   .Fields("yield") = 0#
14970                   .Fields("masterasset_TYPE") = "RA"
14980                   .Update
14990                   DoEvents
                        ' ** ADO: After you call the AddNew method, the new record becomes the current
                        ' ** record and remains current after you call the Update method.
15000                   varAssetNo = .Fields("assetno")
15010                   cnxn.CommitTrans
15020                   blnInTrans = False
15030                   .Close
15040                 End With  ' ** rsx2.
15050               End If
15060               DoEvents

15070               If IsNull(varAssetNo) Then
15080                 blnContinue = False
15090                 DoCmd.Hourglass False
15100                 MsgBox "Null Asset!", vbCritical + vbOKOnly, "Error"
15110               End If

15120               If blnContinue = True Then

15130                 lngCurrID_Old = 0&
15140                 strCurrCode_Old = vbNullString
15150                 blnCurrConvert = False

15160                 If .Fields("RecChanged") > 0 And .Fields("RecChanged") < 3 Then ' ** Edited existing record (already inserted if new).
                        'Set rsx2 = New ADODB.Recordset             ' ** Early binding.
15170                   Set rsx2 = CreateObject("ADODB.Recordset")  ' ** Late binding.
15180                   Set cnxn = CurrentProject.Connection
15190                   rsx2.Open "masterasset", cnxn, adOpenDynamic, adLockOptimistic, adCmdTableDirect
15200                   With rsx2
15210                     .MoveFirst
15220                     DoEvents
15230                     .Find "assetno = " & CStr(varAssetNo)
15240                     If .EOF = False Then
15250                       cnxn.BeginTrans
15260                       blnInTrans = True
15270                       .Fields("cusip") = rsx1.Fields("cusip")
15280                       .Fields("description") = rsx1.Fields("description_masterasset")
15290                       Select Case IsNull(rsx1.Fields("rate"))
                            Case True
15300                         .Fields("rate") = 0#
15310                       Case False
15320                         .Fields("rate") = rsx1.Fields("rate")
15330                       End Select
15340                       If IsNull(rsx1.Fields("due")) = False Then
15350                         .Fields("due") = rsx1.Fields("due")
15360                       Else
15370                         If IsNull(.Fields("due")) = False Then
15380                           .Fields("due") = Null
15390                         End If
15400                       End If
15410                       .Fields("assettype") = rsx1.Fields("assettype")
15420                       If .Fields("curr_id") <> rsx1.Fields("curr_id") Then
15430                         .Fields("curr_id") = rsx1.Fields("curr_id")
15440                       End If
15450                       If IsNull(rsx1.Fields("curr_id_old")) = False Then
                              ' ** It's assumed that curr_code_old will also not be Null.
15460                         blnCurrChange = True
15470                         lngCurrID_New = rsx1.Fields("curr_id")
15480                         lngCurrID_Old = rsx1.Fields("curr_id_old")
15490                         strCurrCode_New = rsx1.Fields("curr_code")
15500                         strCurrCode_Old = rsx1.Fields("curr_code_old")
15510                         blnCurrConvert = rsx1.Fields("curr_convert")
15520                       End If
15530                       .Update
15540                       DoEvents
15550                       cnxn.CommitTrans
15560                       blnInTrans = False
15570                     Else
                            ' ** Certainly should have found it!
15580                     End If
15590                     .Close
15600                   End With  ' ** rsx2.
15610                 ElseIf ![RecChanged] = 3 Then  ' ** Marked for deletion, though this isn't currently allowed.
                        ' ** Delete MasterAsset, by specified [astno] (not currently allowed).
15620                   Set qdf = dbs.QueryDefs("qryAssets_19")
15630                   With qdf.Parameters
15640                     ![astno] = CLng(varAssetNo)
15650                   End With
15660                   qdf.Execute dbFailOnError
15670                 End If
15680                 DoEvents

15690                 If blnCurrChange = True Then
15700                   lngCurrs = lngCurrs + 1&
15710                   lngE = lngCurrs - 1&
15720                   ReDim Preserve arr_varCurr(C_ELEMS, lngE)
15730                   arr_varCurr(C_ASTNO, lngE) = varAssetNo
15740                   arr_varCurr(C_CIDO, lngE) = lngCurrID_Old
15750                   arr_varCurr(C_CIDN, lngE) = lngCurrID_New
15760                   arr_varCurr(C_CODO, lngE) = strCurrCode_Old
15770                   arr_varCurr(C_CODN, lngE) = strCurrCode_New
15780                   arr_varCurr(C_CONV, lngE) = blnCurrConvert
15790                 End If
15800                 DoEvents

                      ' ** Update tmpEdit09, for RecChanged = 0, by specified [astno].
15810                 Set qdf = dbs.QueryDefs("qryAssets_20")
15820                 With qdf.Parameters
15830                   ![astno] = CLng(varAssetNo)
15840                 End With
15850                 qdf.Execute
15860                 Set qdf = Nothing
15870                 DoEvents

15880               End If  ' ** blnContinue.

15890               If blnContinue = False Then
15900                 Exit For
15910               Else
15920                 If lngX < lngRecs Then .MoveNext
15930               End If
15940             Next  ' ** lngX.
15950             DoEvents

                  ' ** An asset currency was changed!
15960             If lngCurrs > 0& Then
15970               UpdateCurrencyChange lngCurrs, arr_varCurr, dbs  ' ** Procedure: Below.
15980             End If
15990             DoEvents

16000             blnCloseForm = True

16010           Else
                  ' ** No, really, this should never happen.
16020             blnContinue = False
16030             DoCmd.Hourglass False
16040             MsgBox "No changes found!", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
16050           End If  ' ** lngRecs.
16060         End If  ' ** BOF, EOF.
16070         .Close
16080       End With  ' ** rsx1.
16090       dbs.Close
16100       DoEvents

16110       If blnCloseForm = True And blnContinue = True Then
16120         If .ForExCnt2 <> .ForExCnt1 Then
                ' ** The number of foreign currency assets has changed.
16130           gblnHasForEx = HasForEx_All  ' ** Module Function: modCurrencyFuncs.
16140           DoEvents
16150           HasForEx_Load  ' ** Module Procedure: modCurrencyFuncs.
16160           DoEvents
16170           gblnHasForExThis = gblnHasForEx
16180           DoEvents
16190           DoCmd.Hourglass True  ' ** Make sure it's still running.
16200           DoEvents
16210         End If
16220       End If

16230     Else
16240       blnContinue = False
16250       MsgBox "Nothing to update!", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
16260     End If  ' ** chkUpdated.

16270     DoCmd.Hourglass False

16280     .cmdClose.SetFocus
16290     If blnCloseForm = True And blnContinue = True Then
16300       .chkUpdated = False
16310       strDocName = "frmAssets_Update"
16320       gblnMessage = True
16330       DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME
16340       Select Case gblnMessage
            Case True
16350         cmdClose_Click  ' ** Procedure: Above.
16360       Case False
16370         .cmdClose.SetFocus
16380         ReloadAssets  ' ** Procedure: Below.
16390         blnHasForEx_Chk = True
16400         .TimerInterval = 200&
16410       End Select
16420     End If

16430   End With  ' ** Me.

EXITP:
16440   Set rsx1 = Nothing
16450   Set rsx2 = Nothing
16460   Set cnxn = Nothing
16470   Set qdf = Nothing
16480   Set dbs = Nothing
16490   Exit Sub

ERRH:
16500   DoCmd.Hourglass False
16510   If blnInTrans Then
          'wrk.Rollback
16520     MsgBox "Error updating: " & vbCrLf & ERR.description & vbCrLf & _
            "Error: " & CStr(ERR.Number) & vbCrLf & "Procedure: " & THIS_PROC & "()" & vbCrLf & _
            "Line: " & Erl, vbCritical + vbOKOnly, "Transaction Error"
16530   End If
16540   THAT_PROC = THIS_PROC
16550   That_Erl = Erl
16560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16570   Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16620   intShiftDown = (Shift And acShiftMask) > 0
16630   intAltDown = (Shift And acAltMask) > 0
16640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16660     Select Case intRetVal
          Case vbKeyTab
16670       With Me
16680         intRetVal = 0
16690         .cmdClose.SetFocus
16700       End With
16710     End Select
16720   End If

        ' ** Shift keys.
16730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16740     Select Case intRetVal
          Case vbKeyTab
16750       With Me
16760         intRetVal = 0
16770         If .cmdAdd.Enabled = True Then
16780           .cmdAdd.SetFocus
16790         ElseIf .cmdEdit.Enabled = True Then
16800           .cmdEdit.SetFocus
16810         ElseIf .cmdPrintReport.Enabled = True Then
16820           .cmdPrintReport.SetFocus
16830         ElseIf .chkShowHidden.Enabled = True Then
16840           .chkShowHidden.SetFocus
16850         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
16860           .chkShowAssetNo.SetFocus
16870         ElseIf .chkIncludeCurrency.Enabled = True Then
16880           .chkIncludeCurrency.SetFocus
16890         Else
16900           .opgShow.SetFocus
16910         End If
16920       End With
16930     End Select
16940   End If

EXITP:
16950   KeyCode = intRetVal
16960   Exit Sub

ERRH:
16970   intRetVal = 0
16980   THAT_PROC = THIS_PROC
16990   That_Erl = Erl
17000   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17010   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

17110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17120   Exit Sub

ERRH:
17130   Select Case ERR.Number
        Case Else
17140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17150   End Select
17160   Resume EXITP

End Sub

Private Sub chkOrderBy_AfterUpdate()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOrderBy_AfterUpdate"

17210   With Me
17220     Select Case .chkOrderBy
          Case True
17230       .chkOrderBy_lbl.FontBold = True
17240     Case False
17250       .chkOrderBy_lbl.FontBold = False
17260     End Select
17270   End With

        ' ** [cusip], [description_masterasset_sort], [assettype]
        ' ** [cusip] DESC, [description_masterasset_sort] DESC, [assettype] DESC
        ' ** [description_masterasset_sort], [cusip], [assettype]
        ' ** [description_masterasset_sort] DESC, [cusip] DESC, [assettype] DESC
        ' ** [rate], [description_masterasset_sort], [assettype]
        ' ** [rate] DESC, [description_masterasset_sort] DESC, [assettype] DESC
        ' ** [due], [description_masterasset_sort], [assettype]
        ' ** [due] DESC, [description_masterasset_sort] DESC, [assettype] DESC
        ' ** [shareface], [description_masterasset_sort], [assettype]
        ' ** [shareface] DESC, [description_masterasset_sort] DESC, [assettype] DESC
        ' ** [assettype], [description_masterasset_sort], [cusip]
        ' ** [assettype] DESC, [description_masterasset_sort] DESC, [cusip] DESC

EXITP:
17280   Exit Sub

ERRH:
17290   Select Case ERR.Number
        Case Else
17300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17310   End Select
17320   Resume EXITP

End Sub

Private Function LoadIAData() As Boolean
' ** Load the current Accrued Interest Asset values into the arr_varIA_Asset() array,
' ** and determine whether it's been changed since installation or conversion.

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "LoadIAData"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim arr_varTmp00() As Variant
        Dim blnRetVal As Boolean

17410   Set dbs = CurrentDb
17420   With dbs
          ' ** tmpEdit09, just the Accrued Interest Asset, 'IA'.
17430     Set qdf = .QueryDefs("qryAssets_08")
17440     Set rst = qdf.OpenRecordset
17450     With rst
17460       If .BOF = True And .EOF = True Then
              ' ** No 'IA' asset?
17470         ReDim arr_varTmp00(11, 0)
17480         arr_varTmp00(IA_MTYPE, 0) = "IA"
17490         arr_varTmp00(IA_ID, 0) = 0&
17500         arr_varTmp00(IA_DESC, 0) = RET_ERR
17510         arr_varIA_Asset = arr_varTmp00
17520       Else
17530         .MoveFirst
17540         arr_varIA_Asset = .GetRows(1)
              ' ************************************************************
              ' ** Array: arr_varIA_Asset()
              ' **
              ' **   Field  Element  Name                       Constant
              ' **   =====  =======  =========================  ==========
              ' **     1       0     masterasset_TYPE           IA_MTYPE
              ' **     2       1     assetno                    IA_ID
              ' **     3       2     cusip                      IA_CUSIP
              ' **     4       3     description_masterasset    IA_DESC
              ' **     5       4     shareface
              ' **     6       5     assettype                  IA_ATYPE
              ' **     7       6     rate                       IA_RATE
              ' **     8       7     due                        IA_DUE
              ' **     9       8     marketvalue
              ' **    10       9     marketvaluecurrent
              ' **    11      10     yield
              ' **    12      11     currentDate
              ' **
              ' ************************************************************
17550       End If
17560       .Close
17570     End With
17580     If arr_varIA_Asset(IA_DESC, 0) <> RET_ERR Then
17590       Set rst = .OpenRecordset("tblTemplate_MasterAsset", dbOpenDynaset, dbReadOnly)
17600       With rst
17610         If .BOF = True And .EOF = True Then
                ' ** Hmmm. Didn't get populated at installation or conversion.
17620         Else
17630           .MoveFirst
17640           If arr_varIA_Asset(IA_CUSIP, 0) <> ![cusip] Then
17650             blnIAHasChanged = True
17660           Else
17670             If arr_varIA_Asset(IA_DESC, 0) <> ![description] Then  ' ** tmpEdit09 has a different name.
17680               blnIAHasChanged = True
17690             Else
17700               If arr_varIA_Asset(IA_ATYPE, 0) <> ![assettype] Then
17710                 blnIAHasChanged = True
17720               Else
17730                 If IsNull(arr_varIA_Asset(IA_RATE, 0)) = True And IsNull(![rate]) = False Then
17740                   blnIAHasChanged = True
17750                 Else
17760                   If IsNull(arr_varIA_Asset(IA_RATE, 0)) = False And IsNull(![rate]) = True Then
17770                     blnIAHasChanged = True
17780                   Else
17790                     If IsNull(arr_varIA_Asset(IA_RATE, 0)) = False And IsNull(![rate]) = False Then
17800                       If arr_varIA_Asset(IA_RATE, 0) <> ![rate] Then
17810                         blnIAHasChanged = True
17820                       End If
17830                     End If
17840                   End If
17850                 End If
17860                 If blnIAHasChanged = False Then
17870                   If IsNull(arr_varIA_Asset(IA_DUE, 0)) = True And IsNull(![due]) = False Then
17880                     blnIAHasChanged = True
17890                   Else
17900                     If IsNull(arr_varIA_Asset(IA_DUE, 0)) = False And IsNull(![due]) = True Then
17910                       blnIAHasChanged = True
17920                     Else
17930                       If IsNull(arr_varIA_Asset(IA_DUE, 0)) = False And IsNull(![due]) = False Then
17940                         If arr_varIA_Asset(IA_DUE, 0) <> ![due] Then
17950                           blnIAHasChanged = True
17960                         End If
17970                       End If
17980                     End If
17990                   End If
18000                 End If
18010               End If
18020             End If
18030           End If
18040         End If
18050         .Close
18060       End With
18070     End If
18080     .Close
18090   End With

18100   Me.frmAssets_Sub.Form.IAHasChanged_Set blnIAHasChanged  ' ** Form Procedure: frmAssets_Sub.

EXITP:
18110   Set rst = Nothing
18120   Set qdf = Nothing
18130   Set dbs = Nothing
18140   LoadIAData = blnRetVal
18150   Exit Function

ERRH:
18160   blnRetVal = False
18170   ReDim arr_varTmp00(11, 0)
18180   arr_varTmp00(IA_MTYPE, 0) = "IA"
18190   arr_varTmp00(IA_ID, 0) = 0&
18200   arr_varTmp00(IA_DESC, 0) = RET_ERR
18210   arr_varIA_Asset = arr_varTmp00
18220   Select Case ERR.Number
        Case Else
18230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18240   End Select
18250   Resume EXITP

End Function

Public Function IAHasChanged_Get() As Boolean

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "IAHasChanged_Get"

        Dim blnRetVal As Boolean

18310   blnRetVal = blnIAHasChanged

EXITP:
18320   IAHasChanged_Get = blnRetVal
18330   Exit Function

ERRH:
18340   blnRetVal = False
18350   Select Case ERR.Number
        Case Else
18360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18370   End Select
18380   Resume EXITP

End Function

Public Sub IAHasChanged_Set(blnSet As Boolean)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "IAHasChanged_Set"

18410   blnIAHasChanged = blnSet

EXITP:
18420   Exit Sub

ERRH:
18430   Select Case ERR.Number
        Case Else
18440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18450   End Select
18460   Resume EXITP

End Sub

Private Sub ReloadAssets()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "ReloadAssets"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim strRecSrc As String

18510   With Me

18520     DoCmd.Hourglass True
18530     DoEvents

18540     .chkReloading = True

18550     Set frm = .frmAssets_Sub.Form

18560     With frm
18570       .FocusHolder.SetFocus
18580       .cusip.ForeColor = CLR_WHT
18590       .description_masterasset.ForeColor = CLR_WHT
18600       .curr_id.ForeColor = CLR_WHT
18610       .rate_entry.ForeColor = CLR_WHT
18620       .due.ForeColor = CLR_WHT
18630       .shareface.ForeColor = CLR_DISABLED_BG
18640       .cmbAssetType.ForeColor = CLR_WHT
18650       DoEvents
18660       strRecSrc = .RecordSource
18670       .RecordSource = vbNullString
18680     End With
18690     DoEvents

18700     Set dbs = CurrentDb
18710     With dbs

            ' ** Empty tmpEdit09.
18720       Set qdf = .QueryDefs("qryAssets_16")
18730       qdf.Execute

            ' ** Append qryAssets_03 (qryAssets_02 (MasterAsset, linked to assettype), with add'l fields),
            ' ** linked to qryAssets_04 (qryAssets_04c (Union of qryAssets_04a (Ledger, grouped by assetno, with cnt),
            ' ** qryAssets_04b (LedgerArchive, grouped by assetno, with cnt)), grouped and summed by assetno), to tmpEdit09.
18740       Set qdf = .QueryDefs("qryAssets_05")
18750       qdf.Execute

            ' ** tmpEdit09, summed by IsHid.
18760       Set qdf = .QueryDefs("qryAssets_07")
18770       Set rst = qdf.OpenRecordset
18780       With rst
18790         If .BOF = True And .EOF = True Then
18800           lngIsHidSum = 0&
18810         Else
18820           .MoveFirst
18830           lngIsHidSum = ![IsHidSum]
18840         End If
18850         .Close
18860       End With

18870       .Close
18880     End With

18890     .IsHidSum = lngIsHidSum
18900     lngIsHid_Width = .chkShowHidden_lbl.Width
18910     If lngIsHidSum = 0& Then
18920       .chkShowHidden.Enabled = False
18930       .chkShowHidden_lbl.Visible = False
18940       .chkShowHidden_lbl2_dim.Visible = True
18950       .chkShowHidden_lbl2_dim_hi.Visible = True
18960       .chkShowHidden_lbl3_dim.Visible = True
18970       .chkShowHidden_lbl3_dim_hi.Visible = True
18980     End If
18990     frm.IsHid_Width_Set lngIsHid_Width  ' ** Form Procedure: frmAssets_Sub.
19000     DoEvents

          ' ** Pass some values to the subform.
19010     frm.IAHasChanged_Set blnIAHasChanged  ' ** Form Procedure: frmAssets_Sub.
19020     frm.ParForm_Width_Set lngForm_Width  ' ** Form Procedure: frmAssets_Sub.

19030     chkShowHidden_AfterUpdate  ' ** Procedure: Below.
19040     .chkReloading = False
19050     DoEvents

19060     With frm
19070       .RecordSource = strRecSrc
19080       DoEvents
19090       .cusip.ForeColor = CLR_BLK
19100       .description_masterasset.ForeColor = CLR_BLK
19110       .curr_id.ForeColor = CLR_BLK
19120       .rate_entry.ForeColor = CLR_BLK
19130       .due.ForeColor = CLR_BLK
19140       .shareface.ForeColor = CLR_DISABLED_FG
19150       .cmbAssetType.ForeColor = CLR_BLK
19160     End With

          ' ** Load arr_varIA_Asset().
19170     blnIAHasChanged = False
19180     LoadIAData  ' ** Function: Above.

19190     RecalcNewAssets  ' ** Procedure: Below.
19200     RecalcForEx  ' ** Procedure: Below.

19210     With frm
19220       .FilterOn = False
19230       .OrderByOn = False
19240       .FilterOn = True
19250       .OrderByOn = True
19260       If lngAssetNo <> 0& Then
19270         .MoveRec 0, lngAssetNo  ' ** Form Procedure: frmAssets_Sub.
19280       End If
19290     End With

          ' ** Close the subform.
19300     cmdEdit_Click  ' ** Procedure: Above.
19310     DoEvents
19320     .cmdEdit.Enabled = True

19330     DoCmd.Hourglass False

19340   End With

EXITP:
19350   Set rst = Nothing
19360   Set qdf = Nothing
19370   Set dbs = Nothing
19380   Exit Sub

ERRH:
19390   DoCmd.Hourglass False
19400   Select Case ERR.Number
        Case Else
19410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19420   End Select
19430   Resume EXITP

End Sub

Private Sub IncludeCurrency(blnShow As Boolean)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngNewFrmWidth As Long, lngNewFrmLeft As Long, lngNewFormWidth As Long
        Dim lngTmp01 As Long

19510   With Me

19520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
19530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
19540     End If

19550     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
19560     lngMonitorNum = 1&: lngTmp01 = 0&
19570     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
19580     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

19590     Select Case blnShow
          Case True
19600       lngNewFrmWidth = lngFrm_Width
19610       lngNewFrmLeft = lngFrm_Left
19620       lngNewFormWidth = lngForm_Width
19630       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
19640       DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
19650       If lngMonitorNum > 1& Then
19660         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
19670       End If
19680       .Width = lngNewFormWidth
19690       .cmdClose.Left = lngClose_Left
19700       .cmdUpdate.Left = lngUpdate_Left
19710       .cmdAdd.Left = lngAdd_Left
19720       .cmdEdit.Left = lngEdit_Left
19730       .frmAssets_Sub.Width = lngSub_Width
19740       .frmAssets_Sub_box.Width = (.frmAssets_Sub.Width + (2& * lngTpp))
19750       .Nav_box01.Width = .frmAssets_Sub.Width
19760       .Nav_hline03.Width = .frmAssets_Sub.Width
19770       .Header_vline01.Left = lngNewFormWidth
19780       .Header_vline02.Left = lngNewFormWidth
19790       .Footer_vline01.Left = lngNewFormWidth
19800       .Footer_vline02.Left = lngNewFormWidth
19810       .Header_hline01.Width = lngNewFormWidth
19820       .Header_hline02.Width = lngNewFormWidth
19830       .Footer_hline01.Width = lngNewFormWidth
19840       .Footer_hline02.Width = lngNewFormWidth
19850     Case False
19860       lngNewFrmWidth = (lngFrm_Width - lngCurrID_Width)
19870       lngNewFrmLeft = (lngFrm_Left + (lngCurrID_Width / 2))
19880       lngNewFormWidth = (lngForm_Width - lngCurrID_Width)
19890       .cmdClose.Left = (lngClose_Left - lngCurrID_Width)
19900       .cmdUpdate.Left = (lngUpdate_Left - lngCurrID_Width)
19910       .cmdAdd.Left = (lngAdd_Left - lngCurrID_Width)
19920       .cmdEdit.Left = (lngEdit_Left - lngCurrID_Width)
19930       .frmAssets_Sub.Width = (lngSub_Width - lngCurrID_Width)
19940       .frmAssets_Sub_box.Width = (.frmAssets_Sub.Width + (2& * lngTpp))
19950       .Nav_box01.Width = .frmAssets_Sub.Width
19960       .Nav_hline03.Width = .frmAssets_Sub.Width
19970       .Header_vline01.Left = lngNewFormWidth
19980       .Header_vline02.Left = lngNewFormWidth
19990       .Footer_vline01.Left = lngNewFormWidth
20000       .Footer_vline02.Left = lngNewFormWidth
20010       .Header_hline01.Width = lngNewFormWidth
20020       .Header_hline02.Width = lngNewFormWidth
20030       .Footer_hline01.Width = lngNewFormWidth
20040       .Footer_hline02.Width = lngNewFormWidth
20050       .Width = lngNewFormWidth
20060       If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
20070       DoCmd.MoveSize lngNewFrmLeft, lngTmp01, lngNewFrmWidth, lngFrm_Height  'lngFrm_Top
20080       If lngMonitorNum > 1& Then
20090         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
20100       End If
20110     End Select

20120   End With

EXITP:
20130   Exit Sub

ERRH:
20140   THAT_PROC = THIS_PROC
20150   That_Erl = Erl
20160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
20170   Resume EXITP

End Sub

Private Sub UpdateCurrencyChange(lngCurrs As Long, arr_varCurr As Variant, dbs As DAO.Database)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateCurrencyChange"

        Dim qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim varRate1 As Variant, varRate2 As Variant
        Dim lngX As Long, lngE As Long

        ' ** Array: arr_varCurr().
        'Const C_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const C_ASTNO As Integer = 0
        Const C_CIDO  As Integer = 1
        Const C_CIDN  As Integer = 2
        'Const C_CODO  As Integer = 3
        'Const C_CODN  As Integer = 4
        Const C_CONV  As Integer = 5

20210   For lngX = 0& To (lngCurrs - 1&)

20220     Select Case arr_varCurr(C_CONV, lngE)
          Case True
            ' ** Convert values from the old currency to the new.

20230       If arr_varCurr(C_CIDO, lngX) = 150& Then
              ' ** Get the rate for converting USD to this new currency.

20240         varRate1 = DLookup("[curr_rate1]", "tblCurrency", "[curr_id] = " & CStr(arr_varCurr(C_CIDN, lngX)))
20250         If IsNull(varRate1) = False Then
20260           If varRate1 > 0 Then

                  ' ** Update qryAssets_Currency_03_01_01 (ActiveAssets, for USD to other,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate1]).
20270             Set qdf = dbs.QueryDefs("qryAssets_Currency_03_01_02")
20280             With qdf.Parameters
20290               ![astno] = arr_varCurr(C_ASTNO, lngX)
20300               ![curridnew] = arr_varCurr(C_CIDN, lngX)
20310               ![crate1] = varRate1
20320             End With
20330             qdf.Execute
20340             Set qdf = Nothing
20350             DoEvents

                  ' ** Update qryAssets_Currency_03_02_01 (Ledger, for USD to other,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate1]).
20360             Set qdf = dbs.QueryDefs("qryAssets_Currency_03_02_02")
20370             With qdf.Parameters
20380               ![astno] = arr_varCurr(C_ASTNO, lngX)
20390               ![curridnew] = arr_varCurr(C_CIDN, lngX)
20400               ![crate1] = varRate1
20410             End With
20420             qdf.Execute
20430             Set qdf = Nothing
20440             DoEvents

                  ' ** Update qryAssets_Currency_03_03_01 (LedgerArchive, for USD to other,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate1]).
20450             Set qdf = dbs.QueryDefs("qryAssets_Currency_03_03_02")
20460             With qdf.Parameters
20470               ![astno] = arr_varCurr(C_ASTNO, lngX)
20480               ![curridnew] = arr_varCurr(C_CIDN, lngX)
20490               ![crate1] = varRate1
20500             End With
20510             qdf.Execute
20520             Set qdf = Nothing
20530             DoEvents

                  ' ** Update qryAssets_Currency_03_04_01 (Journal, for USD to other,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate1]).
20540             Set qdf = dbs.QueryDefs("qryAssets_Currency_03_04_02")
20550             With qdf.Parameters
20560               ![astno] = arr_varCurr(C_ASTNO, lngX)
20570               ![curridnew] = arr_varCurr(C_CIDN, lngX)
20580               ![crate1] = varRate1
20590             End With
20600             qdf.Execute
20610             Set qdf = Nothing
20620             DoEvents

                  ' ** qryAssets_Currency_03_05_01 (tblPricing_MasterAsset_History, for
                  ' ** USD to other, today's date, with .._new fields, by specified [astno],
                  ' ** [curridnew], [crate1]), linked to tblCurrency, with .._new fields.
20630             Set qdf = dbs.QueryDefs("qryAssets_Currency_03_05_02")
20640             With qdf.Parameters
20650               ![astno] = arr_varCurr(C_ASTNO, lngX)
20660               ![curridnew] = arr_varCurr(C_CIDN, lngX)
20670               ![crate1] = varRate1
20680             End With
20690             Set rst1 = qdf.OpenRecordset
20700             If rst1.BOF = True And rst1.EOF = True Then
                    ' ** No applicable entries.
20710               rst1.Close
20720             Else
20730               Set rst2 = dbs.OpenRecordset("tblPricing_MasterAsset_History", dbOpenDynaset, dbConsistent)
20740               With rst1
20750                 .MoveFirst  ' ** There can be only one.
20760                 DoEvents
20770                 With rst2
20780                   .FindFirst "[AssetPricing_ID] = " & CStr(rst1![AssetPricing_ID])
20790                   If .NoMatch = False Then
20800                     .Edit
20810                     ![marketvaluecurrent] = rst1![marketvaluecurrent_new]
20820                     ![curr_id] = rst1![curr_id_new]
20830                     ![curr_code] = rst1![curr_code_new]
20840                     ![currsym_symbol] = rst1![currsym_symbol_new]
20850                     ![curr_date] = rst1![curr_date_new]
20860                     ![curr_rate1] = rst1![curr_rate1_new]
20870                     ![curr_rate2] = rst1![curr_rate2_new]
20880                     ![CurrentJournalUser] = CurrentUser()  ' ** Internal Access Function: Trust Accountant login.
20890                     ![DateModified] = Now()
20900                     .Update
20910                   End If
20920                   .Close
20930                 End With
20940                 .Close
20950               End With
20960             End If
20970             Set rst1 = Nothing
20980             Set rst2 = Nothing
20990             Set qdf = Nothing
21000             DoEvents

21010           Else
                  ' ** Else what?
21020           End If
21030         Else
                ' ** Else what?
21040         End If

21050       ElseIf arr_varCurr(C_CIDN, lngX) = 150& Then
              ' ** Get the rate for converting this currency to USD.

21060         varRate2 = DLookup("[curr_rate2]", "tblCurrency", "[curr_id] = " & CStr(arr_varCurr(C_CIDN, lngX)))
21070         If IsNull(varRate1) = False Then
21080           If varRate2 > 0 Then

                  ' ** Update qryAssets_Currency_04_01_01 (ActiveAssets, for other to USD,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate2]).
21090             Set qdf = dbs.QueryDefs("qryAssets_Currency_04_01_02")
21100             With qdf.Parameters
21110               ![astno] = arr_varCurr(C_ASTNO, lngX)
21120               ![curridnew] = arr_varCurr(C_CIDN, lngX)
21130               ![crate2] = varRate2
21140             End With
21150             qdf.Execute
21160             Set qdf = Nothing
21170             DoEvents

                  ' ** Update qryAssets_Currency_04_02_01 (Ledger, for other to USD,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate2]).
21180             Set qdf = dbs.QueryDefs("qryAssets_Currency_04_02_02")
21190             With qdf.Parameters
21200               ![astno] = arr_varCurr(C_ASTNO, lngX)
21210               ![curridnew] = arr_varCurr(C_CIDN, lngX)
21220               ![crate2] = varRate2
21230             End With
21240             qdf.Execute
21250             Set qdf = Nothing
21260             DoEvents

                  ' ** Update qryAssets_Currency_04_03_01 (LedgerArchive, for other to USD,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate2]).
21270             Set qdf = dbs.QueryDefs("qryAssets_Currency_04_03_02")
21280             With qdf.Parameters
21290               ![astno] = arr_varCurr(C_ASTNO, lngX)
21300               ![curridnew] = arr_varCurr(C_CIDN, lngX)
21310               ![crate2] = varRate2
21320             End With
21330             qdf.Execute
21340             Set qdf = Nothing
21350             DoEvents

                  ' ** Update qryAssets_Currency_04_04_01 (Journal, for other to USD,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate2]).
21360             Set qdf = dbs.QueryDefs("qryAssets_Currency_04_04_02")
21370             With qdf.Parameters
21380               ![astno] = arr_varCurr(C_ASTNO, lngX)
21390               ![curridnew] = arr_varCurr(C_CIDN, lngX)
21400               ![crate2] = varRate2
21410             End With
21420             qdf.Execute
21430             Set qdf = Nothing
21440             DoEvents

                  ' ** qryAssets_Currency_04_05_01 (tblPricing_MasterAsset_History, for
                  ' ** other to USD, today's date, with .._new fields, by specified [astno],
                  ' ** [curridnew], [crate2]), linked to tblCurrency, with .._new fields.
21450             Set qdf = dbs.QueryDefs("qryAssets_Currency_04_05_02")
21460             With qdf.Parameters
21470               ![astno] = arr_varCurr(C_ASTNO, lngX)
21480               ![curridnew] = arr_varCurr(C_CIDN, lngX)
21490               ![crate2] = varRate2
21500             End With
21510             Set rst1 = qdf.OpenRecordset
21520             If rst1.BOF = True And rst1.EOF = True Then
                    ' ** No applicable entries.
21530               rst1.Close
21540             Else
21550               Set rst2 = dbs.OpenRecordset("tblPricing_MasterAsset_History", dbOpenDynaset, dbConsistent)
21560               With rst1
21570                 .MoveFirst  ' ** There can be only one.
21580                 DoEvents
21590                 With rst2
21600                   .FindFirst "[AssetPricing_ID] = " & CStr(rst1![AssetPricing_ID])
21610                   If .NoMatch = False Then
21620                     .Edit
21630                     ![marketvaluecurrent] = rst1![marketvaluecurrent_new]
21640                     ![curr_id] = rst1![curr_id_new]
21650                     ![curr_code] = rst1![curr_code_new]
21660                     ![currsym_symbol] = rst1![currsym_symbol_new]
21670                     ![curr_date] = rst1![curr_date_new]
21680                     ![curr_rate1] = rst1![curr_rate1_new]
21690                     ![curr_rate2] = rst1![curr_rate2_new]
21700                     ![CurrentJournalUser] = CurrentUser()  ' ** Internal Access Function: Trust Accountant login.
21710                     ![DateModified] = Now()
21720                     .Update
21730                   End If
21740                   .Close
21750                 End With
21760                 .Close
21770               End With
21780             End If
21790             Set rst1 = Nothing
21800             Set rst2 = Nothing
21810             Set qdf = Nothing
21820             DoEvents

21830           Else
                  ' ** Else what?
21840           End If
21850         Else
                ' ** Else what?
21860         End If

21870       Else
              ' ** Get the rates for converting this currency to another, non-USD currency.
              ' ** This is other to other, and since I don't have any
              ' ** cross-currency rates, I'll have to do it through USD.

21880         varRate1 = DLookup("[curr_rate1]", "tblCurrency", "[curr_id] = " & CStr(arr_varCurr(C_CIDN, lngX)))
21890         varRate2 = DLookup("[curr_rate2]", "tblCurrency", "[curr_id] = " & CStr(arr_varCurr(C_CIDO, lngX)))
21900         If IsNull(varRate1) = False And IsNull(varRate2) = False Then
21910           If varRate1 > 0 And varRate2 > 0 Then

                  ' ** Update qryAssets_Currency_05_01_01 (ActiveAssets, for other to other,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate1], [crate2]).
21920             Set qdf = dbs.QueryDefs("qryAssets_Currency_05_01_02")
21930             With qdf.Parameters
21940               ![astno] = arr_varCurr(C_ASTNO, lngX)
21950               ![curridnew] = arr_varCurr(C_CIDN, lngX)
21960               ![crate1] = varRate1
21970               ![crate2] = varRate2
21980             End With
21990             qdf.Execute
22000             Set qdf = Nothing
22010             DoEvents

                  ' ** Update qryAssets_Currency_05_02_01 (Ledger, for other to other, with
                  ' ** .._new fields, by specified [astno], [curridnew], [crate1], [crate2]).
22020             Set qdf = dbs.QueryDefs("qryAssets_Currency_05_02_02")
22030             With qdf.Parameters
22040               ![astno] = arr_varCurr(C_ASTNO, lngX)
22050               ![curridnew] = arr_varCurr(C_CIDN, lngX)
22060               ![crate1] = varRate1
22070               ![crate2] = varRate2
22080             End With
22090             qdf.Execute
22100             Set qdf = Nothing
22110             DoEvents

                  ' ** Update qryAssets_Currency_05_03_01 (LedgerArchive, for other to other,
                  ' ** with .._new fields, by specified [astno], [curridnew], [crate1], [crate2]).
22120             Set qdf = dbs.QueryDefs("qryAssets_Currency_05_03_02")
22130             With qdf.Parameters
22140               ![astno] = arr_varCurr(C_ASTNO, lngX)
22150               ![curridnew] = arr_varCurr(C_CIDN, lngX)
22160               ![crate1] = varRate1
22170               ![crate2] = varRate2
22180             End With
22190             qdf.Execute
22200             Set qdf = Nothing
22210             DoEvents

                  ' ** Update qryAssets_Currency_05_04_01 (Journal, for other to other, with
                  ' ** .._new fields, by specified [astno], [curridnew], [crate1], [crate2]).
22220             Set qdf = dbs.QueryDefs("qryAssets_Currency_05_04_02")
22230             With qdf.Parameters
22240               ![astno] = arr_varCurr(C_ASTNO, lngX)
22250               ![curridnew] = arr_varCurr(C_CIDN, lngX)
22260               ![crate1] = varRate1
22270               ![crate2] = varRate2
22280             End With
22290             qdf.Execute
22300             Set qdf = Nothing
22310             DoEvents

                  ' ** tblPricing_MasterAsset_05_05_01 (tblPricing_MasterAsset_History, for
                  ' ** other to other, today's date, with .._new fields, by specified [astno],
                  ' ** [curridnew], [crate1], [crate2]), linked to tblCurrency, with .._new fields.
22320             Set qdf = dbs.QueryDefs("qryAssets_Currency_05_05_02")
22330             With qdf.Parameters
22340               ![astno] = arr_varCurr(C_ASTNO, lngX)
22350               ![curridnew] = arr_varCurr(C_CIDN, lngX)
22360               ![crate1] = varRate1
22370               ![crate2] = varRate2
22380             End With
22390             Set rst1 = qdf.OpenRecordset
22400             If rst1.BOF = True And rst1.EOF = True Then
                    ' ** No applicable entries.
22410               rst1.Close
22420             Else
22430               Set rst2 = dbs.OpenRecordset("tblPricing_MasterAsset_History", dbOpenDynaset, dbConsistent)
22440               With rst1
22450                 .MoveFirst  ' ** There can be only one.
22460                 DoEvents
22470                 With rst2
22480                   .FindFirst "[AssetPricing_ID] = " & CStr(rst1![AssetPricing_ID])
22490                   If .NoMatch = False Then
22500                     .Edit
22510                     ![marketvaluecurrent] = rst1![marketvaluecurrent_new]
22520                     ![curr_id] = rst1![curr_id_new]
22530                     ![curr_code] = rst1![curr_code_new]
22540                     ![currsym_symbol] = rst1![currsym_symbol_new]
22550                     ![curr_date] = rst1![curr_date_new]
22560                     ![curr_rate1] = rst1![curr_rate1_new]
22570                     ![curr_rate2] = rst1![curr_rate2_new]
22580                     ![CurrentJournalUser] = CurrentUser()  ' ** Internal Access Function: Trust Accountant login.
22590                     ![DateModified] = Now()
22600                     .Update
22610                   End If
22620                   .Close
22630                 End With
22640                 .Close
22650               End With
22660             End If
22670             Set rst1 = Nothing
22680             Set rst2 = Nothing
22690             Set qdf = Nothing
22700             DoEvents

22710           Else
                  ' ** Else what?
22720           End If
22730         Else
                ' ** Else what?
22740         End If

22750       End If

22760     Case False
            ' ** No conversion, leave values as-is. (Which may mean cost_usd and market_usd are messed up!)

            ' ** Update qryAssets_Currency_02_01_01 (ActiveAssets, with curr_id_new, by specified [astno], [curridnew]).
22770       Set qdf = dbs.QueryDefs("qryAssets_Currency_02_01_02")
22780       With qdf.Parameters
22790         ![astno] = arr_varCurr(C_ASTNO, lngX)
22800         ![curridnew] = arr_varCurr(C_CIDN, lngX)
22810       End With
22820       qdf.Execute
22830       Set qdf = Nothing
22840       DoEvents

            ' ** Update qryAssets_Currency_02_02_01 (Ledger, with curr_id_new, by specified [astno], [curridnew]).
22850       Set qdf = dbs.QueryDefs("qryAssets_Currency_02_02_02")
22860       With qdf.Parameters
22870         ![astno] = arr_varCurr(C_ASTNO, lngX)
22880         ![curridnew] = arr_varCurr(C_CIDN, lngX)
22890       End With
22900       qdf.Execute
22910       Set qdf = Nothing
22920       DoEvents

            ' ** Update qryAssets_Currency_02_03_01 (LedgerArchive, with curr_id_new, by specified [astno], [curridnew]).
22930       Set qdf = dbs.QueryDefs("qryAssets_Currency_02_03_02")
22940       With qdf.Parameters
22950         ![astno] = arr_varCurr(C_ASTNO, lngX)
22960         ![curridnew] = arr_varCurr(C_CIDN, lngX)
22970       End With
22980       qdf.Execute
22990       Set qdf = Nothing
23000       DoEvents

            ' ** Update qryAssets_Currency_02_04_01 (Journal, for as-is,
            ' ** with curr_id_new, by specified [astno], [curridnew]).
23010       Set qdf = dbs.QueryDefs("qryAssets_Currency_02_04_02")
23020       With qdf.Parameters
23030         ![astno] = arr_varCurr(C_ASTNO, lngX)
23040         ![curridnew] = arr_varCurr(C_CIDN, lngX)
23050       End With
23060       qdf.Execute
23070       Set qdf = Nothing
23080       DoEvents

            ' ** How does tblPricing_MasterAsset_History fit into all this?
            ' ** Any pricing prior to today should remain as they are; they remain accurate.
            ' ** If there's a price with today's date, perhaps that should be updated.

            ' ** qryAssets_Currency_02_05_01 (tblPricing_MasterAsset_History, for
            ' ** as-is, today's date, with curr_id_new, by specified [astno],
            ' ** [curridnew]), linked to tblCurrency, with .._new fields.
23090       Set qdf = dbs.QueryDefs("qryAssets_Currency_02_05_02")
23100       With qdf.Parameters
23110         ![astno] = arr_varCurr(C_ASTNO, lngX)
23120         ![curridnew] = arr_varCurr(C_CIDN, lngX)
23130       End With
23140       Set rst1 = qdf.OpenRecordset
23150       If rst1.BOF = True And rst1.EOF = True Then
              ' ** No applicable entries.
23160         rst1.Close
23170       Else
23180         Set rst2 = dbs.OpenRecordset("tblPricing_MasterAsset_History", dbOpenDynaset, dbConsistent)
23190         With rst1
23200           .MoveFirst  ' ** There can be only one.
23210           DoEvents
23220           With rst2
23230             .FindFirst "[AssetPricing_ID] = " & CStr(rst1![AssetPricing_ID])
23240             If .NoMatch = False Then
23250               .Edit
23260               ![curr_id] = rst1![curr_id_new]
23270               ![curr_code] = rst1![curr_code_new]
23280               ![currsym_symbol] = rst1![currsym_symbol_new]
23290               ![curr_date] = rst1![curr_date_new]
23300               ![curr_rate1] = rst1![curr_rate1_new]
23310               ![curr_rate2] = rst1![curr_rate2_new]
23320               ![CurrentJournalUser] = CurrentUser()  ' ** Internal Access Function: Trust Accountant login.
23330               ![DateModified] = Now()
23340               .Update
23350             End If
23360             .Close
23370           End With
23380           .Close
23390         End With
23400       End If
23410       Set rst1 = Nothing
23420       Set rst2 = Nothing
23430       Set qdf = Nothing
23440       DoEvents

23450     End Select  ' ** C_CONV.

23460   Next  ' ** lngX.

EXITP:
23470   Set rst1 = Nothing
23480   Set rst2 = Nothing
23490   Set qdf = Nothing
23500   Exit Sub

ERRH:
23510   DoCmd.Hourglass False
23520   THAT_PROC = THIS_PROC
23530   That_Erl = Erl
23540   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
23550   Resume EXITP

End Sub

Public Sub RecalcNewAssets()

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcNewAssets"

        Dim varTmp00 As Variant

23610   With Me

23620     DoCmd.Hourglass True
23630     DoEvents

23640     varTmp00 = DCount("*", "masterasset")
23650     lngRecsCur = DCount("*", "tmpEdit09")

23660     .NewAssets = (lngRecsCur - CLng(varTmp00))

23670     If .NewAssets > 0 And .NewAssets.Visible = False Then
23680       .NewAssets.Visible = True
23690     ElseIf .NewAssets = 0 And .NewAssets.Visible = True Then
23700       .NewAssets.Visible = False
23710     End If

23720     DoCmd.Hourglass False

23730   End With

EXITP:
23740   Exit Sub

ERRH:
23750   DoCmd.Hourglass False
23760   Select Case ERR.Number
        Case Else
23770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23780   End Select
23790   Resume EXITP

End Sub

Public Sub RecalcForEx()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcForEx"

        Dim varTmp00 As Variant

23810   With Me

23820     DoCmd.Hourglass True
23830     DoEvents

          ' ** MasterAsset, grouped, just curr_id <> 150, with cnt.
23840     varTmp00 = DLookup("[cnt]", "qryCurrency_08")

23850     Select Case IsNull(varTmp00)
          Case True
23860       .ForExCnt1 = 0
23870     Case False
23880       .ForExCnt1 = varTmp00
23890     End Select

          ' ** tmpEdit09, grouped, just curr_id <> 150, with cnt.
23900     varTmp00 = DLookup("[cnt]", "qryCurrency_15")

23910     Select Case IsNull(varTmp00)
          Case True
23920       .ForExCnt2 = 0
23930     Case False
23940       .ForExCnt2 = varTmp00
23950     End Select

23960     DoCmd.Hourglass False

23970   End With

EXITP:
23980   Exit Sub

ERRH:
23990   DoCmd.Hourglass False
24000   Select Case ERR.Number
        Case Else
24010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24020   End Select
24030   Resume EXITP

End Sub

Public Sub RecalcTots()

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

24110   With Me
24120     blnRecalcTots = True
24130     .TimerInterval = 250&
24140   End With

EXITP:
24150   Exit Sub

ERRH:
24160   Select Case ERR.Number
        Case Else
24170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24180   End Select
24190   Resume EXITP

End Sub
