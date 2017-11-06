VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2_One"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2_One"

'VGC 10/19/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   cmdHelp_Click()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Show All:         A {chkShowAll}
' **   Close:            C {cmdClose}
' **   List Hidden:      D {ledger_HIDDEN on frmAccountHideTrans2_One_Sub_List}
' **   Cancel:           E {cmdCancel}
' **   From:             F {DateStart}
' **   Include Archive:  H {chkIncludeArchive}
' **   CUSIP:            I {opgAssetSource_optCusip}
' **   Show JournalNo:   J {chkShowJournalNo}
' **   Journal Type:     L {cmbJournalTypes}
' **   Name:             M {opgAssetSource_optName}
' **   Pick Hidden:      N {ledger_HIDDEN on frmAccountHideTrans2_One_Sub_Pick}
' **   OK:               O {cmdOK}
' **   Clear All:        R {cmdClearAll}
' **   Asset:            S {cmbAssets}
' **   To:               T {DateEnd}
' **   Include Currency: U {chkIncludeCurrency}
' **   Exit:             X {cmdClose}
' **   Type:             Y {opgAssetSource_optType}

' ** Shortcut Ctrl keys responsive from this form:
' **   Help:             H {cmdHelp}
' **   Save:             S {cmdSave on frmAccountHideTrans2_One_Sub_List, frmAccountHideTrans2_One_Sub_Pick}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnClearAll_Focus As Boolean, blnClearAll_MouseDown As Boolean
Private blnHelp_Focus As Boolean, blnHelp_MouseDown As Boolean

Private lngICashSub_Left As Long, lngPCashSub_Left As Long, lngCostSub_Left As Long, lngSubTotLbl_Offset As Long
Private lngOK_Left As Long, lngCancel_Left As Long, lngPTrans_Left As Long, lngPTransLbl_Offset As Long
Private lngSharefaceSub_Left As Long, lngCurrID_Left As Long

Private lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngDetail_Height As Long, lngSubHead_Diff As Long

' ** Array: arr_varItem().
Private lngItems As Long, arr_varItem() As Variant
Private Const I_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const I_CNAM As Integer = 0
Private Const I_VIS  As Integer = 1

Private strAccountNo As String, strAccountName As String, blnHasArchive As Boolean, blnNoTrans As Boolean
Private blnHasCurrent As Boolean, blnHasCurrentMsgShown As Boolean, datLastBalDate As Date
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String, strCallersCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnHasArchs As Boolean
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then
150           If CharCnt(strTmp01, "~") = 5 Then  ' ** Module Function: modStringFuncs.

160             blnIsOpen = True

170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             strCallersCallingForm = Left(strTmp01, (intPos01 - 1))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             strAccountNo = Left(strTmp01, (intPos01 - 1))  ' ** gstrAccountNo should also already have this.
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             strAccountName = Left(strTmp01, (intPos01 - 1))  ' ** gstrAccountName should also already have this.
280             strTmp01 = Mid(strTmp01, (intPos01 + 1))
290             intPos01 = InStr(strTmp01, "~")
300             blnHasCurrent = CBool(Left(strTmp01, (intPos01 - 1)))  ' ** Has transactions in current statement period, after last statement.
310             strTmp01 = Mid(strTmp01, (intPos01 + 1))
320             If strTmp01 = "0" Then
330               datLastBalDate = 0#
340             Else
350               datLastBalDate = CDate(strTmp01)
360             End If

370             .LastStmt_Trans_lbl.Visible = False
380             If blnHasCurrent = False Then
390               ShowSub False
400             End If

410             gstrFormQuerySpec = THIS_NAME

420             blnClearAll_Focus = False: blnClearAll_MouseDown = False
430             blnHelp_Focus = False: blnHelp_MouseDown = False
440             gblnHasForExThis = False: blnHasArchive = False: blnHasCurrentMsgShown = False: blnNoTrans = False

450             CLR_DISABLED_FG = CLR_DKGRY
460             CLR_DISABLED_BG = CLR_LTTEAL

470             .accountno = strAccountNo
480             .shortname = strAccountName
490             .balance_date_max = datLastBalDate

                ' ** Pixels to move up.
500             lngSubHead_Diff = (.frmAccountHideTrans2_One_Sub_List_box.Top - .Subheader_closed_line.Top)

510             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
520             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

530             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

540             .cmdClearAll.Enabled = False
550             .cmdClearAll_raised_img_dis.Visible = True
560             .cmdClearAll_raised_img.Visible = False
570             .cmdClearAll_raised_semifocus_dots_img.Visible = False
580             .cmdClearAll_raised_focus_img.Visible = False
590             .cmdClearAll_raised_focus_dots_img.Visible = False
600             .cmdClearAll_sunken_focus_dots_img.Visible = False

610             lngSharefaceSub_Left = .shareface_subtot.Left
620             lngICashSub_Left = .icash_subtot.Left
630             lngPCashSub_Left = .pcash_subtot.Left
640             lngCostSub_Left = .cost_subtot.Left
650             lngSubTotLbl_Offset = .shareface_subtot.Left - .subtot_lbl.Left
660             lngOK_Left = .cmdOK.Left
670             lngCancel_Left = .cmdCancel.Left
680             lngPTrans_Left = .pending_trans.Left
690             lngPTransLbl_Offset = (.pending_trans.Left - .pending_trans_lbl.Left)
700             lngCurrID_Left = .curr_id.Left

710             .cmdClearAll_lbl.Top = 0&  ' ** Just get these out of the way.
720             .cmdHelp_lbl.Top = 0&
730             .detail_height_line.Top = 0&

740             lngDetail_Height = .Detail.Height

750             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
760             lngMonitorNum = 1&: lngTmp02 = 0&
770             EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
780             If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
                ' ** Variables are fed empty, then populated ByRef.
790             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.
800             lngTmp02 = lngFrm_Top  'THIS IS SUPPOSED TO BE GETTING A DIFFERENT VALUE FOR >1 MONITOR!
810             If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
820             DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, lngFrm_Height  'lngFrm_Top
830             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.
840             If lngMonitorNum > 1& Then
850               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
860             End If

870             varTmp00 = DCount("*", "LedgerArchive")
880             If IsNull(varTmp00) = False Then
890               If varTmp00 > 0 Then
900                 blnHasArchive = True
910               End If
920             End If

930             .frmAccountHideTrans2_One_Sub_List.Form.Filter = vbNullString
940             .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False

950             If .cmdHelp_IsClosed_chk = True Then
960               cmdHelp_Click  ' ** Procedure: Below.
970             End If

980             lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
990             If lngRecsCur > 0& Then
1000              blnHasArchs = .frmAccountHideTrans2_One_Sub_List.Form.HasArch
1010            End If
1020            ArchiveSet blnHasArchs  ' ** Procedure: Below.

1030            chkShowJournalNo_AfterUpdate  ' ** Procedure: Below.
1040            chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1050            chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
1060            opgAssetSource_AfterUpdate  ' ** Procedure: Below.

1070            .frmAccountHideTrans2_One_Sub_List.SetFocus

1080            blnIsOpen = False

1090          Else
1100            Cancel = -1
1110          End If
1120        Else
1130          Cancel = -1
1140        End If
1150      Else
1160        Cancel = -1
1170      End If
1180    End With

1190    DoCmd.Hourglass False

1200    If Cancel = -1 Then
1210      DoCmd.Hourglass False
1220      Beep
1230      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1240    End If

EXITP:
1250    Exit Sub

ERRH:
1260    DoCmd.Hourglass False
1270    Select Case ERR.Number
        Case Else
1280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1290    End Select
1300    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1410    intShiftDown = (Shift And acShiftMask) > 0
1420    intAltDown = (Shift And acAltMask) > 0
1430    intCtrlDown = (Shift And acCtrlMask) > 0

1440    intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   List Hidden:      D {ledger_HIDDEN on frmAccountHideTrans2_One_Sub_List}
        ' **   Pick Hidden:      N {ledger_HIDDEN on frmAccountHideTrans2_One_Sub_Pick}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Help:             H {cmdHelp}
        ' **   Save:             S {cmdSave on frmAccountHideTrans2_One_Sub_List, frmAccountHideTrans2_One_Sub_Pick}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        intRetVal = 0
1480        cmdClose_Click  ' ** Procedure: Below.
1490      End Select
1500    End If

        ' ** Alt keys.
1510    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1520      Select Case intRetVal
          Case vbKeyD
1530        With Me
1540          intRetVal = 0
1550          lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
1560          If lngRecsCur > 0& Then
1570            .frmAccountHideTrans2_One_Sub_List.SetFocus
1580            .frmAccountHideTrans2_One_Sub_List.Form.ledger_HIDDEN.SetFocus  ' ** Don't invoke.
1590          Else
1600            Beep
1610          End If
1620        End With
1630      Case vbKeyN
1640        With Me
1650          intRetVal = 0
1660          lngRecsCur = .frmAccountHideTrans2_One_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_Pick.
1670          If lngRecsCur > 0& Then
1680            .frmAccountHideTrans2_One_Sub_Pick.SetFocus
1690            .frmAccountHideTrans2_One_Sub_Pick.Form.ledger_HIDDEN.SetFocus  ' ** Don't invoke.
1700          Else
1710            Beep
1720          End If
1730        End With
1740      Case vbKeyX
1750        intRetVal = 0
1760        cmdClose_Click  ' ** Procedure: Below.
1770      End Select
1780    End If

        ' ** Ctrl keys.
1790    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1800      Select Case intRetVal
          Case vbKeyH
1810        With Me
1820          intRetVal = 0
1830          .cmdHelp.SetFocus
1840          cmdHelp_Click  ' ** Procedure: Below.
1850        End With
1860      Case vbKeyS
1870        With Me
1880          intRetVal = 0
1890          .frmAccountHideTrans2_One_Sub_List.Form.cmdSave_Click  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_List.
1900          .frmAccountHideTrans2_One_Sub_Pick.Form.cmdSave_Click  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.
1910        End With
1920      End Select
1930    End If

        ' ** Ctrl-Shift keys.
1940    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1950      Select Case intRetVal
          Case vbKeyF
1960        With Me
1970          intRetVal = 0
1980          .FocusHolder.SetFocus
1990        End With
2000      End Select
2010    End If

EXITP:
2020    KeyCode = intRetVal
2030    Exit Sub

ERRH:
2040    intRetVal = 0
2050    Select Case ERR.Number
        Case Else
2060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2070    End Select
2080    Resume EXITP

End Sub

Private Sub Form_Timer()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim msgResponse As VbMsgBoxResult

2110    With Me
2120      .TimerInterval = 0&
2130      If gblnSetFocus = True Then
2140        gblnSetFocus = False
2150        strFormName = vbNullString
2160  On Error Resume Next
2170        Do Until strFormName = THIS_NAME
2180          DoCmd.SelectObject acForm, THIS_NAME, False
2190          strFormName = Screen.ActiveForm.Name
2200        Loop
2210  On Error GoTo ERRH
2220      End If
2230      Select Case blnHasCurrent
          Case True
2240        .chkShowAll = False
2250        chkShowAll_AfterUpdate  ' ** Procedure: Below.
2260      Case False
2270        If blnHasCurrentMsgShown = False Then
2280          blnHasCurrentMsgShown = True
2290          Beep
2300          DoCmd.Hourglass False
2310          msgResponse = MsgBox("There are no transactions since the last Statement." & vbCrLf & vbCrLf & _
                "Whould you like to see all the transactions?", vbQuestion + vbOKCancel, "No Current Transactions")
2320          DoCmd.Hourglass True
2330          DoEvents
2340          Select Case msgResponse
              Case vbOK
2350            .chkShowAll = True
2360          Case Else
2370            blnNoTrans = True
2380            .chkShowAll = False
2390          End Select
2400          ShowSub True
2410          chkShowAll_AfterUpdate  ' ** Procedure: Below.
2420          DoEvents
2430        End If
2440      End Select
2450      gblnHasForExThis = HasForEx_Acct(gstrAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.
2460      Select Case gblnHasForEx
          Case True
2470        .chkIncludeCurrency.Enabled = True
2480        Select Case gblnHasForExThis
            Case True
2490          .chkIncludeCurrency = True
2500          .chkIncludeCurrency.Locked = True
2510          chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
2520        Case False
2530          .chkIncludeCurrency.Locked = False
2540        End Select
2550      Case False
2560        .chkIncludeCurrency = False
2570        chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
2580        .chkIncludeCurrency.Enabled = False
2590      End Select
2600      DoCmd.Hourglass False
2610    End With

EXITP:
2620    Exit Sub

ERRH:
2630    DoCmd.Hourglass False
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strControl As String
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

2710    With Me
2720      If DataErr <> 0 Then
2730        Select Case DataErr
            Case 2113, 2279
              ' ** 2113: The value you entered isn't valid for this field.
              ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2740          strControl = Screen.ActiveControl.Name
2750          Select Case strControl
              Case "DateStart", "DateEnd"
2760            Beep
2770            MsgBox "Please enter a valid date." & vbCrLf & vbCrLf & vbCrLf & "msg 1", vbInformation + vbOKOnly, "Invalid Date"
2780          End Select
2790          Select Case strControl
              Case "DateStart"
2800            .DateStart.Undo
2810          Case "DateEnd"
2820            .DateEnd.Undo
2830          End Select
2840          Response = acDataErrContinue
2850        Case Else
2860          varTmp00 = DLookup("[ErrorString]", "tblAccessAndJetErrors", "[ErrorCode] = " & CStr(DataErr))
2870          If IsNull(varTmp00) = False Then
2880            Beep
2890            MsgBox "Error: " & CStr(DataErr) & vbCrLf & "Description: " & varTmp00 & vbCrLf & _
                  "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & _
                  "Line: ", vbInformation + vbOKOnly, "Error: " & CStr(DataErr)
2900            Set dbs = CurrentDb
2910            Set rst = dbs.OpenRecordset("tblErrorLog", dbOpenDynaset, dbAppendOnly)
2920            blnRetVal = zErrorWriteRecord(CDbl(DataErr), CStr(varTmp00), THIS_NAME, THIS_PROC, 0, rst)  ' ** Module Function: modErrorHandler.
2930            rst.Close
2940            dbs.Close
2950            Set rst = Nothing
2960            Set dbs = Nothing
2970            Response = acDataErrContinue
2980          Else
2990            Response = acDataErrDisplay
3000          End If
3010        End Select
3020      End If  ' ** DataErr.
3030    End With

        ' ** AcDataError enumeration:
        ' **   0  acDataErrContinue  Ignore the error and continue without displaying the default Microsoft Access
        ' **                         error message. A custom error message may be displayed in place of the default
        ' **                         error message.
        ' **   1  acDataErrDisplay   Display the default Microsoft Access error message. (Default)
        ' **   2  acDataErrAdded     Don't display the default Microsoft Access error message. The entry may be
        ' **                         added to the combo box list in the NotInList event procedure. After the entry
        ' **                         is added, Microsoft Access updates the list by requerying the combo box.
        ' **                         Microsoft Access then rechecks the string against the combo box list, and saves
        ' **                         the value in the NewData argument in the field the combo box is bound to. If
        ' **                         the string is not in the list, then Microsoft Access displays an error message.

EXITP:
3040    Set rst = Nothing
3050    Set dbs = Nothing
3060    Exit Sub

ERRH:
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult

3210    With Me

3220      DoCmd.Hourglass True
3230      DoEvents

3240      If .cmdOK.Enabled = True Then
3250        DoCmd.Hourglass False
3260        Beep
3270        msgResponse = MsgBox("You have chosen a transaction to reverse, but have not" & vbCrLf & _
              "committed your choice by clicking the 'OK' button." & vbCrLf & vbCrLf & _
              "Are you sure you want to abandon this action?", vbQuestion + vbYesNo, "Unposted Reversal")
3280      Else
3290        msgResponse = vbYes
3300      End If

3310      If msgResponse <> vbYes Then
3320        Cancel = -1
3330      Else

3340        DoCmd.Hourglass True
3350        DoEvents

3360        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3370        Set dbs = CurrentDb
3380        With dbs
              ' ** Empty tblLedgerHidden_Staging1.
3390          Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_List_04")
3400          qdf.Execute
3410          Set qdf = Nothing
              ' ** Empty tblLedgerHidden_Staging2.
3420          Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
3430          qdf.Execute
3440          Set qdf = Nothing
3450          .Close
3460        End With
3470        Set dbs = Nothing

            ' ** Make sure all reports are closed.
3480        If Reports.Count > 0 Then
3490          Do While Reports.Count > 0
3500            DoCmd.Close acReport, Reports(0).Name
3510          Loop
3520        End If

3530        gblnSetFocus = True
3540        gstrReturningForm = THIS_NAME
3550        Select Case strCallersCallingForm
            Case "frmMenu_Account"
3560          DoCmd.OpenForm strCallersCallingForm, , , , , , "frmMenu_Main"
3570        Case Else
              ' ** Nothing at the moment.
3580        End Select

3590      End If  ' ** msgResponse.

3600    End With

EXITP:
3610    Set qdf = Nothing
3620    Set dbs = Nothing
3630    Exit Sub

ERRH:
3640    Select Case ERR.Number
        Case Else
3650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3660    End Select
3670    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3710    With Me
3720      If .cmdClearAll_raised_focus_img.Visible = True Or .cmdClearAll_raised_focus_dots_img.Visible = True Then
3730        Select Case blnClearAll_Focus
            Case True
3740          .cmdClearAll_raised_semifocus_dots_img.Visible = True
3750          .cmdClearAll_raised_img.Visible = False
3760        Case False
3770          .cmdClearAll_raised_img.Visible = True
3780          .cmdClearAll_raised_semifocus_dots_img.Visible = False
3790        End Select
3800        .cmdClearAll_raised_focus_img.Visible = False
3810        .cmdClearAll_raised_focus_dots_img.Visible = False
3820        .cmdClearAll_sunken_focus_dots_img.Visible = False
3830        .cmdClearAll_raised_img_dis.Visible = False
3840      End If
3850      If cmdHelp_L_raised_focus_img.Visible = True Or .cmdHelp_L_raised_focus_dots_img.Visible = True Or _
              .cmdHelp_R_raised_focus_img.Visible = True Or .cmdHelp_R_raised_focus_dots_img.Visible = True Then
3860        Select Case .Subheader_box.Visible
            Case True
3870          Select Case blnHelp_Focus
              Case True
3880            .cmdHelp_L_raised_semifocus_dots_img.Visible = True
3890            .cmdHelp_L_raised_img.Visible = False
3900          Case False
3910            .cmdHelp_L_raised_img.Visible = True
3920            .cmdHelp_L_raised_semifocus_dots_img.Visible = False
3930          End Select
3940          .cmdHelp_R_raised_img.Visible = False
3950          .cmdHelp_R_raised_semifocus_dots_img.Visible = False
3960        Case False
3970          Select Case blnHelp_Focus
              Case True
3980            .cmdHelp_R_raised_semifocus_dots_img.Visible = True
3990            .cmdHelp_R_raised_img.Visible = False
4000          Case False
4010            .cmdHelp_R_raised_img.Visible = True
4020            .cmdHelp_R_raised_semifocus_dots_img.Visible = False
4030          End Select
4040          .cmdHelp_L_raised_img.Visible = False
4050          .cmdHelp_L_raised_semifocus_dots_img.Visible = False
4060        End Select
4070        .cmdHelp_L_raised_focus_img.Visible = False
4080        .cmdHelp_L_raised_focus_dots_img.Visible = False
4090        .cmdHelp_L_sunken_focus_dots_img.Visible = False
4100        .cmdHelp_L_raised_img_dis.Visible = False
4110        .cmdHelp_R_raised_focus_img.Visible = False
4120        .cmdHelp_R_raised_focus_dots_img.Visible = False
4130        .cmdHelp_R_sunken_focus_dots_img.Visible = False
4140        .cmdHelp_R_raised_img_dis.Visible = False
4150      End If
4160    End With

EXITP:
4170    Exit Sub

ERRH:
4180    Select Case ERR.Number
        Case Else
4190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4200    End Select
4210    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

4310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Public Sub cmdClose_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim frmList As Access.Form, frmPick As Access.Form

4510    With Me

4520      .TimerInterval = 0&

4530      DoCmd.Hourglass True
4540      DoEvents

4550      Set frmList = .frmAccountHideTrans2_One_Sub_List.Form
4560      Set frmPick = .frmAccountHideTrans2_One_Sub_Pick.Form

          ' ** So they don't see the '#DELETED#' fields.
4570      With frmList
4580        .journalno.ForeColor = CLR_DISABLED_BG
4590        .transdate.ForeColor = CLR_DISABLED_BG
4600        .journaltype.ForeColor = CLR_DISABLED_BG
4610        .shareface_0.ForeColor = CLR_DISABLED_BG
4620        .shareface_1.ForeColor = CLR_DISABLED_BG
4630        .shareface_2.ForeColor = CLR_DISABLED_BG
4640        .shareface_3.ForeColor = CLR_DISABLED_BG
4650        .shareface_4.ForeColor = CLR_DISABLED_BG
4660        .totdesc.ForeColor = CLR_DISABLED_BG
4670        .icash_str.ForeColor = CLR_DISABLED_BG
4680        .ICash.ForeColor = CLR_DISABLED_BG
4690        .pcash_str.ForeColor = CLR_DISABLED_BG
4700        .PCash.ForeColor = CLR_DISABLED_BG
4710        .cost_str.ForeColor = CLR_DISABLED_BG
4720        .Cost.ForeColor = CLR_DISABLED_BG
4730        .curr_id.ForeColor = CLR_DISABLED_BG
4740        lngRecsCur = .RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
4750        If lngRecsCur > 0& Then
4760          .FocusHolder.SetFocus
4770        End If
4780        .ledger_HIDDEN_box.ForeColor = MY_CLR_BGE
4790        .ledger_HIDDEN_dummy.Visible = True
4800        .ledger_HIDDEN.Visible = False
4810      End With  ' ** frmList.

4820      With frmPick
4830        .journalno.ForeColor = CLR_DISABLED_BG
4840        .transdate.ForeColor = CLR_DISABLED_BG
4850        .journaltype.ForeColor = CLR_DISABLED_BG
4860        .shareface_0.ForeColor = CLR_DISABLED_BG
4870        .shareface_1.ForeColor = CLR_DISABLED_BG
4880        .shareface_2.ForeColor = CLR_DISABLED_BG
4890        .shareface_3.ForeColor = CLR_DISABLED_BG
4900        .shareface_4.ForeColor = CLR_DISABLED_BG
4910        .totdesc.ForeColor = CLR_DISABLED_BG
4920        .icash_str.ForeColor = CLR_DISABLED_BG
4930        .pcash_str.ForeColor = CLR_DISABLED_BG
4940        .cost_str.ForeColor = CLR_DISABLED_BG
4950        .curr_id.ForeColor = CLR_DISABLED_BG
4960        .ledger_HIDDEN_box.ForeColor = MY_CLR_BGE
4970      End With  ' ** frmPick.

4980  On Error Resume Next
4990      DoCmd.Close acForm, THIS_NAME
5000  On Error GoTo ERRH

5010    End With

EXITP:
5020    Set frmList = Nothing
5030    Set frmPick = Nothing
5040    Exit Sub

ERRH:
5050    DoCmd.Hourglass False
5060    Select Case ERR.Number
        Case Else
5070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5080    End Select
5090    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
5110    intShiftDown = (Shift And acShiftMask) > 0
5120    intAltDown = (Shift And acAltMask) > 0
5130    intCtrlDown = (Shift And acCtrlMask) > 0

5140    intRetVal = KeyCode

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab
5170        With Me
5180          intRetVal = 0
5190          lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
5200          If lngRecsCur > 0& Then
5210            .frmAccountHideTrans2_One_Sub_List.SetFocus
5220            .frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
5230          Else
5240            If .chkIncludeArchive.Enabled = True Then
5250              .chkIncludeArchive.SetFocus
5260            ElseIf .chkIncludeCurrency.Enabled = True Then
5270              .chkIncludeCurrency.SetFocus
5280            Else
5290              .chkShowAll.SetFocus
5300            End If
5310          End If
5320        End With
5330      End Select
5340    End If

        ' ** Shift keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5360      Select Case intRetVal
          Case vbKeyTab
5370        With Me
5380          intRetVal = 0
5390          If .cmdClearAll.Enabled = True Then
5400            .cmdClearAll.SetFocus
5410          Else
5420            .cmbAssets.SetFocus
5430          End If
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

Private Sub cmdOK_Click()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngJournalNo As Long, lngAssetNo As Long, lngCurrID As Long
        Dim strSource As String, strJType As String
        Dim datTransDate As Date
        Dim varAssetDate1 As Variant, varAssetDate2 As Variant, varPurchaseDate1 As Variant, varPurchaseDate2 As Variant
        Dim dblShareface As Double, dblICash As Double, dblPCash As Double, dblCost As Double
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean
        Dim varTmp00 As Variant, datTmp01 As Date

        'MAKE SURE PERSHARE HAS THE RIGHT VALUE!
        'IN ACTIVE ASSETS, COST_USD AND MARKET_USD ARE EMPTY!

5610    With Me

5620      blnContinue = True

5630      DoCmd.Hourglass True
5640      DoEvents

          ' ** Get some info about the reversing transaction.
5650      Set rst = .frmAccountHideTrans2_One_Sub_Pick.Form.RecordsetClone
5660      rst.MoveFirst
5670      rst.MoveNext
5680      strJType = rst![journaltype]    ' ** This is the new reversing entry.
5690      lngAssetNo = rst![assetno]      ' ** May be Zero.
5700      rst.Close
5710      Set rst = Nothing

          ' ** If it's reversing a 'Sold' or 'Withdrawn' with a 'Purchase' or 'Deposit',
          ' ** the asset must be priced. If it's not, advise and prevent continuing.
5720      If (strJType = "Deposit" Or strJType = "Purchase") And lngAssetNo > 0& Then
5730        Set dbs = CurrentDb
            ' ** MasterAsset, by specified [astno].
5740        Set qdf = dbs.QueryDefs("qryAccountHideTrans2_One_06_12")
5750        With qdf.Parameters
5760          ![astno] = lngAssetNo
5770        End With
5780        Set rst = qdf.OpenRecordset
5790        rst.MoveFirst
5800        varTmp00 = rst![marketvaluecurrent]
5810        rst.Close
5820        Set rst = Nothing
5830        dbs.Close
5840        Set dbs = Nothing
5850        If varTmp00 = 0 Then
5860          blnContinue = False
5870          Beep
5880          DoCmd.Hourglass False
5890          MsgBox "This asset has not been priced," & vbCrLf & _
                "and the procedure cannot continue." & vbCrLf & vbCrLf & _
                "Go to the Asset Pricing form found on the Asset Menu" & vbCrLf & _
                "to price this asset before trying the reversal again.", vbExclamation + vbOKOnly, "Asset Not Priced"
5900          .cmdCancel.SetFocus
5910          DoEvents
5920          .cmdOK.Enabled = False
5930        End If
5940      End If

5950      If blnContinue = True Then

5960        Beep
5970        DoCmd.Hourglass False
5980        msgResponse = MsgBox("This will create, post, and hide a reversing transaction" & vbCrLf & _
              "for the selected entry." & vbCrLf & vbCrLf & "Do you wish to continue?", _
              vbQuestion + vbYesNo, "Create Reversing Transaction")

5990        If msgResponse = vbYes Then

6000          DoCmd.Hourglass True
6010          DoEvents

6020          Set rst = .frmAccountHideTrans2_One_Sub_Pick.Form.RecordsetClone
6030          rst.MoveFirst
              ' ** strAccountNo has accountno.
6040          lngJournalNo = rst![journalno]  ' ** This is the existing entry.
6050          lngAssetNo = rst![assetno]      ' ** May be Zero.
6060          datTransDate = rst![transdate]
6070          varAssetDate1 = rst![assetdate]
6080          varPurchaseDate1 = rst![PurchaseDate]
6090          lngCurrID = rst![curr_id]
6100          strSource = rst![src]           ' ** 'L' = Ledger, 'A' = LedgerArchive.
6110          rst.MoveNext
6120          strJType = rst![journaltype]    ' ** This is the new reversing entry.
6130          dblShareface = rst![shareface]
6140          dblICash = rst![ICash]
6150          dblPCash = rst![PCash]
6160          dblCost = rst![Cost]
6170          varAssetDate2 = rst![assetdate]
6180          varPurchaseDate2 = rst![PurchaseDate]
6190          rst.Close
6200          Set rst = Nothing
6210          DoEvents

6220          Set dbs = CurrentDb
6230          With dbs

                ' ** Update the original ledger_HIDDEN.
6240            Select Case strSource
                Case "L"
                  ' ** Update Ledger, for ledger_HIDDEN = True, by specified [jno].
6250              Set qdf = .QueryDefs("qryAccountHideTrans2_One_04_01")
6260            Case "A"
                  ' ** Update LedgerArchive, for ledger_HIDDEN = True, by specified [jno].
6270              Set qdf = .QueryDefs("qryAccountHideTrans2_One_04_02")
6280            End Select
6290            With qdf.Parameters
6300              ![jno] = lngJournalNo
6310            End With
6320            qdf.Execute
6330            Set qdf = Nothing
6340            DoEvents

                ' ** Append the new reversing entry.
6350            Select Case strSource
                Case "L"
                  ' ** Append qryAccountHideTrans2_One_05_01 (tblLedgerHidden_Staging2,
                  ' ** just sort = 2, src = 'L') to Ledger.
6360              Set qdf = .QueryDefs("qryAccountHideTrans2_One_05_03")
6370            Case "A"
                  ' ** Append qryAccountHideTrans2_One_05_02 (tblLedgerHidden_Staging2,
                  ' ** just sort = 2, src = 'A') to LedgerArchive.
6380              Set qdf = .QueryDefs("qryAccountHideTrans2_One_05_04")
6390            End Select
6400            qdf.Execute dbFailOnError
6410            Set qdf = Nothing
6420            DoEvents

                ' ** Append the new Tax Lot or delete the reversed one.
6430            Select Case strJType
                Case "Purchase", "Deposit"
6440              datTmp01 = CDate(Format(varAssetDate2, "mm/dd/yyyy"))
6450              Select Case lngCurrID
                  Case 150&
                    'IF THE ASSET HAS NO PRICE, IT WILL FAIL THIS APPEND!
                    ' ** Append qryAccountHideTrans2_One_06_05 (qryAccountHideTrans2_One_06_01
                    ' ** (tblLedgerHidden_Staging2, just sort = 2, as new ActiveAssets record,
                    ' ** with cost_usd for curr_id = 150), linked to qryAccountHideTrans2_One_06_04
                    ' ** (qryAccountHideTrans2_One_06_03 (qryAccountHideTrans2_One_06_02
                    ' ** (tblPricing_MasterAsset_History, by specified [astno], [adat]), grouped,
                    ' ** with Max(currentDate)), linked back to qryAccountHideTrans2_One_06_02
                    ' ** (tblPricing_MasterAsset_History, by specified [astno], [adat])),
                    ' ** with market_usd for curr_id = 150) to ActiveAssets.
6460                Set qdf = .QueryDefs("qryAccountHideTrans2_One_06_10")
6470                With qdf.Parameters
6480                  ![astno] = lngAssetNo
6490                  ![adat] = datTmp01
6500                End With
6510              Case Else
                    ' ** Append qryAccountHideTrans2_One_06_09 (qryAccountHideTrans2_One_06_05
                    ' ** (qryAccountHideTrans2_One_06_01 (tblLedgerHidden_Staging2, just sort = 2,
                    ' ** as new ActiveAssets record, with cost_usd for curr_id = 150), linked to
                    ' ** qryAccountHideTrans2_One_06_04 (qryAccountHideTrans2_One_06_03
                    ' ** (qryAccountHideTrans2_One_06_02 (tblPricing_MasterAsset_History, by
                    ' ** specified [astno], [adat]), grouped, with Max(currentDate)), linked back
                    ' ** to qryAccountHideTrans2_One_06_02 (tblPricing_MasterAsset_History, by
                    ' ** specified [astno], [adat])), with market_usd for curr_id = 150), linked
                    ' ** to qryAccountHideTrans2_One_06_08 (qryAccountHideTrans2_One_06_07
                    ' ** (qryAccountHideTrans2_One_06_06 (tblCurrency_History, by specified
                    ' ** [currid], [adat]), grouped, with Max(curr_date])), linked back to
                    ' ** qryAccountHideTrans2_One_06_06 (tblCurrency_History, by specified [currid],
                    ' ** [adat])), with cost_usd, market_usd for curr_id <> 150) to ActiveAssets.
6520                Set qdf = .QueryDefs("qryAccountHideTrans2_One_06_11")
6530                With qdf.Parameters
6540                  ![astno] = lngAssetNo
6550                  ![adat] = datTmp01
6560                  ![currid] = lngCurrID
6570                End With
6580              End Select
6590              qdf.Execute dbFailOnError
6600              Set qdf = Nothing
6610            Case "Sold", "Withdrawn"
                  ' ** Delete ActiveAssets, by specified [astno], [actno], [adat].
6620              Set qdf = .QueryDefs("qryAccountHideTrans2_One_07")
6630              With qdf.Parameters
6640                ![astno] = lngAssetNo
6650                ![actno] = strAccountNo
6660                ![adat] = CDate(varAssetDate1)
6670              End With
6680              qdf.Execute dbFailOnError
6690              Set qdf = Nothing
6700            Case "Liability"
6710              Select Case IsNull(varPurchaseDate2)
                  Case True
                    ' ** Liability (+).
6720                datTmp01 = CDate(Format(varAssetDate2, "mm/dd/yyyy"))
6730                Select Case lngCurrID
                    Case 150&
                      ' ** Append qryAccountHideTrans2_One_06_05 (qryAccountHideTrans2_One_06_01
                      ' ** (tblLedgerHidden_Staging2, just sort = 2, as new ActiveAssets record,
                      ' ** with cost_usd for curr_id = 150), linked to qryAccountHideTrans2_One_06_04
                      ' ** (qryAccountHideTrans2_One_06_03 (qryAccountHideTrans2_One_06_02
                      ' ** (tblPricing_MasterAsset_History, by specified [astno], [adat]), grouped,
                      ' ** with Max(currentDate)), linked back to qryAccountHideTrans2_One_06_02
                      ' ** (tblPricing_MasterAsset_History, by specified [astno], [adat])),
                      ' ** with market_usd for curr_id = 150) to ActiveAssets.
6740                  Set qdf = .QueryDefs("qryAccountHideTrans2_One_06_10")
6750                  With qdf.Parameters
6760                    ![astno] = lngAssetNo
6770                    ![adat] = datTmp01
6780                  End With
6790                Case Else
                      ' ** Append qryAccountHideTrans2_One_06_09 (qryAccountHideTrans2_One_06_05
                      ' ** (qryAccountHideTrans2_One_06_01 (tblLedgerHidden_Staging2, just sort = 2,
                      ' ** as new ActiveAssets record, with cost_usd for curr_id = 150), linked to
                      ' ** qryAccountHideTrans2_One_06_04 (qryAccountHideTrans2_One_06_03
                      ' ** (qryAccountHideTrans2_One_06_02 (tblPricing_MasterAsset_History, by
                      ' ** specified [astno], [adat]), grouped, with Max(currentDate)), linked back
                      ' ** to qryAccountHideTrans2_One_06_02 (tblPricing_MasterAsset_History, by
                      ' ** specified [astno], [adat])), with market_usd for curr_id = 150), linked
                      ' ** to qryAccountHideTrans2_One_06_08 (qryAccountHideTrans2_One_06_07
                      ' ** (qryAccountHideTrans2_One_06_06 (tblCurrency_History, by specified
                      ' ** [currid], [adat]), grouped, with Max(curr_date])), linked back to
                      ' ** qryAccountHideTrans2_One_06_06 (tblCurrency_History, by specified [currid],
                      ' ** [adat])), with cost_usd, market_usd for curr_id <> 150) to ActiveAssets.
6800                  Set qdf = .QueryDefs("qryAccountHideTrans2_One_06_11")
6810                  With qdf.Parameters
6820                    ![astno] = lngAssetNo
6830                    ![adat] = datTmp01
6840                    ![currid] = lngCurrID
6850                  End With
6860                End Select
6870                qdf.Execute dbFailOnError
6880                Set qdf = Nothing
6890              Case False
                    ' ** Liability (-).
                    ' ** Delete ActiveAssets, by specified [astno], [actno], [adat].
6900                Set qdf = .QueryDefs("qryAccountHideTrans2_One_07")
6910                With qdf.Parameters
6920                  ![astno] = lngAssetNo
6930                  ![actno] = strAccountNo
6940                  ![adat] = CDate(varAssetDate1)
6950                End With
6960                qdf.Execute dbFailOnError
6970                Set qdf = Nothing
6980              End Select
6990            End Select
7000            DoEvents

7010            If strJType = "Deposit" Or strJType = "Purchase" Or strJType = "Withdrawn" Or strJType = "Sold" Or strJType = "Liability" Then

                  ' ** Update shareface in MasterAsset.
7020              Select Case strJType
                  Case "Deposit", "Purchase"
                    ' ** Update qryAccountHideTrans2_One_08_01 (MasterAsset,
                    ' ** for shareface increase, by specified [astno], [shrs]).
7030                Set qdf = .QueryDefs("qryAccountHideTrans2_One_08_03")
7040              Case "Withdrawn", "Sold"
                    ' ** Update qryAccountHideTrans2_One_08_02 (MasterAsset,
                    ' ** for shareface decrease, by specified [astno], [shrs]).
7050                Set qdf = .QueryDefs("qryAccountHideTrans2_One_08_04")
7060              Case "Liability"
7070                Select Case IsNull(varPurchaseDate2)
                    Case True
                      ' ** Liability (+).
                      ' ** Update qryAccountHideTrans2_One_08_01 (MasterAsset,
                      ' ** for shareface increase, by specified [astno], [shrs]).
7080                  Set qdf = .QueryDefs("qryAccountHideTrans2_One_08_03")
7090                Case False
                      ' ** Liability (-).
                      ' ** Update qryAccountHideTrans2_One_08_02 (MasterAsset,
                      ' ** for shareface decrease, by specified [astno], [shrs]).
7100                  Set qdf = .QueryDefs("qryAccountHideTrans2_One_08_04")
7110                End Select
7120              End Select
7130              With qdf.Parameters
7140                ![astno] = lngAssetNo
7150                ![shrs] = dblShareface
7160              End With
7170              qdf.Execute dbFailOnError
7180              Set qdf = Nothing
7190              DoEvents

7200            End If

                ' ** Update master totals in Account.
7210            If lngCurrID <> 150& Then
                  ' ** qryAccountHideTrans2_One_09_04 (qryAccountHideTrans2_One_09_03
                  ' ** (tblCurrency_History, by specified [currid], [tdat]), grouped,
                  ' ** with Max(curr_date])), linked back to qryAccountHideTrans2_One_09_03
                  ' ** (tblCurrency_History, by specified [currid], [tdat]).
7220              Set qdf = .QueryDefs("qryAccountHideTrans2_One_09_05")
7230              With qdf.Parameters
7240                ![currid] = lngCurrID
7250                ![tdat] = datTransDate
7260              End With
7270              Set rst = qdf.OpenRecordset
7280              With rst
7290                .MoveFirst
7300                dblICash = Round((dblICash * ![curr_rate2]), 2)
7310                dblPCash = Round((dblPCash * ![curr_rate2]), 2)
7320                dblCost = Round((dblCost * ![curr_rate2]), 2)
7330                .Close
7340              End With
7350              Set rst = Nothing
7360              Set qdf = Nothing
7370            End If
                ' ** Update qryAccountHideTrans2_One_09_01 (Account,
                ' ** by specified [actno], [icsh], [pcsh], [cst]).
7380            Set qdf = .QueryDefs("qryAccountHideTrans2_One_09_02")
7390            With qdf.Parameters
7400              ![actno] = strAccountNo
7410              ![icsh] = dblICash
7420              ![pcsh] = dblPCash
7430              ![CST] = dblCost
7440            End With
7450            qdf.Execute dbFailOnError
7460            Set qdf = Nothing
7470            DoEvents

                ' ** Delete tblLedgerHidden_Staging1, by specified [jno].
7480            Set qdf = .QueryDefs("qryAccountHideTrans2_One_10")
7490            With qdf.Parameters
7500              ![jno] = lngJournalNo
7510            End With
7520            qdf.Execute
7530            Set qdf = Nothing
7540            DoEvents

                ' ** Empty tblLedgerHidden_Staging2.
7550            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
7560            qdf.Execute
7570            Set qdf = Nothing
7580            DoEvents

7590            .Close
7600          End With  ' ** dbs.
7610          Set dbs = Nothing

7620          .cmdClose.SetFocus
7630          DoEvents

7640          .frmAccountHideTrans2_One_Sub_List.Form.Requery
7650          .frmAccountHideTrans2_One_Sub_Pick.Form.Requery
7660          .frmAccountHideTrans2_One_Sub_Pick.Form.LabelEnable False  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.

7670          .cmdOK.Enabled = False
7680          .cmdCancel.Enabled = False

7690          .shareface_subtot_box1.Visible = False
7700          .shareface_subtot_box2.Visible = False
7710          .icash_subtot_box1.Visible = False
7720          .icash_subtot_box2.Visible = False
7730          .pcash_subtot_box1.Visible = False
7740          .pcash_subtot_box2.Visible = False
7750          .cost_subtot_box1.Visible = False
7760          .cost_subtot_box2.Visible = False
7770          .curr_id_box1.Visible = False
7780          .curr_id_box2.Visible = False

7790          .shareface_subtot.Enabled = False
7800          .shareface_subtot.Locked = False
7810          .shareface_divint_subtot.Visible = False
7820          .shareface_subtot = 0#
7830          .shareface_subtot.BorderColor = WIN_CLR_DISR
7840          .shareface_subtot.BackStyle = acBackStyleTransparent
7850          .shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
7860          .shareface_subtot_lbl_dim_hi.Visible = True
7870          .icash_subtot.Enabled = False
7880          .icash_subtot.Locked = False
7890          .icash_subtot = 0@
7900          .icash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
7910          .icash_subtot.BorderColor = WIN_CLR_DISR
7920          .icash_subtot.BackStyle = acBackStyleTransparent
7930          .icash_subtot_lbl.ForeColor = WIN_CLR_DISF
7940          .icash_subtot_lbl_dim_hi.Visible = True
7950          .pcash_subtot.Enabled = False
7960          .pcash_subtot.Locked = False
7970          .pcash_subtot = 0@
7980          .pcash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
7990          .pcash_subtot.BorderColor = WIN_CLR_DISR
8000          .pcash_subtot.BackStyle = acBackStyleTransparent
8010          .pcash_subtot_lbl.ForeColor = WIN_CLR_DISF
8020          .pcash_subtot_lbl_dim_hi.Visible = True
8030          .cost_subtot.Enabled = False
8040          .cost_subtot.Locked = False
8050          .cost_subtot = 0@
8060          .cost_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
8070          .cost_subtot.BorderColor = WIN_CLR_DISR
8080          .cost_subtot.BackStyle = acBackStyleTransparent
8090          .cost_subtot_lbl.ForeColor = WIN_CLR_DISF
8100          .cost_subtot_lbl_dim_hi.Visible = True

8110          If gblnHasForExThis = True Then
8120            .curr_id = Null
8130            .curr_id.BorderColor = WIN_CLR_DISR
8140            .curr_id.BackStyle = acBackStyleTransparent
8150            .curr_id_lbl.ForeColor = WIN_CLR_DISF
8160            .curr_id_lbl.Visible = True
8170            .curr_id_lbl_dim_hi.Visible = True
8180            .curr_id.Visible = True
8190          Else
8200            Select Case .chkIncludeCurrency
                Case True
8210              .curr_id = Null
8220              .curr_id.BorderColor = WIN_CLR_DISR
8230              .curr_id.BackStyle = acBackStyleTransparent
8240              .curr_id_lbl.ForeColor = WIN_CLR_DISF
8250              .curr_id_lbl.Visible = True
8260              .curr_id_lbl_dim_hi.Visible = True
8270              .curr_id.Visible = True
8280            Case False
8290              .curr_id.Visible = False
8300              .curr_id_lbl.Visible = False
8310              .curr_id_lbl_dim_hi.Visible = False
8320            End Select
8330          End If

8340          .subtot_lbl_dim.Visible = True
8350          .subtot_lbl_dim_hi.Visible = True
8360          .subtot_lbl.Visible = False

8370          .pending_trans = 0#
8380          .pending_trans.Enabled = False
8390          .pending_trans.Locked = False
8400          .pending_trans.BorderColor = WIN_CLR_DISR
8410          .pending_trans.BackStyle = acBackStyleTransparent

8420          DoCmd.Hourglass False

8430          Beep
8440          MsgBox "The reversing entry has been successfully created," & vbCrLf & _
                "and both entries have been hidden.", vbInformation + vbOKOnly, "Reversal Successful"

8450        End If  ' ** msgResponse.

8460      End If  ' ** blnContinue.

8470    End With

EXITP:
8480    Set rst = Nothing
8490    Set qdf = Nothing
8500    Set dbs = Nothing
8510    Exit Sub

ERRH:
8520    DoCmd.Hourglass False
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
8610    intShiftDown = (Shift And acShiftMask) > 0
8620    intAltDown = (Shift And acAltMask) > 0
8630    intCtrlDown = (Shift And acCtrlMask) > 0

8640    intRetVal = KeyCode

        ' ** Plain keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8660      Select Case intRetVal
          Case vbKeyTab
8670        With Me
8680          intRetVal = 0
8690          .cmdCancel.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab
8750        With Me
8760          intRetVal = 0
8770          .frmAccountHideTrans2_One_Sub_Pick.SetFocus
8780          .frmAccountHideTrans2_One_Sub_Pick.Form.FocusHolder.SetFocus
8790        End With
8800      End Select
8810    End If

EXITP:
8820    KeyCode = intRetVal
8830    Exit Sub

ERRH:
8840    intRetVal = 0
8850    Select Case ERR.Number
        Case Else
8860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8870    End Select
8880    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngJournalNo As Long

8910    With Me

8920      DoCmd.Hourglass True
8930      DoEvents

8940      Set rst = .frmAccountHideTrans2_One_Sub_Pick.Form.RecordsetClone
8950      rst.MoveFirst
8960      lngJournalNo = rst![journalno]  ' ** This is the existing entry.
8970      rst.Close
8980      Set rst = Nothing

8990      Set dbs = CurrentDb
9000      With dbs
            ' ** Update tblLedgerHidden_Staging1, for ledger_HIDDEN = False.
9010        Set qdf = .QueryDefs("qryAccountHideTrans2_One_03")
9020        qdf.Execute
9030        Set qdf = Nothing
            ' ** Empty tblLedgerHidden_Staging2.
9040        Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
9050        qdf.Execute
9060        Set qdf = Nothing
9070        .Close
9080      End With
9090      Set dbs = Nothing
9100      DoEvents

9110      .frmAccountHideTrans2_One_Sub_List.Form.Requery
9120      .frmAccountHideTrans2_One_Sub_List.Form.MoveRec 0, lngJournalNo  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_List.
9130      .frmAccountHideTrans2_One_Sub_List.SetFocus
9140      .frmAccountHideTrans2_One_Sub_List.Form.ledger_HIDDEN.SetFocus
9150      DoEvents

9160      .frmAccountHideTrans2_One_Sub_Pick.Form.Requery
9170      .frmAccountHideTrans2_One_Sub_Pick.Form.LabelEnable False  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.

9180      .cmdOK.Enabled = False
9190      .cmdCancel.Enabled = False

9200      .shareface_subtot_box1.Visible = False
9210      .shareface_subtot_box2.Visible = False
9220      .icash_subtot_box1.Visible = False
9230      .icash_subtot_box2.Visible = False
9240      .pcash_subtot_box1.Visible = False
9250      .pcash_subtot_box2.Visible = False
9260      .cost_subtot_box1.Visible = False
9270      .cost_subtot_box2.Visible = False
9280      .curr_id_box1.Visible = False
9290      .curr_id_box2.Visible = False

9300      .shareface_subtot.Enabled = False
9310      .shareface_subtot.Locked = False
9320      .shareface_subtot.Visible = True
9330      .shareface_divint_subtot.Visible = False
9340      .shareface_subtot = 0#
9350      .shareface_subtot.BorderColor = WIN_CLR_DISR
9360      .shareface_subtot.BackStyle = acBackStyleTransparent
9370      .shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
9380      .shareface_subtot_lbl_dim_hi.Visible = True
9390      .icash_subtot.Enabled = False
9400      .icash_subtot.Locked = False
9410      .icash_subtot = 0@
9420      .icash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
9430      .icash_subtot.BorderColor = WIN_CLR_DISR
9440      .icash_subtot.BackStyle = acBackStyleTransparent
9450      .icash_subtot_lbl.ForeColor = WIN_CLR_DISF
9460      .icash_subtot_lbl_dim_hi.Visible = True
9470      .pcash_subtot.Enabled = False
9480      .pcash_subtot.Locked = False
9490      .pcash_subtot = 0@
9500      .pcash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
9510      .pcash_subtot.BorderColor = WIN_CLR_DISR
9520      .pcash_subtot.BackStyle = acBackStyleTransparent
9530      .pcash_subtot_lbl.ForeColor = WIN_CLR_DISF
9540      .pcash_subtot_lbl_dim_hi.Visible = True
9550      .cost_subtot.Enabled = False
9560      .cost_subtot.Locked = False
9570      .cost_subtot = 0@
9580      .cost_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
9590      .cost_subtot.BorderColor = WIN_CLR_DISR
9600      .cost_subtot.BackStyle = acBackStyleTransparent
9610      .cost_subtot_lbl.ForeColor = WIN_CLR_DISF
9620      .cost_subtot_lbl_dim_hi.Visible = True

9630      If gblnHasForExThis = True Then
9640        .curr_id = Null
9650        .curr_id.BorderColor = WIN_CLR_DISR
9660        .curr_id.BackStyle = acBackStyleTransparent
9670        .curr_id.Visible = True
9680        .curr_id_lbl.ForeColor = WIN_CLR_DISF
9690        .curr_id_lbl.Visible = True
9700        .curr_id_lbl_dim_hi.Visible = True
9710      Else
9720        Select Case .chkIncludeCurrency
            Case True
9730          .curr_id = Null
9740          .curr_id.BorderColor = WIN_CLR_DISR
9750          .curr_id.BackStyle = acBackStyleTransparent
9760          .curr_id.Visible = True
9770          .curr_id_lbl.ForeColor = WIN_CLR_DISF
9780          .curr_id_lbl.Visible = True
9790          .curr_id_lbl_dim_hi.Visible = True
9800        Case False
9810          .curr_id.Visible = False
9820          .curr_id_lbl.Visible = False
9830          .curr_id_lbl_dim_hi.Visible = False
9840        End Select
9850      End If

9860      .subtot_lbl_dim.Visible = True
9870      .subtot_lbl_dim_hi.Visible = True
9880      .subtot_lbl.Visible = False

9890      .pending_trans.Enabled = False
9900      .pending_trans.Locked = False
9910      .pending_trans.BorderColor = WIN_CLR_DISR
9920      .pending_trans.BackStyle = acBackStyleTransparent
9930      .pending_trans = 0

9940      DoCmd.Hourglass False

9950    End With

EXITP:
9960    Set rst = Nothing
9970    Set qdf = Nothing
9980    Set dbs = Nothing
9990    Exit Sub

ERRH:
10000   DoCmd.Hourglass False
10010   Select Case ERR.Number
        Case Else
10020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10030   End Select
10040   Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
10110   intShiftDown = (Shift And acShiftMask) > 0
10120   intAltDown = (Shift And acAltMask) > 0
10130   intCtrlDown = (Shift And acCtrlMask) > 0

10140   intRetVal = KeyCode

        ' ** Plain keys.
10150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10160     Select Case intRetVal
          Case vbKeyTab
10170       With Me
10180         intRetVal = 0
10190         If .chkIncludeArchive.Enabled = True Then
10200           .chkIncludeArchive.SetFocus
10210         ElseIf .chkIncludeCurrency.Enabled = True Then
10220           .chkIncludeCurrency.SetFocus
10230         Else
10240           .chkShowAll.SetFocus
10250         End If
10260       End With
10270     End Select
10280   End If

        ' ** Shift keys.
10290   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10300     Select Case intRetVal
          Case vbKeyTab
10310       With Me
10320         intRetVal = 0
10330         .cmdOK.SetFocus
10340       End With
10350     End Select
10360   End If

EXITP:
10370   KeyCode = intRetVal
10380   Exit Sub

ERRH:
10390   intRetVal = 0
10400   Select Case ERR.Number
        Case Else
10410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10420   End Select
10430   Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

10510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10520   Exit Sub

ERRH:
10530   Select Case ERR.Number
        Case Else
10540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10550   End Select
10560   Resume EXITP

End Sub

Private Sub Subheader_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "Subheader_box_MouseMove"

10610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10620   Exit Sub

ERRH:
10630   Select Case ERR.Number
        Case Else
10640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10650   End Select
10660   Resume EXITP

End Sub

Private Sub Subheader01_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "Subheader01_lbl_MouseMove"

10710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10720   Exit Sub

ERRH:
10730   Select Case ERR.Number
        Case Else
10740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10750   End Select
10760   Resume EXITP

End Sub

Private Sub Subheader02_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "Subheader02_lbl_MouseMove"

10810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10820   Exit Sub

ERRH:
10830   Select Case ERR.Number
        Case Else
10840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10850   End Select
10860   Resume EXITP

End Sub

Private Sub Subheader03_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "Subheader03_lbl_MouseMove"

10910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10920   Exit Sub

ERRH:
10930   Select Case ERR.Number
        Case Else
10940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10950   End Select
10960   Resume EXITP

End Sub

Private Sub cmdHelp_GotFocus()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_GotFocus"

11010   With Me
11020     blnHelp_Focus = True
11030     Select Case .Subheader_box.Visible
          Case True
11040       .cmdHelp_L_raised_semifocus_dots_img.Visible = True
11050       .cmdHelp_R_raised_semifocus_dots_img.Visible = False
11060     Case False
11070       .cmdHelp_R_raised_semifocus_dots_img.Visible = True
11080       .cmdHelp_L_raised_semifocus_dots_img.Visible = False
11090     End Select
11100     .cmdHelp_L_raised_img.Visible = False
11110     .cmdHelp_L_raised_focus_img.Visible = False
11120     .cmdHelp_L_raised_focus_dots_img.Visible = False
11130     .cmdHelp_L_sunken_focus_dots_img.Visible = False
11140     .cmdHelp_L_raised_img_dis.Visible = False
11150     .cmdHelp_R_raised_img.Visible = False
11160     .cmdHelp_R_raised_focus_img.Visible = False
11170     .cmdHelp_R_raised_focus_dots_img.Visible = False
11180     .cmdHelp_R_sunken_focus_dots_img.Visible = False
11190     .cmdHelp_R_raised_img_dis.Visible = False
11200   End With

        'blnHelp_Focus
        'blnHelp_MouseDown

EXITP:
11210   Exit Sub

ERRH:
11220   Select Case ERR.Number
        Case Else
11230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11240   End Select
11250   Resume EXITP

End Sub

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseDown"

11310   With Me
11320     blnHelp_MouseDown = True
11330     Select Case .Subheader_box.Visible
          Case True
11340       .cmdHelp_L_sunken_focus_dots_img.Visible = True
11350       .cmdHelp_R_sunken_focus_dots_img.Visible = False
11360     Case False
11370       .cmdHelp_R_sunken_focus_dots_img.Visible = True
11380       .cmdHelp_L_sunken_focus_dots_img.Visible = False
11390     End Select
11400     .cmdHelp_L_raised_img.Visible = False
11410     .cmdHelp_L_raised_semifocus_dots_img.Visible = False
11420     .cmdHelp_L_raised_focus_img.Visible = False
11430     .cmdHelp_L_raised_focus_dots_img.Visible = False
11440     .cmdHelp_L_raised_img_dis.Visible = False
11450     .cmdHelp_R_raised_img.Visible = False
11460     .cmdHelp_R_raised_semifocus_dots_img.Visible = False
11470     .cmdHelp_R_raised_focus_img.Visible = False
11480     .cmdHelp_R_raised_focus_dots_img.Visible = False
11490     .cmdHelp_R_raised_img_dis.Visible = False
11500   End With

EXITP:
11510   Exit Sub

ERRH:
11520   Select Case ERR.Number
        Case Else
11530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11540   End Select
11550   Resume EXITP

End Sub

Public Sub cmdHelp_Click()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_Click"

        Dim ctl As Access.Control
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long

11610   With Me

11620     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
11630       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
11640     End If

          ' ** Pixels to move up.
11650     lngTmp01 = (lngSubHead_Diff / lngTpp)
11660     lngTmp02 = (lngTmp01 * lngTpp)

11670     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
11680     lngMonitorNum = 1&: lngTmp03 = 0&
11690     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
11700     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

11710     Select Case .Subheader_box.Visible
          Case True

11720       .Subheader03_lbl.Visible = False
11730       .Subheader03_bullet.Visible = False
11740       .Subheader02_lbl.Visible = False
11750       .Subheader02_bullet.Visible = False
11760       .Subheader01_lbl.Visible = False
11770       .Subheader01_bullet.Visible = False
11780       .Subheader_box.Visible = False

11790       For Each ctl In .Detail.Controls
11800         With ctl
11810           If .Name = "detail_height_line" Or .Name = "cmdClearAll_lbl" Then
                  ' ** Skip these.
11820           ElseIf Left(.Name, 7) = "cmdHelp" Then
                  ' ** I'll move these.
11830           ElseIf Left(.Name, 9) = "Subheader" Or Left(.Name, 14) = "opgAssetSource" Then
                  ' ** These are handled differently.
11840           Else
11850             .Top = (.Top - lngTmp02)
11860           End If
11870         End With
11880       Next

11890       .opgAssetSource.Top = (.opgAssetSource.Top - lngTmp02)
11900       .opgAssetSource_box.Top = (.opgAssetSource_box.Top - lngTmp02)
11910       .opgAssetSource_optType.Top = (.opgAssetSource_optType.Top - lngTmp02)
11920       .opgAssetSource_optType_lbl.Top = (.opgAssetSource_optType_lbl.Top - lngTmp02)
11930       .opgAssetSource_optName.Top = (.opgAssetSource_optName.Top - lngTmp02)
11940       .opgAssetSource_optName_lbl.Top = (.opgAssetSource_optName_lbl.Top - lngTmp02)
11950       .opgAssetSource_optCusip.Top = (.opgAssetSource_optCusip.Top - lngTmp02)
11960       .opgAssetSource_optCusip_lbl.Top = (.opgAssetSource_optCusip_lbl.Top - lngTmp02)
11970       .opgAssetSource.Height = .opgAssetSource_box.Height

11980       .cmdHelp.Top = lngTpp
11990       .cmdHelp.Left = (3& * lngTpp)
12000       .cmdHelp_L_raised_img.Top = .cmdHelp.Top
12010       .cmdHelp_L_raised_img.Left = .cmdHelp.Left
12020       .cmdHelp_L_raised_semifocus_dots_img.Top = .cmdHelp.Top
12030       .cmdHelp_L_raised_semifocus_dots_img.Left = .cmdHelp.Left
12040       .cmdHelp_L_raised_focus_img.Top = .cmdHelp.Top
12050       .cmdHelp_L_raised_focus_img.Left = .cmdHelp.Left
12060       .cmdHelp_L_raised_focus_dots_img.Top = .cmdHelp.Top
12070       .cmdHelp_L_raised_focus_dots_img.Left = .cmdHelp.Left
12080       .cmdHelp_L_sunken_focus_dots_img.Top = .cmdHelp.Top
12090       .cmdHelp_L_sunken_focus_dots_img.Left = .cmdHelp.Left
12100       .cmdHelp_L_raised_img_dis.Top = .cmdHelp.Top
12110       .cmdHelp_L_raised_img_dis.Left = .cmdHelp.Left
12120       .cmdHelp_R_raised_img.Top = .cmdHelp.Top
12130       .cmdHelp_R_raised_img.Left = .cmdHelp.Left
12140       .cmdHelp_R_raised_semifocus_dots_img.Top = .cmdHelp.Top
12150       .cmdHelp_R_raised_semifocus_dots_img.Left = .cmdHelp.Left
12160       .cmdHelp_R_raised_focus_img.Top = .cmdHelp.Top
12170       .cmdHelp_R_raised_focus_img.Left = .cmdHelp.Left
12180       .cmdHelp_R_raised_focus_dots_img.Top = .cmdHelp.Top
12190       .cmdHelp_R_raised_focus_dots_img.Left = .cmdHelp.Left
12200       .cmdHelp_R_sunken_focus_dots_img.Top = .cmdHelp.Top
12210       .cmdHelp_R_sunken_focus_dots_img.Left = .cmdHelp.Left
12220       .cmdHelp_R_raised_img_dis.Top = .cmdHelp.Top
12230       .cmdHelp_R_raised_img_dis.Left = .cmdHelp.Left

12240       .cmdHelp_R_raised_img.Visible = True
12250       .cmdHelp_L_raised_img.Visible = False
12260       .cmdHelp_L_raised_semifocus_dots_img.Visible = False
12270       .cmdHelp_L_raised_focus_img.Visible = False
12280       .cmdHelp_L_raised_focus_dots_img.Visible = False
12290       .cmdHelp_L_sunken_focus_dots_img.Visible = False
12300       .cmdHelp_L_raised_img_dis.Visible = False
12310       .cmdHelp_R_raised_semifocus_dots_img.Visible = False
12320       .cmdHelp_R_raised_focus_img.Visible = False
12330       .cmdHelp_R_raised_focus_dots_img.Visible = False
12340       .cmdHelp_R_sunken_focus_dots_img.Visible = False
12350       .cmdHelp_R_raised_img_dis.Visible = False

12360       .Detail.Height = (.Detail.Height - lngTmp02)

12370       If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
12380       DoCmd.MoveSize lngFrm_Left, lngTmp03, lngFrm_Width, (lngFrm_Height - lngTmp02)  'lngFrm_Top
12390       If lngMonitorNum > 1& Then
12400         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
12410       End If

12420       .cmdHelp_IsClosed_chk = True

12430     Case False

12440       .Detail.Height = lngDetail_Height
12450       If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
12460       DoCmd.MoveSize lngFrm_Left, lngTmp03, lngFrm_Width, lngFrm_Height  'lngFrm_Top
12470       If lngMonitorNum > 1& Then
12480         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
12490       End If

12500       For Each ctl In .Detail.Controls
12510         With ctl
12520           If .Name = "detail_height_line" Or .Name = "cmdClearAll_lbl" Or .Name = "cmdHelp_box" Or .Name = "cmdHelp_lbl" Then
                  ' ** Skip these.
12530           ElseIf Left(.Name, 9) = "Subheader" Or Left(.Name, 14) = "opgAssetSource" Then
                  ' ** These are handled differently.
12540           Else
12550             .Top = (.Top + lngTmp02)
12560           End If
12570         End With
12580       Next

12590       Do Until .opgAssetSource.Top = .cmbAssets.Top
12600         .opgAssetSource.Top = (.opgAssetSource.Top + lngTpp)
12610         .opgAssetSource_box.Top = (.opgAssetSource_box.Top + lngTpp)
12620         .opgAssetSource_optType.Top = (.opgAssetSource_optType.Top + lngTpp)
12630         .opgAssetSource_optType_lbl.Top = (.opgAssetSource_optType_lbl.Top + lngTpp)
12640         .opgAssetSource_optName.Top = (.opgAssetSource_optName.Top + lngTpp)
12650         .opgAssetSource_optName_lbl.Top = (.opgAssetSource_optName_lbl.Top + lngTpp)
12660         .opgAssetSource_optCusip.Top = (.opgAssetSource_optCusip.Top + lngTpp)
12670         .opgAssetSource_optCusip_lbl.Top = (.opgAssetSource_optCusip_lbl.Top + lngTpp)
12680       Loop

12690       .Subheader_box.Visible = True
12700       .Subheader01_bullet.Visible = True
12710       .Subheader01_lbl.Visible = True
12720       .Subheader02_bullet.Visible = True
12730       .Subheader02_lbl.Visible = True
12740       .Subheader03_bullet.Visible = True
12750       .Subheader03_lbl.Visible = True

12760       .cmdHelp.Top = (.Subheader_box.Top + lngTpp)
12770       .cmdHelp.Left = (((.Subheader_box.Left + .Subheader_box.Width) - .cmdHelp.Width) - lngTpp)
12780       .cmdHelp_L_raised_img.Top = .cmdHelp.Top
12790       .cmdHelp_L_raised_img.Left = .cmdHelp.Left
12800       .cmdHelp_L_raised_semifocus_dots_img.Top = .cmdHelp.Top
12810       .cmdHelp_L_raised_semifocus_dots_img.Left = .cmdHelp.Left
12820       .cmdHelp_L_raised_focus_img.Top = .cmdHelp.Top
12830       .cmdHelp_L_raised_focus_img.Left = .cmdHelp.Left
12840       .cmdHelp_L_raised_focus_dots_img.Top = .cmdHelp.Top
12850       .cmdHelp_L_raised_focus_dots_img.Left = .cmdHelp.Left
12860       .cmdHelp_L_sunken_focus_dots_img.Top = .cmdHelp.Top
12870       .cmdHelp_L_sunken_focus_dots_img.Left = .cmdHelp.Left
12880       .cmdHelp_L_raised_img_dis.Top = .cmdHelp.Top
12890       .cmdHelp_L_raised_img_dis.Left = .cmdHelp.Left
12900       .cmdHelp_R_raised_img.Top = .cmdHelp.Top
12910       .cmdHelp_R_raised_img.Left = .cmdHelp.Left
12920       .cmdHelp_R_raised_semifocus_dots_img.Top = .cmdHelp.Top
12930       .cmdHelp_R_raised_semifocus_dots_img.Left = .cmdHelp.Left
12940       .cmdHelp_R_raised_focus_img.Top = .cmdHelp.Top
12950       .cmdHelp_R_raised_focus_img.Left = .cmdHelp.Left
12960       .cmdHelp_R_raised_focus_dots_img.Top = .cmdHelp.Top
12970       .cmdHelp_R_raised_focus_dots_img.Left = .cmdHelp.Left
12980       .cmdHelp_R_sunken_focus_dots_img.Top = .cmdHelp.Top
12990       .cmdHelp_R_sunken_focus_dots_img.Left = .cmdHelp.Left
13000       .cmdHelp_R_raised_img_dis.Top = .cmdHelp.Top
13010       .cmdHelp_R_raised_img_dis.Left = .cmdHelp.Left

13020       .cmdHelp_L_raised_img.Visible = True
13030       .cmdHelp_L_raised_semifocus_dots_img.Visible = False
13040       .cmdHelp_L_raised_focus_img.Visible = False
13050       .cmdHelp_L_raised_focus_dots_img.Visible = False
13060       .cmdHelp_L_sunken_focus_dots_img.Visible = False
13070       .cmdHelp_L_raised_img_dis.Visible = False
13080       .cmdHelp_R_raised_img.Visible = False
13090       .cmdHelp_R_raised_semifocus_dots_img.Visible = False
13100       .cmdHelp_R_raised_focus_img.Visible = False
13110       .cmdHelp_R_raised_focus_dots_img.Visible = False
13120       .cmdHelp_R_sunken_focus_dots_img.Visible = False
13130       .cmdHelp_R_raised_img_dis.Visible = False

13140       .cmdHelp_IsClosed_chk = False

13150     End Select

13160   End With

EXITP:
13170   Exit Sub

ERRH:
13180   Select Case ERR.Number
        Case Else
13190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13200   End Select
13210   Resume EXITP

End Sub

Private Sub cmdHelp_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
13310   intShiftDown = (Shift And acShiftMask) > 0
13320   intAltDown = (Shift And acAltMask) > 0
13330   intCtrlDown = (Shift And acCtrlMask) > 0

13340   intRetVal = KeyCode

        ' ** Plain keys.
13350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13360     Select Case intRetVal
          Case vbKeyTab
13370       With Me
13380         intRetVal = 0
13390         lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
13400         If lngRecsCur > 0 Then
13410           .frmAccountHideTrans2_One_Sub_List.SetFocus
13420         Else
13430           .DateStart.SetFocus
13440         End If
13450       End With
13460     End Select
13470   End If

        ' ** Shift keys.
13480   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13490     Select Case intRetVal
          Case vbKeyTab
13500       With Me
13510         intRetVal = 0
13520         .cmdClose.SetFocus
13530       End With
13540     End Select
13550   End If

EXITP:
13560   KeyCode = intRetVal
13570   Exit Sub

ERRH:
13580   intRetVal = 0
13590   Select Case ERR.Number
        Case Else
13600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13610   End Select
13620   Resume EXITP

End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseMove"

13710   With Me
13720     If blnHelp_MouseDown = False Then
13730       Select Case .Subheader_box.Visible
            Case True
13740         Select Case blnHelp_Focus
              Case True
13750           .cmdHelp_L_raised_focus_dots_img.Visible = True
13760           .cmdHelp_L_raised_focus_img.Visible = False
13770         Case False
13780           .cmdHelp_L_raised_focus_img.Visible = True
13790           .cmdHelp_L_raised_focus_dots_img.Visible = False
13800         End Select
13810         .cmdHelp_R_raised_focus_img.Visible = False
13820         .cmdHelp_R_raised_focus_dots_img.Visible = False
13830       Case False
13840         Select Case blnHelp_Focus
              Case True
13850           .cmdHelp_R_raised_focus_dots_img.Visible = True
13860           .cmdHelp_R_raised_focus_img.Visible = False
13870         Case False
13880           .cmdHelp_R_raised_focus_img.Visible = True
13890           .cmdHelp_R_raised_focus_dots_img.Visible = False
13900         End Select
13910         .cmdHelp_L_raised_focus_img.Visible = False
13920         .cmdHelp_L_raised_focus_dots_img.Visible = False
13930       End Select
13940       .cmdHelp_L_raised_img.Visible = False
13950       .cmdHelp_L_raised_semifocus_dots_img.Visible = False
13960       .cmdHelp_L_sunken_focus_dots_img.Visible = False
13970       .cmdHelp_L_raised_img_dis.Visible = False
13980       .cmdHelp_R_raised_img.Visible = False
13990       .cmdHelp_R_raised_semifocus_dots_img.Visible = False
14000       .cmdHelp_R_sunken_focus_dots_img.Visible = False
14010       .cmdHelp_R_raised_img_dis.Visible = False
14020     End If
14030   End With

EXITP:
14040   Exit Sub

ERRH:
14050   Select Case ERR.Number
        Case Else
14060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14070   End Select
14080   Resume EXITP

End Sub

Private Sub cmdHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseUp"

14110   With Me
14120     Select Case .Subheader_box.Visible
          Case True
14130       .cmdHelp_L_raised_focus_dots_img.Visible = True
14140       .cmdHelp_R_raised_focus_dots_img.Visible = False
14150     Case False
14160       .cmdHelp_R_raised_focus_dots_img.Visible = True
14170       .cmdHelp_L_raised_focus_dots_img.Visible = False
14180     End Select
14190     .cmdHelp_L_raised_img.Visible = False
14200     .cmdHelp_L_raised_semifocus_dots_img.Visible = False
14210     .cmdHelp_L_raised_focus_img.Visible = False
14220     .cmdHelp_L_sunken_focus_dots_img.Visible = False
14230     .cmdHelp_L_raised_img_dis.Visible = False
14240     .cmdHelp_R_raised_img.Visible = False
14250     .cmdHelp_R_raised_semifocus_dots_img.Visible = False
14260     .cmdHelp_R_raised_focus_img.Visible = False
14270     .cmdHelp_R_sunken_focus_dots_img.Visible = False
14280     .cmdHelp_R_raised_img_dis.Visible = False
14290     blnHelp_MouseDown = False
14300   End With

EXITP:
14310   Exit Sub

ERRH:
14320   Select Case ERR.Number
        Case Else
14330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14340   End Select
14350   Resume EXITP

End Sub

Private Sub cmdHelp_LostFocus()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_LostFocus"

14410   With Me
14420     Select Case .Subheader_box.Visible
          Case True
14430       .cmdHelp_L_raised_img.Visible = True
14440       .cmdHelp_R_raised_img.Visible = False
14450     Case False
14460       .cmdHelp_R_raised_img.Visible = True
14470       .cmdHelp_L_raised_img.Visible = False
14480     End Select
14490     .cmdHelp_L_raised_semifocus_dots_img.Visible = False
14500     .cmdHelp_L_raised_focus_img.Visible = False
14510     .cmdHelp_L_raised_focus_dots_img.Visible = False
14520     .cmdHelp_L_sunken_focus_dots_img.Visible = False
14530     .cmdHelp_L_raised_img_dis.Visible = False
14540     .cmdHelp_R_raised_semifocus_dots_img.Visible = False
14550     .cmdHelp_R_raised_focus_img.Visible = False
14560     .cmdHelp_R_raised_focus_dots_img.Visible = False
14570     .cmdHelp_R_sunken_focus_dots_img.Visible = False
14580     .cmdHelp_R_raised_img_dis.Visible = False
14590     blnHelp_Focus = False
14600   End With

EXITP:
14610   Exit Sub

ERRH:
14620   Select Case ERR.Number
        Case Else
14630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14640   End Select
14650   Resume EXITP

End Sub

Private Sub frmAccountHideTrans2_One_Sub_List_Enter()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "frmAccountHideTrans2_One_Sub_List_Enter"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

14710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14720   Exit Sub

ERRH:
14730   Select Case ERR.Number
        Case Else
14740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14750   End Select
14760   Resume EXITP

End Sub

Private Sub frmAccountHideTrans2_One_Sub_Pick_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "frmAccountHideTrans2_One_Sub_Pick_box_MouseMove"

14810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14820   Exit Sub

ERRH:
14830   Select Case ERR.Number
        Case Else
14840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14850   End Select
14860   Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

        Dim strFilter As String
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, strTmp02 As String

14910   With Me
14920     If blnHasArchive = True Then
14930       strFilter = .frmAccountHideTrans2_One_Sub_List.Form.Filter
14940       Select Case .chkIncludeArchive
            Case True

14950         .chkIncludeArchive_lbl.FontBold = True
14960         .archive_trans_lbl.Visible = True
14970         .archive_trans_lbl_box.Visible = True

14980         If strFilter <> vbNullString Then
14990           intPos01 = InStr(strFilter, "[src]")
15000           If intPos01 > 0 Then
15010             If intPos01 = 1 Then
                    ' ** It's the first thing on the line.
15020               intPos02 = InStr(strFilter, " And ")
15030               If intPos02 = 0 Then
                      ' ** It's the only thing.
15040                 strFilter = vbNullString
15050               Else
                      ' ** There were others after it.
15060                 strTmp01 = Trim(Mid(strFilter, intPos02))
15070                 If Left(strTmp01, 4) = "And " Then strTmp01 = Trim(Mid(strTmp01, 4))
15080                 strFilter = strTmp01
15090               End If
15100             Else
                    ' ** Other criteria are before it.
15110               strTmp01 = Trim(Left(strFilter, (intPos01 - 1)))  ' ** Includes preceding 'And'.
15120               strTmp02 = Mid(strFilter, intPos01)  ' ** From [src] to end of line.
15130               intPos02 = InStr(strTmp02, " And ")
15140               If intPos02 > 0 Then
                      ' ** There are others after it.
15150                 strTmp02 = Trim(Mid(strTmp02, (intPos02 + 5)))  ' ** Without the trailing 'And'.
15160                 strFilter = strTmp01 & " " & strTmp02
15170               Else
                      ' ** It's the last thing on the line.
15180                 If Right(strTmp01, 4) = " And" Then strTmp01 = Trim(Left(strTmp01, (Len(strTmp01) - 3)))
15190                 strFilter = strTmp01
15200               End If
15210             End If
15220           Else
                  ' ** Not filtered anyway.
15230           End If
15240         End If

15250       Case False

15260         .chkIncludeArchive_lbl.FontBold = False
15270         .archive_trans_lbl.Visible = False
15280         .archive_trans_lbl_box.Visible = False

15290         If strFilter = vbNullString Then
                ' ** This will be the only thing.
15300           strFilter = "[src] <> 'A'"
15310         Else
15320           intPos01 = InStr(strFilter, "[src]")
15330           If intPos01 > 0 Then
                  ' ** It seems to already be filtered!
15340           Else
15350             strFilter = strFilter & " And [src] <> 'A'"
15360           End If
15370         End If
15380       End Select

15390       .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
15400       If strFilter <> vbNullString Then
15410         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = True
15420       Else
15430         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False
15440       End If

            'RESET SORT!

15450     End If
15460   End With

EXITP:
15470   Exit Sub

ERRH:
15480   Select Case ERR.Number
        Case Else
15490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15500   End Select
15510   Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
15610   intShiftDown = (Shift And acShiftMask) > 0
15620   intAltDown = (Shift And acAltMask) > 0
15630   intCtrlDown = (Shift And acCtrlMask) > 0

15640   intRetVal = KeyCode

        ' ** Plain keys.
15650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15660     Select Case intRetVal
          Case vbKeyTab
15670       With Me
15680         intRetVal = 0
15690         If .chkIncludeCurrency.Enabled = True Then
15700           .chkIncludeCurrency.SetFocus
15710         Else
15720           .chkShowAll.SetFocus
15730         End If
15740       End With
15750     End Select
15760   End If

        ' ** Shift keys.
15770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15780     Select Case intRetVal
          Case vbKeyTab
15790       With Me
15800         intRetVal = 0
15810         If .cmdCancel.Enabled = True Then
15820           .cmdCancel.SetFocus
15830         Else
15840           lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
15850           If lngRecsCur > 0& Then
15860             .frmAccountHideTrans2_One_Sub_List.SetFocus
15870             .frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
15880           Else
15890             .cmdClose.SetFocus
15900           End If
15910         End If
15920       End With
15930     End Select
15940   End If

EXITP:
15950   KeyCode = intRetVal
15960   Exit Sub

ERRH:
15970   intRetVal = 0
15980   Select Case ERR.Number
        Case Else
15990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16000   End Select
16010   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

        Dim lngCurrID As Long
        Dim lngTmp01 As Long

16110   With Me
16120     DoCmd.Hourglass True
16130     DoEvents
16140     gblnHasForExThis = HasForEx_Acct(gstrAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.
16150     Select Case .chkIncludeCurrency
          Case True
16160       .chkIncludeCurrency_lbl.FontBold = True
16170       .frmAccountHideTrans2_One_Sub_List.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_List.
16180       .frmAccountHideTrans2_One_Sub_Pick.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.
16190       If blnIsOpen = True And blnHasCurrent = False Then
16200         .frmAccountHideTrans2_One_Sub_List.Form.curr_id.Visible = False
16210         .frmAccountHideTrans2_One_Sub_List.Form.curr_id_box.Visible = False
16220         .frmAccountHideTrans2_One_Sub_List.Form.curr_id_forex.Visible = False
16230       End If
16240       DoEvents
16250       .shareface_subtot.Left = lngSharefaceSub_Left
16260       .shareface_divint_subtot.Left = .shareface_subtot.Left
16270       .shareface_subtot_box1.Left = (.shareface_subtot.Left - (2& * lngTpp))
16280       .shareface_subtot_box2.Left = (.shareface_subtot.Left - lngTpp)
16290       .shareface_subtot_lbl.Left = .shareface_subtot.Left
16300       .shareface_subtot_lbl_dim_hi.Left = (.shareface_subtot_lbl.Left + lngTpp)
16310       .icash_subtot.Left = lngICashSub_Left
16320       .icash_subtot_box1.Left = (.icash_subtot.Left - (2& * lngTpp))
16330       .icash_subtot_box2.Left = (.icash_subtot.Left - lngTpp)
16340       .icash_subtot_lbl.Left = .icash_subtot.Left
16350       .icash_subtot_lbl_dim_hi.Left = (.icash_subtot_lbl.Left + lngTpp)
16360       .pcash_subtot.Left = lngPCashSub_Left
16370       .pcash_subtot_box1.Left = (.pcash_subtot.Left - (2& * lngTpp))
16380       .pcash_subtot_box2.Left = (.pcash_subtot.Left - lngTpp)
16390       .pcash_subtot_lbl.Left = .pcash_subtot.Left
16400       .pcash_subtot_lbl_dim_hi.Left = (.pcash_subtot_lbl.Left + lngTpp)
16410       .cost_subtot.Left = lngCostSub_Left
16420       .cost_subtot_box1.Left = (.cost_subtot.Left - (2& * lngTpp))
16430       .cost_subtot_box2.Left = (.cost_subtot.Left - lngTpp)
16440       .cost_subtot_lbl.Left = .cost_subtot.Left
16450       .cost_subtot_lbl_dim_hi.Left = (.cost_subtot_lbl.Left + lngTpp)
16460       .curr_id.Left = lngCurrID_Left
16470       .curr_id_box1.Left = (.curr_id.Left - (2& * lngTpp))
16480       .curr_id_box2.Left = (.curr_id.Left - lngTpp)
16490       .curr_id_lbl.Left = .curr_id.Left
16500       .curr_id_lbl_dim_hi.Left = (.curr_id_lbl.Left + lngTpp)
16510       .subtot_lbl.Left = (.shareface_subtot.Left - lngSubTotLbl_Offset)
16520       .subtot_lbl_dim.Left = .subtot_lbl.Left
16530       .subtot_lbl_dim_hi.Left = (.subtot_lbl_dim.Left + lngTpp)
16540       .pending_trans.Left = lngPTrans_Left
16550       .pending_trans_lbl.Left = (.pending_trans.Left - lngPTransLbl_Offset)
16560       .cmdOK.Left = lngOK_Left
16570       .cmdCancel.Left = lngCancel_Left
16580     Case False
16590       .chkIncludeCurrency_lbl.FontBold = False
16600       .frmAccountHideTrans2_One_Sub_List.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_List.
16610       .frmAccountHideTrans2_One_Sub_Pick.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.
16620       DoEvents
16630       lngTmp01 = (57& * 15&)
16640       .shareface_subtot.Left = (lngSharefaceSub_Left + lngTmp01)
16650       .shareface_divint_subtot.Left = .shareface_subtot.Left
16660       .shareface_subtot_box1.Left = (.shareface_subtot.Left - (2& * lngTpp))
16670       .shareface_subtot_box2.Left = (.shareface_subtot.Left - lngTpp)
16680       .shareface_subtot_lbl.Left = .shareface_subtot.Left
16690       .shareface_subtot_lbl_dim_hi.Left = (.shareface_subtot_lbl.Left + lngTpp)
16700       .icash_subtot.Left = (lngICashSub_Left + lngTmp01)
16710       .icash_subtot_box1.Left = (.icash_subtot.Left - (2& * lngTpp))
16720       .icash_subtot_box2.Left = (.icash_subtot.Left - lngTpp)
16730       .icash_subtot_lbl.Left = .icash_subtot.Left
16740       .icash_subtot_lbl_dim_hi.Left = (.icash_subtot_lbl.Left + lngTpp)
16750       .pcash_subtot.Left = (lngPCashSub_Left + lngTmp01)
16760       .pcash_subtot_box1.Left = (.pcash_subtot.Left - (2& * lngTpp))
16770       .pcash_subtot_box2.Left = (.pcash_subtot.Left - lngTpp)
16780       .pcash_subtot_lbl.Left = .pcash_subtot.Left
16790       .pcash_subtot_lbl_dim_hi.Left = (.pcash_subtot_lbl.Left + lngTpp)
16800       .cost_subtot.Left = (lngCostSub_Left + lngTmp01)
16810       .cost_subtot_box1.Left = (.cost_subtot.Left - (2& * lngTpp))
16820       .cost_subtot_box2.Left = (.cost_subtot.Left - lngTpp)
16830       .cost_subtot_lbl.Left = .cost_subtot.Left
16840       .cost_subtot_lbl_dim_hi.Left = (.cost_subtot_lbl.Left + lngTpp)
16850       .curr_id.Left = (lngCurrID_Left + lngTmp01)
16860       .curr_id_box1.Left = (.curr_id.Left - (2& * lngTpp))
16870       .curr_id_box2.Left = (.curr_id.Left - lngTpp)
16880       .curr_id_lbl.Left = .curr_id.Left
16890       .curr_id_lbl_dim_hi.Left = (.curr_id_lbl.Left + lngTpp)
16900       .subtot_lbl.Left = (.shareface_subtot.Left - lngSubTotLbl_Offset)
16910       .subtot_lbl_dim.Left = .subtot_lbl.Left
16920       .subtot_lbl_dim_hi.Left = (.subtot_lbl_dim.Left + lngTpp)
16930       .pending_trans.Left = (lngPTrans_Left + lngTmp01)
16940       .pending_trans_lbl.Left = (.pending_trans.Left - lngPTransLbl_Offset)
16950       .cmdOK.Left = (lngOK_Left + lngTmp01)
16960       .cmdCancel.Left = (lngCancel_Left + lngTmp01)
16970     End Select
16980     DoEvents

16990     If (gblnHasForExThis = True) Or (gblnHasForExThis = False And .chkIncludeCurrency = True) Then
17000       lngRecsCur = .frmAccountHideTrans2_One_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_Pick.
17010       If lngRecsCur > 0& Then
17020         lngCurrID = .frmAccountHideTrans2_One_Sub_Pick.Form.curr_id
17030         .curr_id = lngCurrID
17040         .curr_id.Enabled = False
17050         .curr_id.Locked = True
17060         .curr_id.BorderColor = CLR_LTBLU2
17070         .curr_id.BackStyle = acBackStyleNormal
17080         If lngCurrID = 150& Then
17090           .curr_id.BackColor = CLR_DISABLED_BG
17100         Else
17110           .curr_id.BackColor = CLR_LTGRN
17120         End If
17130         .curr_id_lbl.ForeColor = CLR_DKGRY3
17140         .curr_id_lbl_dim_hi.Visible = False
17150         .curr_id_box1.Visible = True
17160         .curr_id_box2.Visible = True
17170       Else
17180         .curr_id = Null
17190         .curr_id.Enabled = False
17200         .curr_id.Locked = False
17210         .curr_id.BorderColor = WIN_CLR_DISR
17220         .curr_id.BackStyle = acBackStyleTransparent
17230         .curr_id_lbl.ForeColor = WIN_CLR_DISF
17240         .curr_id_lbl_dim_hi.Visible = True
17250         .curr_id_box1.Visible = False
17260         .curr_id_box2.Visible = False
17270       End If
17280       .curr_id.Visible = True
17290       .curr_id_lbl.Visible = True
17300     Else
17310       .curr_id.Visible = False
17320       .curr_id_lbl.Visible = False
17330       .curr_id_lbl_dim_hi.Visible = False
17340       .curr_id_box1.Visible = False
17350       .curr_id_box2.Visible = False
17360     End If
17370     DoEvents

17380     lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
17390     If lngRecsCur > 0& Then
17400       .frmAccountHideTrans2_One_Sub_List.SetFocus
17410       .frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
17420     End If

17430     DoCmd.Hourglass False

17440   End With

EXITP:
17450   Exit Sub

ERRH:
17460   DoCmd.Hourglass False
17470   Select Case ERR.Number
        Case Else
17480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17490   End Select
17500   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
17610   intShiftDown = (Shift And acShiftMask) > 0
17620   intAltDown = (Shift And acAltMask) > 0
17630   intCtrlDown = (Shift And acCtrlMask) > 0

17640   intRetVal = KeyCode

        ' ** Plain keys.
17650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17660     Select Case intRetVal
          Case vbKeyTab
17670       With Me
17680         intRetVal = 0
17690         .chkShowAll.SetFocus
17700       End With
17710     End Select
17720   End If

        ' ** Shift keys.
17730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17740     Select Case intRetVal
          Case vbKeyTab
17750       With Me
17760         intRetVal = 0
17770         If .chkIncludeArchive.Enabled = True Then
17780           .chkIncludeArchive.SetFocus
17790         ElseIf .cmdCancel.Enabled = True Then
17800           .cmdCancel.SetFocus
17810         Else
17820           lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
17830           If lngRecsCur > 0& Then
17840             .frmAccountHideTrans2_One_Sub_List.SetFocus
17850             .frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
17860           Else
17870             .cmdClose.SetFocus
17880           End If
17890         End If
17900       End With
17910     End Select
17920   End If

EXITP:
17930   KeyCode = intRetVal
17940   Exit Sub

ERRH:
17950   intRetVal = 0
17960   Select Case ERR.Number
        Case Else
17970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17980   End Select
17990   Resume EXITP

End Sub

Private Sub chkShowAll_AfterUpdate()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAll_AfterUpdate"

        Dim strFilter As String

18010   With Me
18020     strFilter = vbNullString
18030     Select Case .chkShowAll
          Case True
            ' ** This will override all the other filters, except Archive.
18040       .chkShowAll_lbl.FontBold = True
18050       .LastStmt_Trans_lbl.Visible = False
18060       DoEvents
18070       If blnHasArchive = True Then
18080         .chkIncludeArchive.Enabled = True
18090         If .chkIncludeArchive = False Then
18100           strFilter = "[src] <> 'A'"
18110         End If
18120       End If
18130       .DateStart = Null
18140       .DateEnd = Null
18150       .cmbJournalTypes = Null
18160       .cmbAssets = Null
18170       DoEvents
18180       If strFilter = vbNullString Then
18190         .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
18200         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False
18210       Else
18220         .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
18230         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = True
18240       End If
18250       DoEvents
18260       If .cmdClearAll.Enabled = True Then
18270         .cmdClearAll.Enabled = False
18280         .cmdClearAll_raised_img_dis.Visible = True
18290         .cmdClearAll_raised_img.Visible = False
18300         .cmdClearAll_raised_semifocus_dots_img.Visible = False
18310         .cmdClearAll_raised_focus_img.Visible = False
18320         .cmdClearAll_raised_focus_dots_img.Visible = False
18330         .cmdClearAll_sunken_focus_dots_img.Visible = False
18340       End If
18350       DoEvents
18360       .frmAccountHideTrans2_One_Sub_List.Form.Requery
18370       DoEvents
18380       .frmAccountHideTrans2_One_Sub_List.SetFocus
18390     Case False
            ' ** By definition, this is just since last balance date.
18400       .chkShowAll_lbl.FontBold = False
18410       If datLastBalDate = 0 Then
              ' ** No filter necessary, all are valid.
18420       Else
18430         strFilter = "[transdate] > #" & Format(datLastBalDate, "mm/dd/yyyy") & "#"
18440       End If
18450       .LastStmt_Trans_lbl.Visible = True
            ' ** Now check the other filters.
18460       If .chkIncludeArchive = True Then
18470         .chkIncludeArchive = False
18480         .chkIncludeArchive_lbl.FontBold = False
18490         .archive_trans_lbl.Visible = False
18500         .archive_trans_lbl_box.Visible = False
18510       End If
18520       If blnHasArchive = True Then
18530         .chkIncludeArchive.Enabled = False
18540       End If
18550       If IsNull(.DateStart) = False Then
18560         If .DateStart <= datLastBalDate Then
18570           .DateStart = DateAdd("y", 1, datLastBalDate)
18580         ElseIf .DateStart > datLastBalDate Then
                ' ** Replace the first filter.
18590           strFilter = "[transdate] >= #" & Format(.DateStart, "mm/dd/yyyy") & "#"
18600         End If
18610       End If
18620       If IsNull(.DateEnd) = False Then
18630         If .DateEnd <= datLastBalDate Then
18640           .DateEnd = Null
18650         Else
18660           If strFilter <> vbNullString Then strFilter = strFilter & " And "
18670           strFilter = strFilter & "[transdate] <= #" & Format(.DateEnd, "mm/dd/yyyy") & "#"
18680         End If
18690       End If
18700       If IsNull(.cmbJournalTypes) = False Then
18710         If strFilter <> vbNullString Then strFilter = strFilter & " And "
18720         strFilter = strFilter & "[journaltype] = '" & .cmbJournalTypes & "'"
18730       End If
18740       If IsNull(.cmbAssets) = False Then
18750         If strFilter <> vbNullString Then strFilter = strFilter & " And "
18760         strFilter = strFilter & "[assetno] = " & CStr(.cmbAssets)
18770       End If
18780       If strFilter = vbNullString Then
18790         .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
18800         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False
18810       Else
18820         .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
18830         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = True
18840       End If
18850       DoEvents
18860       If IsNull(.DateStart) = True And IsNull(.DateEnd) = True And IsNull(.cmbJournalTypes) = True And _
                IsNull(.cmbAssets) = True Then
18870         If .cmdClearAll.Enabled = True Then
18880           .cmdClearAll.Enabled = False
18890           .cmdClearAll_raised_img_dis.Visible = True
18900           .cmdClearAll_raised_img.Visible = False
18910           .cmdClearAll_raised_semifocus_dots_img.Visible = False
18920           .cmdClearAll_raised_focus_img.Visible = False
18930           .cmdClearAll_raised_focus_dots_img.Visible = False
18940           .cmdClearAll_sunken_focus_dots_img.Visible = False
18950         End If
18960       End If
18970       .frmAccountHideTrans2_One_Sub_List.Form.Requery
18980       DoEvents
18990       .frmAccountHideTrans2_One_Sub_List.SetFocus
19000     End Select
19010   End With

EXITP:
19020   Exit Sub

ERRH:
19030   Select Case ERR.Number
        Case Else
19040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19050   End Select
19060   Resume EXITP

End Sub

Private Sub chkShowAll_KeyDown(KeyCode As Integer, Shift As Integer)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
19110   intShiftDown = (Shift And acShiftMask) > 0
19120   intAltDown = (Shift And acAltMask) > 0
19130   intCtrlDown = (Shift And acCtrlMask) > 0

19140   intRetVal = KeyCode

        ' ** Plain keys.
19150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19160     Select Case intRetVal
          Case vbKeyTab
19170       With Me
19180         intRetVal = 0
19190         .chkShowJournalNo.SetFocus
19200       End With
19210     End Select
19220   End If

        ' ** Shift keys.
19230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19240     Select Case intRetVal
          Case vbKeyTab
19250       With Me
19260         intRetVal = 0
19270         If .chkIncludeCurrency.Enabled = True Then
19280           .chkIncludeCurrency.SetFocus
19290         ElseIf .chkIncludeArchive.Enabled = True Then
19300           .chkIncludeArchive.SetFocus
19310         ElseIf .cmdCancel.Enabled = True Then
19320           .cmdCancel.SetFocus
19330         Else
19340           lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
19350           If lngRecsCur > 0& Then
19360             .frmAccountHideTrans2_One_Sub_List.SetFocus
19370             .frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
19380           Else
19390             .cmdClose.SetFocus
19400           End If
19410         End If
19420       End With
19430     End Select
19440   End If

EXITP:
19450   KeyCode = intRetVal
19460   Exit Sub

ERRH:
19470   intRetVal = 0
19480   Select Case ERR.Number
        Case Else
19490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19500   End Select
19510   Resume EXITP

End Sub

Private Sub chkShowJournalNo_AfterUpdate()

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowJournalNo_AfterUpdate"

19610   With Me
19620     Select Case .chkShowJournalNo
          Case True
19630       .chkShowJournalNo_lbl.FontBold = True
19640       .frmAccountHideTrans2_One_Sub_List.Form.ShowJournalNo_Sub True  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_List.
19650       .frmAccountHideTrans2_One_Sub_Pick.Form.ShowJournalNo_Sub True  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.
19660       If blnIsOpen = True And blnHasCurrent = False Then
19670         .frmAccountHideTrans2_One_Sub_List.Form.journalno.Visible = False
19680       End If
19690     Case False
19700       .chkShowJournalNo_lbl.FontBold = False
19710       .frmAccountHideTrans2_One_Sub_List.Form.ShowJournalNo_Sub False  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_List.
19720       .frmAccountHideTrans2_One_Sub_Pick.Form.ShowJournalNo_Sub False  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.
19730     End Select
19740     lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
19750     If lngRecsCur > 0& Then
19760       .frmAccountHideTrans2_One_Sub_List.SetFocus
19770       .frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
19780     End If
19790   End With

EXITP:
19800   Exit Sub

ERRH:
19810   Select Case ERR.Number
        Case Else
19820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19830   End Select
19840   Resume EXITP

End Sub

Private Sub chkShowJournalNo_KeyDown(KeyCode As Integer, Shift As Integer)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowJournalNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
19910   intShiftDown = (Shift And acShiftMask) > 0
19920   intAltDown = (Shift And acAltMask) > 0
19930   intCtrlDown = (Shift And acCtrlMask) > 0

19940   intRetVal = KeyCode

        ' ** Plain keys.
19950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19960     Select Case intRetVal
          Case vbKeyTab
19970       With Me
19980         intRetVal = 0
19990         .DateStart.SetFocus
20000       End With
20010     End Select
20020   End If

        ' ** Shift keys.
20030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20040     Select Case intRetVal
          Case vbKeyTab
20050       With Me
20060         intRetVal = 0
20070         .chkShowAll.SetFocus
20080       End With
20090     End Select
20100   End If

EXITP:
20110   KeyCode = intRetVal
20120   Exit Sub

ERRH:
20130   intRetVal = 0
20140   Select Case ERR.Number
        Case Else
20150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20160   End Select
20170   Resume EXITP

End Sub

Private Sub Filter_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "Filter_box_MouseMove"

20210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
20220   Exit Sub

ERRH:
20230   Select Case ERR.Number
        Case Else
20240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20250   End Select
20260   Resume EXITP

End Sub

Private Sub DateStart_BeforeUpdate(Cancel As Integer)

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_BeforeUpdate"

20310   With Me
20320     If IsNull(.DateStart) = False Then
20330       If IsDate(.DateStart) = True Then
              ' ** Continue.
20340       Else
20350         Beep
20360         MsgBox "Please enter a valid date." & vbCrLf & vbCrLf & vbCrLf & "msg 2", vbInformation + vbOKOnly, "Invalid Date"
20370         Cancel = -1
20380       End If
20390     End If
20400   End With

EXITP:
20410   Exit Sub

ERRH:
20420   Select Case ERR.Number
        Case Else
20430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20440   End Select
20450   Resume EXITP

End Sub

Private Sub DateStart_AfterUpdate()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_AfterUpdate"

        Dim strFilter As String
        Dim datStartDate As Date, datEndDate As Date
        Dim blnContinue As Boolean
        Dim strTmp01 As String

20510   With Me
20520     blnContinue = True
20530     strFilter = vbNullString: strTmp01 = vbNullString

20540     If IsNull(.DateStart) = False Then
20550       Select Case IsDate(.DateStart)
            Case True
20560         datStartDate = .DateStart
20570         If IsNull(.DateEnd) = False Then
20580           datEndDate = .DateEnd
20590           If datStartDate <= datEndDate Then
20600             strTmp01 = "[transdate] >= #" & Format(datStartDate, "mm/dd/yyyy") & "# And " & _
                    "[transdate] <= #" & Format(datEndDate, "mm/dd/yyyy") & "#"
20610           Else
20620             .DateEnd = Null
20630             strTmp01 = "[transdate] >= #" & Format(datStartDate, "mm/dd/yyyy") & "#"
20640           End If
20650         Else
20660           strTmp01 = "[transdate] >= #" & Format(datStartDate, "mm/dd/yyyy") & "#"
20670         End If
20680       Case False
20690         blnContinue = False
20700         Beep
20710         MsgBox "Please enter a valid date." & vbCrLf & vbCrLf & vbCrLf & "msg 3", vbInformation + vbOKOnly, "Invalid Date"
20720         .DateStart.Undo
20730         .DateStart.SetFocus
20740       End Select
20750     Else
20760       If IsNull(.DateEnd) = False Then
20770         datEndDate = .DateEnd
20780         strTmp01 = "[transdate] <= #" & Format(datEndDate, "mm/dd/yyyy") & "#"
20790       End If
20800     End If

20810     If blnContinue = True Then
20820       If strTmp01 <> vbNullString Then
20830         strFilter = strTmp01
20840       End If
20850       If IsNull(.cmbJournalTypes) = False Then
20860         If strFilter <> vbNullString Then strFilter = strFilter & " And "
20870         strFilter = strFilter & "[journaltype] = '" & .cmbJournalTypes & "'"
20880       End If
20890       If IsNull(.cmbAssets) = False Then
20900         If strFilter <> vbNullString Then strFilter = strFilter & " And "
20910         strFilter = strFilter & "[assetno] = " & CStr(.cmbAssets)
20920       End If
20930       If .chkIncludeArchive.Enabled = True Then
20940         If .chkIncludeArchive = False Then
20950           If strFilter <> vbNullString Then strFilter = strFilter & " And "
20960           strFilter = strFilter & "[src] <> 'A'"
20970         End If
20980       End If
20990       .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
21000       If strFilter = vbNullString Then
21010         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False
21020       Else
21030         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = True
21040       End If
21050       .frmAccountHideTrans2_One_Sub_List.Form.Requery
21060     End If  ' ** blnContinue.

21070     If IsNull(.DateStart) = True And IsNull(.DateEnd) = True And IsNull(.cmbJournalTypes) = True And IsNull(.cmbAssets) = True Then
21080       If .cmdClearAll.Enabled = True Then
21090         .cmdClearAll.Enabled = False
21100         .cmdClearAll_raised_img_dis.Visible = True
21110         .cmdClearAll_raised_img.Visible = False
21120         .cmdClearAll_raised_semifocus_dots_img.Visible = False
21130         .cmdClearAll_raised_focus_img.Visible = False
21140         .cmdClearAll_raised_focus_dots_img.Visible = False
21150         .cmdClearAll_sunken_focus_dots_img.Visible = False
21160       End If
21170     ElseIf .cmdClearAll.Enabled = False Then
21180       .cmdClearAll.Enabled = True
21190       .cmdClearAll_raised_img.Visible = True
21200       .cmdClearAll_raised_semifocus_dots_img.Visible = False
21210       .cmdClearAll_raised_focus_img.Visible = False
21220       .cmdClearAll_raised_focus_dots_img.Visible = False
21230       .cmdClearAll_sunken_focus_dots_img.Visible = False
21240       .cmdClearAll_raised_img_dis.Visible = False
21250     End If

21260   End With

EXITP:
21270   Exit Sub

ERRH:
21280   Select Case ERR.Number
        Case Else
21290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21300   End Select
21310   Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

21400 On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
21410   intShiftDown = (Shift And acShiftMask) > 0
21420   intAltDown = (Shift And acAltMask) > 0
21430   intCtrlDown = (Shift And acCtrlMask) > 0

21440   intRetVal = KeyCode

        ' ** Plain keys.
21450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21470       With Me
21480         intRetVal = 0
21490         .DateEnd.SetFocus
21500       End With
21510     End Select
21520   End If

        ' ** Shift keys.
21530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21550       With Me
21560         intRetVal = 0
21570         .chkShowJournalNo.SetFocus
21580       End With
21590     End Select
21600   End If

        ' ** Alt keys.
21610   If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
21620     Select Case intRetVal
          Case vbKeyD  ' ** I don't know why the Form_KeyDown() doesn't catch this.
21630       With Me
21640         intRetVal = 0
21650         lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
21660         If lngRecsCur > 0& Then
21670           .frmAccountHideTrans2_One_Sub_List.SetFocus
21680           .frmAccountHideTrans2_One_Sub_List.Form.ledger_HIDDEN.SetFocus
21690         Else
21700           Beep
21710         End If
21720       End With
21730     Case vbKeyN
21740       With Me
21750         intRetVal = 0
21760         lngRecsCur = .frmAccountHideTrans2_One_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_Pick.
21770         If lngRecsCur > 0& Then
21780           .frmAccountHideTrans2_One_Sub_Pick.SetFocus
21790           .frmAccountHideTrans2_One_Sub_Pick.Form.ledger_HIDDEN.SetFocus
21800         Else
21810           Beep
21820         End If
21830       End With
21840     End Select
21850   End If

EXITP:
21860   KeyCode = intRetVal
21870   Exit Sub

ERRH:
21880   intRetVal = 0
21890   Select Case ERR.Number
        Case Else
21900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21910   End Select
21920   Resume EXITP

End Sub

Private Sub DateEnd_BeforeUpdate(Cancel As Integer)

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_BeforeUpdate"

22010   With Me
22020     If IsNull(.DateEnd) = False Then
22030       If IsDate(.DateEnd) = True Then
              ' ** Continue.
22040       Else
22050         Beep
22060         MsgBox "Please enter a valid date." & vbCrLf & vbCrLf & vbCrLf & "msg 4", vbInformation + vbOKOnly, "Invalid Date"
22070         Cancel = -1
22080       End If
22090     End If
22100   End With

EXITP:
22110   Exit Sub

ERRH:
22120   Select Case ERR.Number
        Case Else
22130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22140   End Select
22150   Resume EXITP

End Sub

Private Sub DateEnd_AfterUpdate()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_AfterUpdate"

        Dim strFilter As String
        Dim datStartDate As Date, datEndDate As Date
        Dim blnContinue As Boolean
        Dim strTmp01 As String

22210   With Me
22220     blnContinue = True
22230     strFilter = vbNullString: strTmp01 = vbNullString

22240     If IsNull(.DateEnd) = False Then
22250       Select Case IsDate(.DateEnd)
            Case True
22260         datEndDate = .DateEnd
22270         If IsNull(.DateStart) = False Then
22280           datStartDate = .DateStart
22290           If datStartDate <= datEndDate Then
22300             strTmp01 = "[transdate] >= #" & Format(datStartDate, "mm/dd/yyyy") & "# And " & _
                    "[transdate] <= #" & Format(datEndDate, "mm/dd/yyyy") & "#"
22310           Else
22320             blnContinue = False
22330             Beep
22340             MsgBox "'To' date must be greater than or equal to 'From' Date.", vbInformation + vbOKOnly, "Invalid Entry"
22350             .DateEnd.Undo
22360             .DateEnd.SetFocus
22370           End If
22380         Else
22390           strTmp01 = "[transdate] <= #" & Format(datEndDate, "mm/dd/yyyy") & "#"
22400         End If
22410       Case False
22420         blnContinue = False
22430         Beep
22440         MsgBox "Please enter a valid date." & vbCrLf & vbCrLf & vbCrLf & "msg 5", vbInformation + vbOKOnly, "Invalid Date"
22450         .DateEnd.Undo
22460         .DateEnd.SetFocus
22470       End Select
22480     Else
22490       If IsNull(.DateStart) = False Then
22500         datStartDate = .DateStart
22510         strTmp01 = "[transdate] >= #" & Format(datStartDate, "mm/dd/yyyy") & "#"
22520       End If
22530     End If

22540     If blnContinue = True Then
22550       If strTmp01 <> vbNullString Then
22560         strFilter = strTmp01
22570       End If
22580       If IsNull(.cmbJournalTypes) = False Then
22590         If strFilter <> vbNullString Then strFilter = strFilter & " And "
22600         strFilter = strFilter & "[journaltype] = '" & .cmbJournalTypes & "'"
22610       End If
22620       If IsNull(.cmbAssets) = False Then
22630         If strFilter <> vbNullString Then strFilter = strFilter & " And "
22640         strFilter = strFilter & "[assetno] = " & CStr(.cmbAssets)
22650       End If
22660       If .chkIncludeArchive.Enabled = True Then
22670         If .chkIncludeArchive = False Then
22680           If strFilter <> vbNullString Then strFilter = strFilter & " And "
22690           strFilter = strFilter & "[src] <> 'A'"
22700         End If
22710       End If
22720       .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
22730       If strFilter = vbNullString Then
22740         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False
22750       Else
22760         .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = True
22770       End If
22780       .frmAccountHideTrans2_One_Sub_List.Form.Requery
22790     End If  ' ** blnContinue.

22800     If IsNull(.DateStart) = True And IsNull(.DateEnd) = True And IsNull(.cmbJournalTypes) = True And IsNull(.cmbAssets) = True Then
22810       If .cmdClearAll.Enabled = True Then
22820         .cmdClearAll.Enabled = False
22830         .cmdClearAll_raised_img_dis.Visible = True
22840         .cmdClearAll_raised_img.Visible = False
22850         .cmdClearAll_raised_semifocus_dots_img.Visible = False
22860         .cmdClearAll_raised_focus_img.Visible = False
22870         .cmdClearAll_raised_focus_dots_img.Visible = False
22880         .cmdClearAll_sunken_focus_dots_img.Visible = False
22890       End If
22900     ElseIf .cmdClearAll.Enabled = False Then
22910       .cmdClearAll.Enabled = True
22920       .cmdClearAll_raised_img.Visible = True
22930       .cmdClearAll_raised_semifocus_dots_img.Visible = False
22940       .cmdClearAll_raised_focus_img.Visible = False
22950       .cmdClearAll_raised_focus_dots_img.Visible = False
22960       .cmdClearAll_sunken_focus_dots_img.Visible = False
22970       .cmdClearAll_raised_img_dis.Visible = False
22980     End If

22990   End With

EXITP:
23000   Exit Sub

ERRH:
23010   Select Case ERR.Number
        Case Else
23020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23030   End Select
23040   Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
23110   intShiftDown = (Shift And acShiftMask) > 0
23120   intAltDown = (Shift And acAltMask) > 0
23130   intCtrlDown = (Shift And acCtrlMask) > 0

23140   intRetVal = KeyCode

        ' ** Plain keys.
23150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23170       With Me
23180         intRetVal = 0
23190         .cmbJournalTypes.SetFocus
23200       End With
23210     End Select
23220   End If

        ' ** Shift keys.
23230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23250       With Me
23260         intRetVal = 0
23270         .DateStart.SetFocus
23280       End With
23290     End Select
23300   End If

        ' ** Alt keys.
23310   If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
23320     Select Case intRetVal
          Case vbKeyD  ' ** I don't know why the Form_KeyDown() doesn't catch this.
23330       With Me
23340         intRetVal = 0
23350         lngRecsCur = .frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
23360         If lngRecsCur > 0& Then
23370           .frmAccountHideTrans2_One_Sub_List.SetFocus
23380           .frmAccountHideTrans2_One_Sub_List.Form.ledger_HIDDEN.SetFocus
23390         Else
23400           Beep
23410         End If
23420       End With
23430     Case vbKeyN
23440       With Me
23450         intRetVal = 0
23460         lngRecsCur = .frmAccountHideTrans2_One_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_Pick.
23470         If lngRecsCur > 0& Then
23480           .frmAccountHideTrans2_One_Sub_Pick.SetFocus
23490           .frmAccountHideTrans2_One_Sub_Pick.Form.ledger_HIDDEN.SetFocus
23500         Else
23510           Beep
23520         End If
23530       End With
23540     End Select
23550   End If

EXITP:
23560   KeyCode = intRetVal
23570   Exit Sub

ERRH:
23580   intRetVal = 0
23590   Select Case ERR.Number
        Case Else
23600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23610   End Select
23620   Resume EXITP

End Sub

Private Sub cmbJournalTypes_DblClick(Cancel As Integer)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalTypes_DblClick"

23710   With Me
23720     .cmbJournalTypes = Null
23730     cmbJournalTypes_AfterUpdate  ' ** Procedure: Below.
23740   End With

EXITP:
23750   Exit Sub

ERRH:
23760   Select Case ERR.Number
        Case Else
23770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23780   End Select
23790   Resume EXITP

End Sub

Private Sub cmbJournalTypes_AfterUpdate()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalTypes_AfterUpdate"

        Dim strFilter As String

23810   With Me
23820     strFilter = vbNullString

23830     If IsNull(.DateStart) = False Then
23840       strFilter = "[transdate] >= #" & Format(.DateStart, "mm/dd/yyyy") & "#"
23850     End If
23860     If IsNull(.DateEnd) = False Then
23870       If strFilter <> vbNullString Then strFilter = strFilter & " And "
23880       strFilter = strFilter & "[transdate] <= #" & Format(.DateEnd, "mm/dd/yyyy") & "#"
23890     End If
23900     If IsNull(.cmbJournalTypes) = False Then
23910       If strFilter <> vbNullString Then strFilter = strFilter & " And "
23920       strFilter = strFilter & "[journaltype] = '" & .cmbJournalTypes & "'"
23930     End If
23940     If IsNull(.cmbAssets) = False Then
23950       If strFilter <> vbNullString Then strFilter = strFilter & " And "
23960       strFilter = strFilter & "[assetno] = " & CStr(.cmbAssets)
23970     End If
23980     If .chkIncludeArchive.Enabled = True Then
23990       If .chkIncludeArchive = False Then
24000         If strFilter <> vbNullString Then strFilter = strFilter & " And "
24010         strFilter = strFilter & "[src] <> 'A'"
24020       End If
24030     End If

24040     .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
24050     If strFilter = vbNullString Then
24060       .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False
24070     Else
24080       .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = True
24090     End If
24100     .frmAccountHideTrans2_One_Sub_List.Form.Requery

24110     If IsNull(.DateStart) = True And IsNull(.DateEnd) = True And IsNull(.cmbJournalTypes) = True And IsNull(.cmbAssets) = True Then
24120       If .cmdClearAll.Enabled = True Then
24130         .cmdClearAll.Enabled = False
24140         .cmdClearAll_raised_img_dis.Visible = True
24150         .cmdClearAll_raised_img.Visible = False
24160         .cmdClearAll_raised_semifocus_dots_img.Visible = False
24170         .cmdClearAll_raised_focus_img.Visible = False
24180         .cmdClearAll_raised_focus_dots_img.Visible = False
24190         .cmdClearAll_sunken_focus_dots_img.Visible = False
24200       End If
24210     ElseIf .cmdClearAll.Enabled = False Then
24220       .cmdClearAll.Enabled = True
24230       .cmdClearAll_raised_img.Visible = True
24240       .cmdClearAll_raised_semifocus_dots_img.Visible = False
24250       .cmdClearAll_raised_focus_img.Visible = False
24260       .cmdClearAll_raised_focus_dots_img.Visible = False
24270       .cmdClearAll_sunken_focus_dots_img.Visible = False
24280       .cmdClearAll_raised_img_dis.Visible = False
24290     End If

24300   End With

EXITP:
24310   Exit Sub

ERRH:
24320   Select Case ERR.Number
        Case Else
24330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24340   End Select
24350   Resume EXITP

End Sub

Private Sub cmbJournalTypes_KeyDown(KeyCode As Integer, Shift As Integer)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalTypes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
24410   intShiftDown = (Shift And acShiftMask) > 0
24420   intAltDown = (Shift And acAltMask) > 0
24430   intCtrlDown = (Shift And acCtrlMask) > 0

24440   intRetVal = KeyCode

        ' ** Plain keys.
24450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24470       With Me
24480         intRetVal = 0
24490         .cmbAssets.SetFocus
24500       End With
24510     End Select
24520   End If

        ' ** Shift keys.
24530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24550       With Me
24560         intRetVal = 0
24570         .DateEnd.SetFocus
24580       End With
24590     End Select
24600   End If

        ' ** Ctrl keys.
24610   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
24620     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24630       With Me
24640         intRetVal = 0
24650         .opgAssetSource.SetFocus
24660       End With
24670     End Select
24680   End If

EXITP:
24690   KeyCode = intRetVal
24700   Exit Sub

ERRH:
24710   intRetVal = 0
24720   Select Case ERR.Number
        Case Else
24730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24740   End Select
24750   Resume EXITP

End Sub

Private Sub opgAssetSource_AfterUpdate()

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_AfterUpdate"

24810   With Me
24820     Select Case .opgAssetSource
          Case .opgAssetSource_optType.OptionValue
24830       .opgAssetSource_optType_lbl.FontBold = True
24840       .opgAssetSource_optName_lbl.FontBold = False
24850       .opgAssetSource_optCusip_lbl.FontBold = False
24860       If .cmbAssets.RowSource <> "qryMasterAsset_01_01" Then
24870         .cmbAssets.RowSource = "qryMasterAsset_01_01"
24880       End If
24890     Case .opgAssetSource_optName.OptionValue
24900       .opgAssetSource_optType_lbl.FontBold = False
24910       .opgAssetSource_optName_lbl.FontBold = True
24920       .opgAssetSource_optCusip_lbl.FontBold = False
24930       If .cmbAssets.RowSource <> "qryMasterAsset_01_02" Then
24940         .cmbAssets.RowSource = "qryMasterAsset_01_02"
24950       End If
24960     Case .opgAssetSource_optCusip.OptionValue
24970       .opgAssetSource_optType_lbl.FontBold = False
24980       .opgAssetSource_optName_lbl.FontBold = False
24990       .opgAssetSource_optCusip_lbl.FontBold = True
25000       If .cmbAssets.RowSource <> "qryMasterAsset_01_03" Then
25010         .cmbAssets.RowSource = "qryMasterAsset_01_03"
25020       End If
25030     End Select
25040   End With

EXITP:
25050   Exit Sub

ERRH:
25060   Select Case ERR.Number
        Case Else
25070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25080   End Select
25090   Resume EXITP

End Sub

Private Sub opgAssetSource_optType_KeyDown(KeyCode As Integer, Shift As Integer)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optType_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
25110   intShiftDown = (Shift And acShiftMask) > 0
25120   intAltDown = (Shift And acAltMask) > 0
25130   intCtrlDown = (Shift And acCtrlMask) > 0

25140   intRetVal = KeyCode

        ' ** Plain keys.
25150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25170       With Me
25180         intRetVal = 0
25190         .cmbAssets.SetFocus
25200       End With
25210     End Select
25220   End If

        ' ** Shift keys.
25230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25250       With Me
25260         intRetVal = 0
25270         .cmbJournalTypes.SetFocus
25280       End With
25290     End Select
25300   End If

EXITP:
25310   KeyCode = intRetVal
25320   Exit Sub

ERRH:
25330   intRetVal = 0
25340   Select Case ERR.Number
        Case Else
25350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25360   End Select
25370   Resume EXITP

End Sub

Private Sub opgAssetSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
25410   intShiftDown = (Shift And acShiftMask) > 0
25420   intAltDown = (Shift And acAltMask) > 0
25430   intCtrlDown = (Shift And acCtrlMask) > 0

25440   intRetVal = KeyCode

        ' ** Plain keys.
25450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25470       With Me
25480         intRetVal = 0
25490         .cmbAssets.SetFocus
25500       End With
25510     End Select
25520   End If

        ' ** Shift keys.
25530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25550       With Me
25560         intRetVal = 0
25570         .cmbJournalTypes.SetFocus
25580       End With
25590     End Select
25600   End If

EXITP:
25610   KeyCode = intRetVal
25620   Exit Sub

ERRH:
25630   intRetVal = 0
25640   Select Case ERR.Number
        Case Else
25650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25660   End Select
25670   Resume EXITP

End Sub

Private Sub opgAssetSource_optCusip_KeyDown(KeyCode As Integer, Shift As Integer)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssetSource_optCusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
25710   intShiftDown = (Shift And acShiftMask) > 0
25720   intAltDown = (Shift And acAltMask) > 0
25730   intCtrlDown = (Shift And acCtrlMask) > 0

25740   intRetVal = KeyCode

        ' ** Plain keys.
25750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25770       With Me
25780         intRetVal = 0
25790         .cmbAssets.SetFocus
25800       End With
25810     End Select
25820   End If

        ' ** Shift keys.
25830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25850       With Me
25860         intRetVal = 0
25870         .cmbJournalTypes.SetFocus
25880       End With
25890     End Select
25900   End If

EXITP:
25910   KeyCode = intRetVal
25920   Exit Sub

ERRH:
25930   intRetVal = 0
25940   Select Case ERR.Number
        Case Else
25950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25960   End Select
25970   Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

        Dim strFilter As String

26010   With Me
26020     strFilter = vbNullString

26030     If IsNull(.DateStart) = False Then
26040       strFilter = "[transdate] >= #" & Format(.DateStart, "mm/dd/yyyy") & "#"
26050     End If
26060     If IsNull(.DateEnd) = False Then
26070       If strFilter <> vbNullString Then strFilter = strFilter & " And "
26080       strFilter = strFilter & "[transdate] <= #" & Format(.DateEnd, "mm/dd/yyyy") & "#"
26090     End If
26100     If IsNull(.cmbJournalTypes) = False Then
26110       If strFilter <> vbNullString Then strFilter = strFilter & " And "
26120       strFilter = strFilter & "[journaltype] = '" & .cmbJournalTypes & "'"
26130     End If
26140     If IsNull(.cmbAssets) = False Then
26150       If strFilter <> vbNullString Then strFilter = strFilter & " And "
26160       strFilter = strFilter & "[assetno] = " & CStr(.cmbAssets)
26170     End If
26180     If .chkIncludeArchive.Enabled = True Then
26190       If .chkIncludeArchive = False Then
26200         If strFilter <> vbNullString Then strFilter = strFilter & " And "
26210         strFilter = strFilter & "[src] <> 'A'"
26220       End If
26230     End If

26240     .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
26250     If strFilter = vbNullString Then
26260       .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False
26270     Else
26280       .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = True
26290     End If
26300     .frmAccountHideTrans2_One_Sub_List.Form.Requery

26310     If IsNull(.DateStart) = True And IsNull(.DateEnd) = True And IsNull(.cmbJournalTypes) = True And IsNull(.cmbAssets) = True Then
26320       If .cmdClearAll.Enabled = True Then
26330         .cmdClearAll.Enabled = False
26340         .cmdClearAll_raised_img_dis.Visible = True
26350         .cmdClearAll_raised_img.Visible = False
26360         .cmdClearAll_raised_semifocus_dots_img.Visible = False
26370         .cmdClearAll_raised_focus_img.Visible = False
26380         .cmdClearAll_raised_focus_dots_img.Visible = False
26390         .cmdClearAll_sunken_focus_dots_img.Visible = False
26400       End If
26410     ElseIf .cmdClearAll.Enabled = False Then
26420       .cmdClearAll.Enabled = True
26430       .cmdClearAll_raised_img.Visible = True
26440       .cmdClearAll_raised_semifocus_dots_img.Visible = False
26450       .cmdClearAll_raised_focus_img.Visible = False
26460       .cmdClearAll_raised_focus_dots_img.Visible = False
26470       .cmdClearAll_sunken_focus_dots_img.Visible = False
26480       .cmdClearAll_raised_img_dis.Visible = False
26490     End If

26500   End With

EXITP:
26510   Exit Sub

ERRH:
26520   Select Case ERR.Number
        Case Else
26530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26540   End Select
26550   Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
26610   intShiftDown = (Shift And acShiftMask) > 0
26620   intAltDown = (Shift And acAltMask) > 0
26630   intCtrlDown = (Shift And acCtrlMask) > 0

26640   intRetVal = KeyCode

        ' ** Plain keys.
26650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26670       With Me
26680         intRetVal = 0
26690         If .cmdClearAll.Enabled = True Then
26700           .cmdClearAll.SetFocus
26710         Else
26720           .cmdClose.SetFocus
26730         End If
26740       End With
26750     End Select
26760   End If

        ' ** Shift keys.
26770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26790       With Me
26800         intRetVal = 0
26810         .cmbJournalTypes.SetFocus
26820       End With
26830     End Select
26840   End If

        ' ** Ctrl-Shift keys.
26850   If intCtrlDown And (Not intAltDown) And intShiftDown Then
26860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26870       With Me
26880         intRetVal = 0
26890         .opgAssetSource.SetFocus
26900       End With
26910     End Select
26920   End If

EXITP:
26930   KeyCode = intRetVal
26940   Exit Sub

ERRH:
26950   intRetVal = 0
26960   Select Case ERR.Number
        Case Else
26970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26980   End Select
26990   Resume EXITP

End Sub

Private Sub cmdClearAll_GotFocus()

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_GotFocus"

27010   With Me
27020     blnClearAll_Focus = True
27030     .cmdClearAll_raised_semifocus_dots_img.Visible = True
27040     .cmdClearAll_raised_img.Visible = False
27050     .cmdClearAll_raised_focus_img.Visible = False
27060     .cmdClearAll_raised_focus_dots_img.Visible = False
27070     .cmdClearAll_sunken_focus_dots_img.Visible = False
27080     .cmdClearAll_raised_img_dis.Visible = False
27090   End With

EXITP:
27100   Exit Sub

ERRH:
27110   Select Case ERR.Number
        Case Else
27120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27130   End Select
27140   Resume EXITP

End Sub

Private Sub cmdClearAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_MouseDown"

27210   With Me
27220     blnClearAll_MouseDown = True
27230     .cmdClearAll_sunken_focus_dots_img.Visible = True
27240     .cmdClearAll_raised_img.Visible = False
27250     .cmdClearAll_raised_semifocus_dots_img.Visible = False
27260     .cmdClearAll_raised_focus_img.Visible = False
27270     .cmdClearAll_raised_focus_dots_img.Visible = False
27280     .cmdClearAll_raised_img_dis.Visible = False
27290   End With

EXITP:
27300   Exit Sub

ERRH:
27310   Select Case ERR.Number
        Case Else
27320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27330   End Select
27340   Resume EXITP

End Sub

Private Sub cmdClearAll_Click()

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_Click"

        Dim strFilter As String

27410   With Me
27420     .DateStart = Null
27430     .DateEnd = Null
27440     .cmbJournalTypes = Null
27450     .cmbAssets = Null
27460     strFilter = vbNullString
27470     If .chkIncludeArchive.Enabled = True Then
27480       If .chkIncludeArchive = False Then
27490         strFilter = "[src] <> 'A'"
27500       End If
27510     End If
27520     .frmAccountHideTrans2_One_Sub_List.Form.Filter = strFilter
27530     If strFilter = vbNullString Then
27540       .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = False
27550     Else
27560       .frmAccountHideTrans2_One_Sub_List.Form.FilterOn = True
27570     End If
27580     .frmAccountHideTrans2_One_Sub_List.Form.Requery
27590     .DateStart.SetFocus
27600     DoEvents
27610     .cmdClearAll.Enabled = False
27620     .cmdClearAll_raised_img_dis.Visible = True
27630     .cmdClearAll_raised_img.Visible = False
27640     .cmdClearAll_raised_semifocus_dots_img.Visible = False
27650     .cmdClearAll_raised_focus_img.Visible = False
27660     .cmdClearAll_raised_focus_dots_img.Visible = False
27670     .cmdClearAll_sunken_focus_dots_img.Visible = False
27680   End With

EXITP:
27690   Exit Sub

ERRH:
27700   Select Case ERR.Number
        Case Else
27710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27720   End Select
27730   Resume EXITP

End Sub

Private Sub cmdClearAll_KeyDown(KeyCode As Integer, Shift As Integer)

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
27810   intShiftDown = (Shift And acShiftMask) > 0
27820   intAltDown = (Shift And acAltMask) > 0
27830   intCtrlDown = (Shift And acCtrlMask) > 0

27840   intRetVal = KeyCode

        ' ** Plain keys.
27850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27860     Select Case intRetVal
          Case vbKeyTab
27870       With Me
27880         intRetVal = 0
27890         .cmdClose.SetFocus
27900       End With
27910     End Select
27920   End If

        ' ** Shift keys.
27930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27940     Select Case intRetVal
          Case vbKeyTab
27950       With Me
27960         intRetVal = 0
27970         .cmbAssets.SetFocus
27980       End With
27990     End Select
28000   End If

EXITP:
28010   KeyCode = intRetVal
28020   Exit Sub

ERRH:
28030   intRetVal = 0
28040   Select Case ERR.Number
        Case Else
28050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28060   End Select
28070   Resume EXITP

End Sub

Private Sub cmdClearAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_MouseMove"

28110   With Me
28120     If blnClearAll_MouseDown = False Then
28130       Select Case blnClearAll_Focus
            Case True
28140         .cmdClearAll_raised_focus_dots_img.Visible = True
28150         .cmdClearAll_raised_focus_img.Visible = False
28160       Case False
28170         .cmdClearAll_raised_focus_img.Visible = True
28180         .cmdClearAll_raised_focus_dots_img.Visible = False
28190       End Select
28200       .cmdClearAll_raised_img.Visible = False
28210       .cmdClearAll_raised_semifocus_dots_img.Visible = False
28220       .cmdClearAll_sunken_focus_dots_img.Visible = False
28230       .cmdClearAll_raised_img_dis.Visible = False
28240     End If
28250   End With

EXITP:
28260   Exit Sub

ERRH:
28270   Select Case ERR.Number
        Case Else
28280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28290   End Select
28300   Resume EXITP

End Sub

Private Sub cmdClearAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_MouseUp"

28410   With Me
28420     .cmdClearAll_raised_focus_dots_img.Visible = True
28430     .cmdClearAll_raised_img.Visible = False
28440     .cmdClearAll_raised_semifocus_dots_img.Visible = False
28450     .cmdClearAll_raised_focus_img.Visible = False
28460     .cmdClearAll_sunken_focus_dots_img.Visible = False
28470     .cmdClearAll_raised_img_dis.Visible = False
28480     blnClearAll_MouseDown = False
28490   End With

EXITP:
28500   Exit Sub

ERRH:
28510   Select Case ERR.Number
        Case Else
28520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28530   End Select
28540   Resume EXITP

End Sub

Private Sub cmdClearAll_LostFocus()

28600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_LostFocus"

28610   With Me
28620     .cmdClearAll_raised_img.Visible = True
28630     .cmdClearAll_raised_semifocus_dots_img.Visible = False
28640     .cmdClearAll_raised_focus_img.Visible = False
28650     .cmdClearAll_raised_focus_dots_img.Visible = False
28660     .cmdClearAll_sunken_focus_dots_img.Visible = False
28670     .cmdClearAll_raised_img_dis.Visible = False
28680     blnClearAll_Focus = False
28690   End With

EXITP:
28700   Exit Sub

ERRH:
28710   Select Case ERR.Number
        Case Else
28720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28730   End Select
28740   Resume EXITP

End Sub

Public Function CreateReversingEntry(lngJournalNo As Long, strSource As String) As Boolean

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "CreateReversingEntry"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngJNo As Long, strJType As String, lngAssetNo As Long
        Dim dblICash As Double, dblPCash As Double, dblCost As Double, dblShareface As Double, dblRate As Double, dblPerShare As Double
        Dim datTransDate As Date
        Dim lngTaxcode As Long, lngRevcodeID As Long, lngLocationID As Long, lngCurrID As Long
        Dim strDescription As String, strJournal_User As String, strTotDesc As String
        Dim strCurrCode As String, strCurrSym As String
        Dim blnLedger_Hidden As Boolean, blnCheckPaid As Boolean, blnFound As Boolean, blnIsAverage As Boolean
        Dim varAssetDescription As Variant, varDue As Variant, varRecurItem As Variant
        Dim varAssetDate As Variant, varPurchaseDate As Variant, varCheckNum As Variant
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String
        Dim blnRetVal As Boolean

28810   With Me

28820     blnRetVal = True

28830     DoCmd.Hourglass True
28840     DoEvents

28850     Select Case strSource
          Case "L"
28860       varTmp00 = DMax("[journalno]", "ledger")
28870       lngJNo = (varTmp00 + 1&)
28880     Case "A"
28890       lngJNo = ArchiveJNoHoles  ' ** Function: Below.
28900       If lngJNo = 0& Then
28910         varTmp00 = DMax("[journalno]", "ledger")
28920         lngJNo = (varTmp00 + 1&)
28930       End If
28940     End Select

28950     Set dbs = CurrentDb
28960     With dbs

28970       Select Case strSource
            Case "L"
              ' ** Ledger, by specified [jno].
28980         Set qdf1 = .QueryDefs("qryAccountHideTrans2_One_01_01")
28990       Case "A"
              ' ** LedgerArchive, by specified [jno].
29000         Set qdf1 = .QueryDefs("qryAccountHideTrans2_One_01_02")
29010       End Select
29020       With qdf1.Parameters
29030         ![jno] = lngJournalNo
29040       End With
29050       Set rst1 = qdf1.OpenRecordset
29060       With rst1
29070         If .BOF = True And .EOF = True Then
                ' ** Something's horribly wrong!
29080           blnRetVal = False
29090           Beep
29100           MsgBox "EEEK!"
29110         Else
29120           .MoveFirst

29130           strAccountNo = ![accountno]
29140           lngAssetNo = ![assetno]  ' ** May be Zero.
29150           varAssetDescription = Null
29160           varDue = Null
29170           dblRate = 0#
29180           blnIsAverage = False
29190           If lngAssetNo > 0& Then
29200             varAssetDescription = DLookup("[description]", "masterasset", "[assetno] = " & CStr(lngAssetNo))
29210             varDue = DLookup("[due]", "masterasset", "[assetno] = " & CStr(lngAssetNo))
29220             varTmp00 = DLookup("[rate]", "masterasset", "[assetno] = " & CStr(lngAssetNo))
29230             Select Case IsNull(varTmp00)
                  Case True
29240               dblRate = 0#
29250             Case False
29260               dblRate = varTmp00
29270             End Select
29280           End If

29290           dblPerShare = 0#
29300           varAssetDate = Null
29310           lngTaxcode = ![taxcode]
29320           lngLocationID = ![Location_ID]
29330           varRecurItem = Null
29340           varPurchaseDate = Null
29350           blnLedger_Hidden = True
29360           lngRevcodeID = ![revcode_ID]
29370           strJournal_User = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
29380           varCheckNum = Null
29390           blnCheckPaid = False

29400           lngCurrID = ![curr_id]
29410           varTmp00 = DLookup("[curr_code]", "tblCurrency", "[curr_id] = " & CStr(lngCurrID))
29420           If IsNull(varTmp00) = False Then
29430             strCurrCode = varTmp00
29440           End If
29450           varTmp00 = DLookup("[currsym_symbol]", "tblCurrency_Symbol", "[curr_id] = " & CStr(lngCurrID))
29460           If IsNull(varTmp00) = False Then
29470             strCurrSym = varTmp00
29480           End If

29490           Select Case ![journaltype]
                Case "Dividend"
                  ' ***********************
                  ' ** Dividend
                  ' ***********************
29500             strJType = "Misc."
29510             datTransDate = ![transdate]
29520             dblShareface = 0#
29530             dblICash = (-![ICash])  ' ** Should be positive.
29540             dblPCash = 0#           ' ** Should be Zero.
29550             dblCost = 0#            ' ** Should be Zero.
29560             varRecurItem = "Dividend Reversal"
29570             strDescription = "Reversal of " & varAssetDescription & " Dividend"
29580             With Me
29590               .shareface_divint_subtot.Visible = True
29600               .shareface_subtot.Visible = False
29610             End With

29620           Case "Interest"
                  ' ***********************
                  ' ** Interest
                  ' ***********************
29630             strJType = "Misc."
29640             datTransDate = ![transdate]
29650             dblShareface = 0#
29660             dblICash = (-![ICash])  ' ** Should be positive.
29670             dblPCash = 0#           ' ** Should be Zero.
29680             dblCost = 0#            ' ** Should be Zero.
29690             varRecurItem = "Interest Reversal"
29700             strDescription = "Reversal of " & varAssetDescription & " Interest"
29710             With Me
29720               .shareface_divint_subtot.Visible = True
29730               .shareface_subtot.Visible = False
29740             End With

29750           Case "Purchase"
                  ' ***********************
                  ' ** Purchase
                  ' ***********************
29760             strJType = "Sold"
29770             datTransDate = ![transdate]
29780             dblShareface = ![shareface]
29790             dblPerShare = ![pershare]
                  ' ** ICash should be negative or Zero.
29800             If ![ICash] = 0 Then
29810               dblICash = 0#
29820             Else
29830               dblICash = Abs(![ICash])
29840             End If
                  ' ** PCash should be negative or Zero.
29850             If ![PCash] = 0 Then
29860               dblPCash = 0#
29870             Else
29880               dblPCash = Abs(![PCash])
29890             End If
                  ' ** Cost should be positive.
29900             If ![Cost] = 0 Then
29910               dblCost = 0#
29920             Else
29930               dblCost = (-![Cost])
29940             End If
29950             If dblICash = 0# And dblPCash = 0# Then
29960               strJType = "Withdrawn"
29970             End If
29980             If dblPerShare = 0# Then
29990               dblPerShare = Abs(dblCost / dblShareface)
30000             End If
                  ' ** Ideally, we want to sell the purchased Tax Lot.
                  ' ** If that's not available, then it gets messy!
30010             Set rst2 = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset, dbReadOnly)
30020             rst2.MoveFirst
30030             blnFound = False
30040             rst2.FindFirst "[assetno] = " & CStr(lngAssetNo) & " And [accountno] = '" & strAccountNo & "' And " & _
                    "[assetdate] = #" & Format(![assetdate], "mm/dd/yyyy hh:nn:ss") & "#"
30050             If rst2.NoMatch = False Then
30060               If rst2![IsAverage] = True Then
30070                 blnIsAverage = True
30080               Else
30090                 If rst2![shareface] = ![shareface] Then
                        ' ** Hooray!
30100                   blnFound = True
30110                   varPurchaseDate = rst2![assetdate]
30120                 Else
                        ' ** Rats!
30130                 End If
30140               End If
30150             Else
                    ' ** Dagnabbit!
30160             End If
30170             rst2.Close
30180             Set rst2 = Nothing
30190             If blnFound = False Then
30200               blnRetVal = False
30210               Beep
30220               Select Case blnIsAverage
                    Case True
30230                 MsgBox "The asset for the Purchase you wish to reverse" & vbCrLf & _
                        "has been specified as using Average Cost Per Unit." & vbCrLf & vbCrLf & _
                        "You will need to construct a reversal in a different manner.", _
                        vbInformation + vbOKOnly, "Average Costing"
30240               Case False
30250                 MsgBox "The original Tax Lot for the Purchase you wish to reverse" & vbCrLf & _
                        "has been partially or wholly sold via another transaction." & vbCrLf & vbCrLf & _
                        "You will need to construct a reversal in a different manner.", _
                        vbInformation + vbOKOnly, "Purchased Tax Lot Unavailable"
30260               End Select
30270             End If
30280             If blnRetVal = True Then
30290               varAssetDate = CDate(DateAdd("n", 5, ![assetdate]))  ' ** 5 minutes after Purchase.
30300               strDescription = "Reversal of " & varAssetDescription & " Purchase"
30310             End If

30320           Case "Deposit"
                  ' ***********************
                  ' ** Deposit
                  ' ***********************
30330             strJType = "Withdrawn"
30340             datTransDate = ![transdate]
30350             dblShareface = ![shareface]
30360             dblPerShare = ![pershare]
                  ' ** ICash should Zero.
30370             If ![ICash] = 0 Then
30380               dblICash = 0#
30390             Else
30400               dblICash = Abs(![ICash])
30410             End If
                  ' ** PCash should be Zero.
30420             If ![PCash] = 0 Then
30430               dblPCash = 0#
30440             Else
30450               dblPCash = Abs(![PCash])
30460             End If
                  ' ** Cost should be positive.
30470             If ![Cost] = 0 Then
30480               dblCost = 0#
30490             Else
30500               dblCost = (-![Cost])
30510             End If
30520             If dblICash <> 0# Or dblPCash <> 0# Then
30530               strJType = "Sold"
30540             End If
30550             If dblPerShare = 0# Then
30560               dblPerShare = Abs(dblCost / dblShareface)
30570             End If
                  ' ** Ideally, we want to sell the deposited Tax Lot.
                  ' ** If that's not available, then it gets messy!
30580             Set rst2 = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset, dbReadOnly)
30590             rst2.MoveFirst
30600             blnFound = False
30610             rst2.FindFirst "[assetno] = " & CStr(lngAssetNo) & " And [accountno] = '" & strAccountNo & "' And " & _
                    "[assetdate] = #" & Format(![assetdate], "mm/dd/yyyy hh:nn:ss") & "#"
30620             If rst2.NoMatch = False Then
30630               If rst2![IsAverage] = True Then
30640                 blnIsAverage = True
30650               Else
30660                 If rst2![shareface] = ![shareface] Then
                        ' ** Hooray!
30670                   blnFound = True
30680                   varPurchaseDate = rst2![assetdate]
30690                 Else
                        ' ** Rats!
30700                 End If
30710               End If
30720             Else
                    ' ** Dagnabbit!
30730             End If
30740             rst2.Close
30750             Set rst2 = Nothing
30760             If blnFound = False Then
30770               blnRetVal = False
30780               Beep
30790               Select Case blnIsAverage
                    Case True
30800                 MsgBox "The asset for the Deposit you wish to reverse" & vbCrLf & _
                        "has been specified as using Average Cost Per Unit." & vbCrLf & vbCrLf & _
                        "You will need to construct a reversal in a different manner.", _
                        vbInformation + vbOKOnly, "Average Costing"
30810               Case False
30820                 MsgBox "The original Tax Lot for the Deposit you wish to reverse" & vbCrLf & _
                        "has been partially or wholly sold via another transaction." & vbCrLf & vbCrLf & _
                        "You will need to construct a reversal in a different manner.", _
                        vbInformation + vbOKOnly, "Deposited Tax Lot Unavailable"
30830               End Select
30840             End If
30850             If blnRetVal = True Then
30860               varAssetDate = CDate(DateAdd("n", 5, ![assetdate]))  ' ** 5 minutes after Purchase.
30870               strDescription = "Reversal of " & varAssetDescription & " Deposit"
30880             End If

30890           Case "Sold"
                  ' ***********************
                  ' ** Sold
                  ' ***********************
30900             strJType = "Purchase"
30910             datTransDate = ![transdate]
30920             dblShareface = ![shareface]
30930             dblPerShare = ![pershare]
                  ' ** ICash should be positive or Zero.
30940             If ![ICash] = 0 Then
30950               dblICash = 0#
30960             Else
30970               dblICash = (-![ICash])
30980             End If
                  ' ** PCash should be positive or Zero.
30990             If ![PCash] = 0 Then
31000               dblPCash = 0#
31010             Else
31020               dblPCash = (-![PCash])
31030             End If
                  ' ** Cost should be negative.
31040             If ![Cost] = 0 Then
31050               dblCost = 0#
31060             Else
31070               dblCost = Abs(![Cost])
31080             End If
31090             If dblICash = 0# And dblPCash = 0# Then
31100               strJType = "Deposit"
31110             End If
31120             If dblPerShare = 0# Then
31130               dblPerShare = Abs(dblCost / dblShareface)
31140             End If
                  ' ** ActiveAssets, grouped, by accountno, assetno, just
                  ' ** IsAverage = True, by specified [actno], [astno].
31150             Set qdf2 = dbs.QueryDefs("qryAccountHideTrans2_One_11")
31160             With qdf2.Parameters
31170               ![actno] = strAccountNo
31180               ![astno] = lngAssetNo
31190             End With
31200             Set rst2 = qdf2.OpenRecordset
31210             With rst2
31220               If .BOF = True And .EOF = True Then
                      ' ** No evidence of averaging.
31230               Else
31240                 .MoveFirst
31250                 blnIsAverage = True
31260               End If
31270               .Close
31280             End With
31290             Set rst2 = Nothing
31300             Set qdf2 = Nothing
31310             Select Case blnIsAverage
                  Case True
31320               blnRetVal = False
31330               Beep
31340               MsgBox "The asset for the Sold you wish to reverse" & vbCrLf & _
                      "has been specified as using Average Cost Per Unit." & vbCrLf & vbCrLf & _
                      "You will need to construct a reversal in a different manner.", _
                      vbInformation + vbOKOnly, "Average Costing"
31350             Case False
                    ' ** Whether or not the Tax Lot still exists,
                    ' ** this Purchase will have to create a new one.
31360               varAssetDate = CDate(DateAdd("n", 5, ![PurchaseDate]))  ' ** 5 minutes after Sold.
31370               strDescription = "Reversal of " & varAssetDescription & " Sold"
31380             End Select

31390           Case "Withdrawn"
                  ' ***********************
                  ' ** Withdrawn
                  ' ***********************
31400             strJType = "Deposit"
31410             datTransDate = ![transdate]
31420             dblShareface = ![shareface]
31430             dblPerShare = ![pershare]
                  ' ** ICash should be Zero.
31440             If ![ICash] = 0 Then
31450               dblICash = 0#
31460             Else
31470               dblICash = (-![ICash])
31480             End If
                  ' ** PCash should be Zero.
31490             If ![PCash] = 0 Then
31500               dblPCash = 0#
31510             Else
31520               dblPCash = (-![PCash])
31530             End If
                  ' ** Cost should be positive.
31540             If ![Cost] = 0 Then
31550               dblCost = 0#
31560             Else
31570               dblCost = Abs(![Cost])
31580             End If
31590             If dblICash <> 0# Or dblPCash <> 0# Then
31600               strJType = "Purchase"
31610             End If
31620             If dblPerShare = 0# Then
31630               dblPerShare = Abs(dblCost / dblShareface)
31640             End If
                  ' ** ActiveAssets, grouped, by accountno, assetno, just
                  ' ** IsAverage = True, by specified [actno], [astno].
31650             Set qdf2 = dbs.QueryDefs("qryAccountHideTrans2_One_11")
31660             With qdf2.Parameters
31670               ![actno] = strAccountNo
31680               ![astno] = lngAssetNo
31690             End With
31700             Set rst2 = qdf2.OpenRecordset
31710             With rst2
31720               If .BOF = True And .EOF = True Then
                      ' ** No evidence of averaging.
31730               Else
31740                 .MoveFirst
31750                 blnIsAverage = True
31760               End If
31770               .Close
31780             End With
31790             Set rst2 = Nothing
31800             Set qdf2 = Nothing
31810             Select Case blnIsAverage
                  Case True
31820               blnRetVal = False
31830               Beep
31840               MsgBox "The asset for the Withdrawn you wish to reverse" & vbCrLf & _
                      "has been specified as using Average Cost Per Unit." & vbCrLf & vbCrLf & _
                      "You will need to construct a reversal in a different manner.", _
                      vbInformation + vbOKOnly, "Average Costing"
31850             Case False
                    ' ** Whether or not the Tax Lot still exists,
                    ' ** this Deposit will have to create a new one.
31860               varAssetDate = CDate(DateAdd("n", 5, ![PurchaseDate]))  ' ** 5 minutes after Withdrawn.
31870               strDescription = "Reversal of " & varAssetDescription & " Withdrawn"
31880             End Select

31890           Case "Misc."
                  ' ***********************
                  ' ** Misc.
                  ' ***********************
31900             strJType = "Misc."
31910             datTransDate = ![transdate]
31920             dblShareface = 0#
31930             If ![RecurringItem] = RECUR_I_TO_P Then
31940               varRecurItem = RECUR_P_TO_I
                    ' ** ICash should be negative.
31950               dblICash = Abs(![ICash])
                    ' ** PCash should be positive.
31960               dblPCash = (-![PCash])
                    ' ** Cost should be Zero.
31970               dblCost = 0#
31980             ElseIf ![RecurringItem] = RECUR_P_TO_I Then
31990               varRecurItem = RECUR_I_TO_P
                    ' ** ICash should be positive.
32000               dblICash = (-![ICash])
                    ' ** PCash should be negative.
32010               dblPCash = Abs(![PCash])
                    ' ** Cost should be Zero.
32020               dblCost = 0#
32030             Else
32040               If ![ICash] = 0 Then
32050                 dblICash = 0#
32060               ElseIf ![ICash] > 0 Then
32070                 dblICash = (-![ICash])
32080               ElseIf ![ICash] < 0 Then
32090                 dblICash = Abs(![ICash])
32100               End If
32110               If ![PCash] = 0 Then
32120                 dblPCash = 0#
32130               ElseIf ![PCash] > 0 Then
32140                 dblPCash = (-![PCash])
32150               ElseIf ![PCash] < 0 Then
32160                 dblPCash = Abs(![PCash])
32170               End If
32180               dblCost = 0#
32190             End If
32200             strDescription = "Reversal of previous Misc."
32210             With Me
32220               .shareface_divint_subtot.Visible = True
32230               .shareface_subtot.Visible = False
32240             End With

32250           Case "Paid"
                  ' ***********************
                  ' ** Paid
                  ' ***********************
32260             strJType = "Misc."
32270             datTransDate = ![transdate]
32280             dblShareface = 0#
                  ' ** ICash should be negative or Zero.
32290             If ![ICash] = 0 Then
32300               dblICash = 0#
32310             Else
32320               dblICash = Abs(![ICash])
32330             End If
                  ' ** PCash should be negative or Zero.
32340             If ![PCash] = 0 Then
32350               dblPCash = 0#
32360             Else
32370               dblPCash = Abs(![PCash])
32380             End If
                  ' ** Cost should be Zero.
32390             dblCost = 0#
32400             varRecurItem = "Paid Reversal"
32410             strDescription = "Reversal of Paid to " & ![RecurringItem]
32420             With Me
32430               .shareface_divint_subtot.Visible = True
32440               .shareface_subtot.Visible = False
32450             End With

32460           Case "Received"
                  ' ***********************
                  ' ** Received
                  ' ***********************
                  ' ** Only Capital Gain gets assetno, assetdate, and shareface.
32470             strJType = "Misc."
32480             datTransDate = ![transdate]
32490             dblShareface = 0#
                  ' ** ICash should be positive or Zero.
32500             If ![ICash] = 0 Then
32510               dblICash = 0#
32520             Else
32530               dblICash = (-![ICash])
32540             End If
                  ' ** PCash should be positive or Zero.
32550             If ![PCash] = 0 Then
32560               dblPCash = 0#
32570             Else
32580               dblPCash = (-![PCash])
32590             End If
                  ' ** Cost should be Zero.
32600             dblCost = 0#
32610             varRecurItem = "Received Reversal"
32620             strDescription = "Reversal of Received from " & ![RecurringItem]
32630             With Me
32640               .shareface_divint_subtot.Visible = True
32650               .shareface_subtot.Visible = False
32660             End With

32670           Case "Liability"
                  ' ***********************
                  ' ** Liability
                  ' ***********************
                  ' ** This will assume the original transaction was entered correctly.
32680             strJType = "Liability"
32690             datTransDate = ![transdate]
32700             Select Case IsNull(![PurchaseDate])
                  Case True
                    ' ** Was 'Liability (+)'.
                    ' ** Now 'Liability (-)'.
                    ' ** ICash should be positive or Zero.
32710               If ![ICash] = 0 Then
32720                 dblICash = 0#
32730               Else
32740                 dblICash = (-![ICash])
32750               End If
                    ' ** PCash should be positive or Zero.
32760               If ![PCash] = 0 Then
32770                 dblPCash = 0#
32780               Else
32790                 dblPCash = (-![PCash])
32800               End If
                    ' ** Cost should be negative.
32810               dblCost = Abs(![Cost])
                    ' ** Again, ideally, we want to pay off the original Liability.
                    ' ** If it's not available, then it's too messy!
32820               Set rst2 = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset, dbReadOnly)
32830               rst2.MoveFirst
32840               blnFound = False
32850               rst2.FindFirst "[assetno] = " & CStr(lngAssetNo) & " And [accountno] = '" & strAccountNo & "' And " & _
                      "[assetdate] = #" & Format(![assetdate], "mm/dd/yyyy hh:nn:ss") & "#"
32860               If rst2.NoMatch = False Then
32870                 If rst2![shareface] = ![shareface] Then
                        ' ** Hooray!
32880                   blnFound = True
32890                   varPurchaseDate = rst2![assetdate]
32900                 Else
                        ' ** Rats!
32910                 End If
32920               Else
                      ' ** Dagnabbit!
32930               End If
32940               rst2.Close
32950               Set rst2 = Nothing
32960               If blnFound = False Then
32970                 blnRetVal = False
32980                 Beep
32990                 MsgBox "The original Tax Lot for the Liability you wish to reverse" & vbCrLf & _
                        "has been partially or wholly paid off via another transaction." & vbCrLf & vbCrLf & _
                        "You will need to construct a reversal in a different manner.", _
                        vbInformation + vbOKOnly, "Liability Tax Lot Unavailable"
33000               End If
33010               If blnRetVal = True Then
33020                 varAssetDate = CDate(DateAdd("n", 5, ![assetdate]))  ' ** 5 minutes after Liability.
33030                 strDescription = "Reversal of " & varAssetDescription & " Liability increase"
33040               End If
33050             Case False
                    ' ** Was 'Liability (-)'.
                    ' ** Now 'Liability (+)'.
                    ' ** ICash should be negative or Zero.
33060               If ![ICash] = 0 Then
33070                 dblICash = 0#
33080               Else
33090                 dblICash = Abs(![ICash])
33100               End If
                    ' ** PCash should be negative or Zero.
33110               If ![PCash] = 0 Then
33120                 dblPCash = 0#
33130               Else
33140                 dblPCash = Abs(![PCash])
33150               End If
                    ' ** Cost should be positive.
33160               dblCost = (-![Cost])
                    ' ** Whether or not the Tax Lot still exists,
                    ' ** this Liability will have to create a new one.
33170               varAssetDate = CDate(DateAdd("n", 5, ![PurchaseDate]))  ' ** 5 minutes after Liability.
33180               strDescription = "Reversal of " & varAssetDescription & " Liability reduction"
33190             End Select

33200           Case "Cost Adj."
                  ' ***********************
                  ' ** Cost Adj.
                  ' ***********************
33210             datTransDate = ![transdate]
33220             dblShareface = 0#
33230             blnRetVal = False
33240             Beep
33250             MsgBox "Cost Adj. reversal not implemented yet.", vbInformation + vbOKOnly, "Reversal Unavailable"

33260           End Select

33270         End If  ' ** BOF, EOF.

33280         .Close
33290       End With  ' ** rst1.
33300       Set rst1 = Nothing
33310       Set qdf1 = Nothing
33320       DoEvents

33330       If blnRetVal = True Then

33340         strTotDesc = vbNullString
33350         If lngAssetNo = 0& Then
33360           If IsNull(varRecurItem) = False Then
33370             strTotDesc = varRecurItem
33380           End If
33390         End If
33400         If lngAssetNo > 0& Then
33410           strTotDesc = strTotDesc & " " & varAssetDescription
33420           If dblRate > 0# Then
33430             strTotDesc = strTotDesc & " " & Format(dblRate, "0.000%")
33440           End If
33450           If IsNull(varDue) = False Then
33460             strTotDesc = strTotDesc & "  " & Format(varDue, "mm/dd/yyyy")
33470           End If
33480         End If
33490         strTotDesc = strTotDesc & " " & strDescription
33500         strTotDesc = Trim(strTotDesc)

              ' ** All variables should be properly populated and ready for creation of new reversing entry.
33510         Set rst1 = .OpenRecordset("tblLedgerHidden_Staging2", dbOpenDynaset, dbAppendOnly)
33520         With rst1
33530           .AddNew
                ' ** ![ledghids2_id] : AutoNumber.
33540           ![journalno] = lngJNo
33550           ![journaltype] = strJType
33560           ![assetno] = lngAssetNo
33570           ![asset_description] = varAssetDescription
33580           ![totdesc] = strTotDesc
33590           ![transdate] = datTransDate
33600           ![postdate] = Null
33610           ![accountno] = strAccountNo
33620           ![shareface] = dblShareface
33630           strTmp01 = CStr(dblShareface)
33640           intPos01 = InStr(strTmp01, ".")
33650           If intPos01 = 0 Then
33660             ![shareface_0] = dblShareface
33670           Else
33680             strTmp01 = Mid(strTmp01, (intPos01 + 1))
33690             Select Case Len(strTmp01)
                  Case 1
33700               ![shareface_1] = dblShareface
33710             Case 2
33720               ![shareface_2] = dblShareface
33730             Case 3
33740               ![shareface_3] = dblShareface
33750             Case 4
33760               ![shareface_4] = dblShareface
33770             End Select
33780           End If
33790           ![due] = varDue
33800           ![rate] = dblRate
33810           ![pershare] = dblPerShare
33820           ![ICash] = dblICash
33830           strTmp01 = Format(dblICash, strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00")
33840           If dblICash >= 0# Then
33850             strTmp01 = strTmp01 & Chr(160)  ' ** Add a hard-space for alignment.
33860           End If
33870           ![icash_str] = strTmp01
33880           ![PCash] = dblPCash
33890           strTmp01 = Format(dblPCash, strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00")
33900           If dblPCash >= 0# Then
33910             strTmp01 = strTmp01 & Chr(160)  ' ** Add a hard-space for alignment.
33920           End If
33930           ![pcash_str] = strTmp01
33940           ![Cost] = dblCost
33950           strTmp01 = Format(dblCost, strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00")
33960           If dblCost >= 0# Then
33970             strTmp01 = strTmp01 & Chr(160)  ' ** Add a hard-space for alignment.
33980           End If
33990           ![cost_str] = strTmp01
34000           ![assetdate] = varAssetDate
34010           ![ledger_description] = strDescription
34020           ![posted] = Now()
34030           ![taxcode] = lngTaxcode
34040           ![Location_ID] = lngLocationID
34050           ![RecurringItem] = varRecurItem
34060           ![PurchaseDate] = varPurchaseDate
34070           ![ledger_HIDDEN] = blnLedger_Hidden
34080           ![revcode_ID] = lngRevcodeID
34090           ![journal_USER] = strJournal_User
34100           ![CheckNum] = varCheckNum
34110           ![CheckPaid] = blnCheckPaid
34120           ![curr_id] = lngCurrID
34130           ![curr_code] = strCurrCode
34140           ![currsym_symbol] = strCurrSym
34150           ![src] = "L"
34160           ![sort] = 2&
34170           ![ledghids2_datemodified] = Now()
34180           .Update
34190           .Close
34200         End With
34210         Set rst1 = Nothing
34220         DoEvents

34230       End If  ' ** blnRetVal.

34240       .Close
34250     End With  ' ** dbs.
34260     Set dbs = Nothing

34270     Select Case blnRetVal
          Case True
34280       .frmAccountHideTrans2_One_Sub_Pick.Form.Requery

34290       .cmdOK.Enabled = True
34300       .cmdCancel.Enabled = True

34310       If .shareface_subtot.Visible = True Then
34320         .shareface_subtot_box1.Visible = True
34330         .shareface_subtot_box2.Visible = True
34340       End If
34350       .icash_subtot_box1.Visible = True
34360       .icash_subtot_box2.Visible = True
34370       .pcash_subtot_box1.Visible = True
34380       .pcash_subtot_box2.Visible = True
34390       .cost_subtot_box1.Visible = True
34400       .cost_subtot_box2.Visible = True

34410       .curr_id = lngCurrID
34420       If gblnHasForExThis = True Then
34430         .curr_id.Visible = True
34440         .curr_id_box1.Visible = True
34450         .curr_id_box2.Visible = True
34460       Else
34470         Select Case .chkIncludeCurrency
              Case True
34480           .curr_id.Visible = True
34490           .curr_id_box1.Visible = True
34500           .curr_id_box2.Visible = True
34510         Case False
34520           .curr_id.Visible = False
34530           .curr_id_box1.Visible = False
34540           .curr_id_box2.Visible = False
34550         End Select
34560       End If

34570       If .shareface_subtot.Visible = True Then
34580         .shareface_subtot.Enabled = False
34590         .shareface_subtot.Locked = True
34600         .shareface_subtot.BorderColor = CLR_LTBLU2
34610         .shareface_subtot.BackStyle = acBackStyleNormal
34620       End If
34630       .shareface_subtot_lbl.ForeColor = CLR_DKGRY3
34640       .shareface_subtot_lbl_dim_hi.Visible = False

34650       .icash_subtot.Enabled = False
34660       .icash_subtot.Locked = True
34670       .icash_subtot.Format = strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00"
34680       .icash_subtot.BorderColor = CLR_LTBLU2
34690       .icash_subtot.BackStyle = acBackStyleNormal
34700       .icash_subtot_lbl.ForeColor = CLR_DKGRY3
34710       .icash_subtot_lbl_dim_hi.Visible = False

34720       .pcash_subtot.Enabled = False
34730       .pcash_subtot.Locked = True
34740       .pcash_subtot.Format = strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00"
34750       .pcash_subtot.BorderColor = CLR_LTBLU2
34760       .pcash_subtot.BackStyle = acBackStyleNormal
34770       .pcash_subtot_lbl.ForeColor = CLR_DKGRY3
34780       .pcash_subtot_lbl_dim_hi.Visible = False

34790       .cost_subtot.Enabled = False
34800       .cost_subtot.Locked = True
34810       .cost_subtot.Format = strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00"
34820       .cost_subtot.BorderColor = CLR_LTBLU2
34830       .cost_subtot.BackStyle = acBackStyleNormal
34840       .cost_subtot_lbl.ForeColor = CLR_DKGRY3
34850       .cost_subtot_lbl_dim_hi.Visible = False

34860       Select Case .chkIncludeCurrency
            Case True
34870         .curr_id.Enabled = False
34880         .curr_id.Locked = True
34890         .curr_id.BorderColor = CLR_LTBLU2
34900         .curr_id.BackStyle = acBackStyleNormal
34910         .curr_id_lbl.ForeColor = CLR_DKGRY3
34920         .curr_id_lbl_dim_hi.Visible = False
34930         If lngCurrID = 150 Then
34940           .curr_id.BackColor = CLR_DISABLED_BG
34950           .icash_subtot.BackColor = CLR_WHT
34960           .pcash_subtot.BackColor = CLR_WHT
34970           .cost_subtot.BackColor = CLR_WHT
34980         Else
34990           .curr_id.BackColor = CLR_LTGRN
35000           .icash_subtot.BackColor = CLR_VLTGRN
35010           .pcash_subtot.BackColor = CLR_VLTGRN
35020           .cost_subtot.BackColor = CLR_VLTGRN
35030         End If
35040       Case False
35050         .curr_id_lbl.Visible = False
35060         .curr_id_lbl_dim_hi.Visible = False
35070       End Select

35080       .subtot_lbl.Visible = True
35090       .subtot_lbl_dim.Visible = False
35100       .subtot_lbl_dim_hi.Visible = False

35110       .pending_trans.Enabled = False
35120       .pending_trans.Locked = True
35130       .pending_trans.BorderColor = CLR_LTBLU2
35140       .pending_trans.BackStyle = acBackStyleNormal
35150       .pending_trans = 2

35160       DoEvents
35170       If .shareface_subtot.Visible = True Then
35180         .shareface_subtot = .frmAccountHideTrans2_One_Sub_Pick.Form.shareface_subtot
35190       Else
              ' ** It'll just be Zero if not Dividend or Interest.
35200         .shareface_subtot = .frmAccountHideTrans2_One_Sub_Pick.Form.shareface_divint_subtot
35210       End If
35220       .icash_subtot = .frmAccountHideTrans2_One_Sub_Pick.Form.icash_subtot
35230       .pcash_subtot = .frmAccountHideTrans2_One_Sub_Pick.Form.pcash_subtot
35240       .cost_subtot = .frmAccountHideTrans2_One_Sub_Pick.Form.cost_subtot

35250       .cmdOK.SetFocus

35260     Case False
35270       Set dbs = CurrentDb
35280       With dbs
              ' ** Empty tblLedgerHidden_Staging2.
35290         Set qdf1 = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
35300         qdf1.Execute
35310         Set qdf1 = Nothing
              ' ** Update tblLedgerHidden_Staging1, for ledger_HIDDEN = False.
35320         Set qdf1 = .QueryDefs("qryAccountHideTrans2_One_03")
35330         qdf1.Execute
35340         Set qdf1 = Nothing
35350         .Close
35360       End With
35370       Set dbs = Nothing
35380       DoEvents

35390       .frmAccountHideTrans2_One_Sub_Pick.Form.Requery
35400       .frmAccountHideTrans2_One_Sub_Pick.Form.LabelEnable False  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.

35410       .cmdOK.Enabled = False
35420       .cmdCancel.Enabled = False

35430       .shareface_subtot_box1.Visible = False
35440       .shareface_subtot_box2.Visible = False
35450       .icash_subtot_box1.Visible = False
35460       .icash_subtot_box2.Visible = False
35470       .pcash_subtot_box1.Visible = False
35480       .pcash_subtot_box2.Visible = False
35490       .cost_subtot_box1.Visible = False
35500       .cost_subtot_box2.Visible = False
35510       .curr_id_box1.Visible = False
35520       .curr_id_box2.Visible = False

35530       .shareface_subtot.Visible = True
35540       .shareface_divint_subtot.Visible = False
35550       .shareface_subtot.Enabled = False
35560       .shareface_subtot.Locked = False
35570       .shareface_subtot.BorderColor = WIN_CLR_DISR
35580       .shareface_subtot.BackStyle = acBackStyleTransparent
35590       .shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
35600       .shareface_subtot_lbl_dim_hi.Visible = True
35610       .icash_subtot.Enabled = False
35620       .icash_subtot.Locked = False
35630       .icash_subtot = 0
35640       .icash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
35650       .icash_subtot.BorderColor = WIN_CLR_DISR
35660       .icash_subtot.BackStyle = acBackStyleTransparent
35670       .icash_subtot_lbl.ForeColor = WIN_CLR_DISF
35680       .icash_subtot_lbl_dim_hi.Visible = True
35690       .pcash_subtot.Enabled = False
35700       .pcash_subtot.Locked = False
35710       .pcash_subtot = 0
35720       .pcash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
35730       .pcash_subtot.BorderColor = WIN_CLR_DISR
35740       .pcash_subtot.BackStyle = acBackStyleTransparent
35750       .pcash_subtot_lbl.ForeColor = WIN_CLR_DISF
35760       .pcash_subtot_lbl_dim_hi.Visible = True
35770       .cost_subtot.Enabled = False
35780       .cost_subtot.Locked = False
35790       .cost_subtot = 0
35800       .cost_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
35810       .cost_subtot.BorderColor = WIN_CLR_DISR
35820       .cost_subtot.BackStyle = acBackStyleTransparent
35830       .cost_subtot_lbl.ForeColor = WIN_CLR_DISF
35840       .cost_subtot_lbl_dim_hi.Visible = True

35850       .curr_id = Null
35860       If gblnHasForExThis = True Then
35870         .curr_id.Enabled = False
35880         .curr_id.Locked = False
35890         .curr_id.BorderColor = WIN_CLR_DISR
35900         .curr_id.BackStyle = acBackStyleTransparent
35910         .curr_id_lbl.ForeColor = WIN_CLR_DISF
35920         .curr_id_lbl_dim_hi.Visible = True
35930       Else
35940         Select Case .chkIncludeCurrency
              Case True
35950           .curr_id.Enabled = False
35960           .curr_id.Locked = False
35970           .curr_id.BorderColor = WIN_CLR_DISR
35980           .curr_id.BackStyle = acBackStyleTransparent
35990           .curr_id_lbl.ForeColor = WIN_CLR_DISF
36000           .curr_id_lbl_dim_hi.Visible = True
36010         Case False
36020           .curr_id.Visible = False
36030           .curr_id_lbl.Visible = False
36040           .curr_id_lbl_dim_hi.Visible = False
36050         End Select
36060       End If

36070       .subtot_lbl_dim.Visible = True
36080       .subtot_lbl_dim_hi.Visible = True
36090       .subtot_lbl.Visible = False

36100       .pending_trans.Enabled = False
36110       .pending_trans.Locked = False
36120       .pending_trans.BorderColor = WIN_CLR_DISR
36130       .pending_trans.BackStyle = acBackStyleTransparent
36140       .pending_trans = 0

36150     End Select  ' ** blnRetVal.

          ' ** Ledger fields:
          ' **   journalno
          ' **   journaltype
          ' **   assetno
          ' **   transdate
          ' **   postdate
          ' **   accountno
          ' **   shareface
          ' **   due
          ' **   rate
          ' **   pershare
          ' **   icash
          ' **   pcash
          ' **   cost
          ' **   assetdate
          ' **   description
          ' **   posted
          ' **   taxcode
          ' **   Location_ID
          ' **   RecurringItem
          ' **   PurchaseDate
          ' **   ledger_Hidden
          ' **   revcode_ID
          ' **   journal_USER
          ' **   CheckNum
          ' **   CheckPaid
          ' **   curr_id

36160   End With

EXITP:
36170   Set rst1 = Nothing
36180   Set rst2 = Nothing
36190   Set qdf1 = Nothing
36200   Set qdf2 = Nothing
36210   Set dbs = Nothing
36220   CreateReversingEntry = blnRetVal
36230   Exit Function

ERRH:
36240   blnRetVal = False
36250   DoCmd.Hourglass False
36260   Select Case ERR.Number
        Case Else
36270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
36280   End Select
36290   Resume EXITP

End Function

Private Sub ArchiveSet(blnHasArchs As Boolean)

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "ArchiveSet"

36310   With Me
36320     Select Case blnHasArchive
          Case True
36330       Select Case blnHasArchs
            Case True
36340         .chkIncludeArchive_lbl.Visible = True
36350         .chkIncludeArchive_lbl2_dim.Visible = False
36360         .chkIncludeArchive_lbl2_dim_hi.Visible = False
36370         .chkIncludeArchive_lbl3_dim.Visible = False
36380         .chkIncludeArchive_lbl3_dim_hi.Visible = False
36390         .chkIncludeArchive_lbl4_dim.Visible = False
36400         .chkIncludeArchive_lbl4_dim_hi.Visible = False
36410         .chkIncludeArchive_lbl5_dim.Visible = False
36420         .chkIncludeArchive_lbl5_dim_hi.Visible = False
36430       Case False
              ' ** This Account Has No Archived Transactions.
36440         .chkIncludeArchive_lbl.Visible = False
36450         .chkIncludeArchive_lbl2_dim.Visible = False
36460         .chkIncludeArchive_lbl2_dim_hi.Visible = False
36470         .chkIncludeArchive_lbl3_dim.Visible = False
36480         .chkIncludeArchive_lbl3_dim_hi.Visible = False
36490         .chkIncludeArchive_lbl4_dim.Visible = True
36500         .chkIncludeArchive_lbl4_dim_hi.Visible = True
36510         .chkIncludeArchive_lbl5_dim.Visible = True
36520         .chkIncludeArchive_lbl5_dim_hi.Visible = True
36530       End Select
36540     Case False
            ' ** There Are No Archived Transactions.
36550       .chkIncludeArchive_lbl.Visible = False
36560       .chkIncludeArchive_lbl2_dim.Visible = True
36570       .chkIncludeArchive_lbl2_dim_hi.Visible = True
36580       .chkIncludeArchive_lbl3_dim.Visible = True
36590       .chkIncludeArchive_lbl3_dim_hi.Visible = True
36600       .chkIncludeArchive_lbl4_dim.Visible = False
36610       .chkIncludeArchive_lbl4_dim_hi.Visible = False
36620       .chkIncludeArchive_lbl5_dim.Visible = False
36630       .chkIncludeArchive_lbl5_dim_hi.Visible = False
36640     End Select
36650   End With

EXITP:
36660   Exit Sub

ERRH:
36670   Select Case ERR.Number
        Case Else
36680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
36690   End Select
36700   Resume EXITP

End Sub

Private Function ArchiveJNoHoles() As Long
' ** Look for a journalno hole in LedgerArchive.
' ** This will be a simple search, without combining
' ** LedgerArchive with Ledger.
' ** It's highly unlikely that they'll reverse
' ** an archived entry.

36800 On Error GoTo ERRH

        Const THIS_PROC As String = "ArchiveJNoHoles"

        Dim dbs As DAO.Database, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngRecs As Long, lngJournalNo As Long
        Dim lngX As Long
        Dim lngRetVal As Long

36810   lngRetVal = 0&

36820   Set dbs = CurrentDb
36830   With dbs
36840     Set rst1 = .OpenRecordset("LedgerArchive", dbOpenDynaset, dbReadOnly)
36850     If rst1.BOF = True And rst1.EOF = True Then
            ' ** No archive.
36860       rst1.Close
36870       Set rst1 = Nothing
36880     Else
36890       rst1.sort = "[journalno]"
36900       Set rst2 = rst1.OpenRecordset
36910       With rst2
36920         .MoveLast
36930         lngRecs = .RecordCount
36940         .MoveFirst
36950         lngJournalNo = 0&
36960         For lngX = 1& To lngRecs
36970           If ![journalno] = (lngJournalNo + 1&) Then
                  ' ** Keep looking.
36980             lngJournalNo = ![journalno]
36990           Else
                  ' ** A hole's be found.
37000             lngRetVal = (lngJournalNo + 1&)
37010             Exit For
37020           End If
37030           If lngX < lngRecs Then .MoveNext
37040         Next
37050         .Close
37060       End With
37070       rst1.Close
37080       Set rst1 = Nothing
37090       Set rst2 = Nothing
37100       If lngRetVal > 0& Then
              ' ** Check for this journalno in Ledger.
37110         Set rst1 = .OpenRecordset("ledger", dbOpenDynaset, dbReadOnly)
37120         With rst1
37130           .MoveFirst
37140           .FindFirst "[journalno] = " & CStr(lngRetVal)
37150           Select Case .NoMatch
                Case True
                  ' ** Not found, so we're good!
37160           Case False
                  ' ** It's in Ledger, so no go.
37170             lngRetVal = 0&
37180           End Select
37190           .Close
37200         End With
37210         Set rst1 = Nothing
37220       Else
              ' ** See if there's anything missing between LedgerArchive and Ledger.
              ' ** lngJournalNo is the last entry in LedgerArchive.
37230         Set rst1 = .OpenRecordset("ledger", dbOpenDynaset, dbReadOnly)
37240         With rst1
37250           .MoveFirst
37260           .FindFirst "[journalno] = " & CStr(lngJournalNo + 1&)
37270           Select Case .NoMatch
                Case True
                  ' ** A hole between the tables
37280             lngRetVal = (lngJournalNo + 1&)
37290           Case False
                  ' ** No holes.
37300           End Select
37310           .Close
37320         End With
37330         Set rst1 = Nothing
37340       End If
37350     End If
37360     .Close
37370   End With

EXITP:
37380   Set rst1 = Nothing
37390   Set rst2 = Nothing
37400   Set dbs = Nothing
37410   ArchiveJNoHoles = lngRetVal
37420   Exit Function

ERRH:
37430   lngRetVal = 0&
37440   DoCmd.Hourglass False
37450   Select Case ERR.Number
        Case Else
37460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37470   End Select
37480   Resume EXITP

End Function

Private Sub ShowSub(blnShow As Boolean)

37500 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowSub"

        Dim frm As Access.Form, ctl As Access.Control
        Dim lngX As Long, lngE As Long

37510   With Me
37520     Set frm = .frmAccountHideTrans2_One_Sub_List.Form
37530     If lngItems = 0& Then
37540       ReDim arr_varItem(I_ELEMS, 0)
37550       For Each ctl In frm.Detail.Controls
37560         With ctl
37570           If .ControlType = acTextBox Then
37580             If Left(.Name, 5) = "icash" Or Left(.Name, 5) = "pcash" Or Left(.Name, 4) = "cost" Then
37590               lngItems = lngItems + 1&
37600               lngE = lngItems - 1&
37610               ReDim Preserve arr_varItem(I_ELEMS, lngE)
37620               arr_varItem(I_CNAM, lngE) = .Name
37630               arr_varItem(I_VIS, lngE) = .Visible
37640             End If
37650           End If
37660         End With
37670       Next
37680       Set ctl = Nothing
37690     End If

37700     Select Case blnShow
          Case True
37710       If .chkShowJournalNo = True Then
37720         With frm
37730           .journalno.Visible = True
37740           If blnNoTrans = False Then
37750             .journalno_lbl.ForeColor = CLR_GRY4
37760             .journalno_lbl2.ForeColor = CLR_GRY4
37770             .journalno_lbl_line.BorderColor = CLR_GRY5
37780             .journalno_lbl_dim_hi.Visible = False
37790             .journalno_lbl2_dim_hi.Visible = False
37800             .journalno_lbl_line_dim_hi.Visible = False
37810           End If
37820         End With
37830       End If
37840       If .chkIncludeCurrency = True Then
37850         With frm
37860           .curr_id.Visible = True
37870           .curr_id_box.Visible = True
37880           .curr_id_forex.Visible = True
37890           If blnNoTrans = False Then
37900             .curr_id_lbl.ForeColor = CLR_DKGRY2
37910             .curr_id_lbl_line.BorderColor = CLR_DKGRY
37920             .curr_id_lbl_dim_hi.Visible = False
37930             .curr_id_lbl_line_dim_hi.Visible = False
37940           End If
37950         End With
37960       End If
37970       With frm
37980         .transdate.Visible = True
37990         .journaltype.Visible = True
38000         .totdesc.Visible = True
38010         .shareface_box.Visible = True
38020         .shareface_0.Visible = True
38030         .shareface_1.Visible = True
38040         .shareface_2.Visible = True
38050         .shareface_3.Visible = True
38060         .shareface_4.Visible = True
38070         .icash_box.Visible = True
38080         .pcash_box.Visible = True
38090         .cost_box.Visible = True
38100         For lngX = 0& To (lngItems - 1&)
38110           .Controls(arr_varItem(I_CNAM, lngX)).Visible = arr_varItem(I_VIS, lngX)
38120         Next
38130         .ledger_HIDDEN.Visible = True
38140         If blnNoTrans = False Then
38150           .transdate_lbl.ForeColor = CLR_DKGRY2
38160           .transdate_lbl_line.BorderColor = CLR_DKGRY
38170           .transdate_lbl_dim_hi.Visible = False
38180           .transdate_lbl_line_dim_hi.Visible = False
38190           .journaltype_lbl.ForeColor = CLR_DKGRY2
38200           .journaltype_lbl_line.BorderColor = CLR_DKGRY
38210           .journaltype_lbl_dim_hi.Visible = False
38220           .journaltype_lbl_line_dim_hi.Visible = False
38230           .totdesc_lbl.ForeColor = CLR_DKGRY2
38240           .totdesc_lbl_line.BorderColor = CLR_DKGRY
38250           .totdesc_lbl_dim_hi.Visible = False
38260           .totdesc_lbl_line_dim_hi.Visible = False
38270           .shareface_lbl.ForeColor = CLR_DKGRY2
38280           .shareface_lbl_line.BorderColor = CLR_DKGRY
38290           .shareface_lbl_dim_hi.Visible = False
38300           .shareface_lbl_line_dim_hi.Visible = False
38310           .icash_str_lbl.ForeColor = CLR_DKGRY2
38320           .icash_str_lbl_line.BorderColor = CLR_DKGRY
38330           .icash_str_lbl_dim_hi.Visible = False
38340           .icash_str_lbl_line_dim_hi.Visible = False
38350           .pcash_str_lbl.ForeColor = CLR_DKGRY2
38360           .pcash_str_lbl_line.BorderColor = CLR_DKGRY
38370           .pcash_str_lbl_dim_hi.Visible = False
38380           .pcash_str_lbl_line_dim_hi.Visible = False
38390           .cost_str_lbl.ForeColor = CLR_DKGRY2
38400           .cost_str_lbl_line.BorderColor = CLR_DKGRY
38410           .cost_str_lbl_dim_hi.Visible = False
38420           .cost_str_lbl_line_dim_hi.Visible = False
38430           .ledger_HIDDEN_lbl.ForeColor = CLR_DKGRY2
38440           .ledger_HIDDEN_lbl_line.BorderColor = CLR_DKGRY
38450           .ledger_HIDDEN_lbl_dim_hi.Visible = False
38460           .ledger_HIDDEN_lbl_line_dim_hi.Visible = False
38470         End If
38480       End With
38490     Case False
38500       If .chkShowJournalNo = True Then
38510         With frm
38520           .journalno.Visible = False
38530           .journalno_lbl.ForeColor = WIN_CLR_DISF
38540           .journalno_lbl2.ForeColor = WIN_CLR_DISF
38550           .journalno_lbl_line.BorderColor = WIN_CLR_DISR  'MAY BE WRONG!
38560           .journalno_lbl_dim_hi.Visible = True
38570           .journalno_lbl2_dim_hi.Visible = True
38580           .journalno_lbl_line_dim_hi.Visible = True
38590         End With
38600       End If
38610       If .chkIncludeCurrency = True Then
38620         With frm
38630           .curr_id.Visible = False
38640           .curr_id_box.Visible = False
38650           .curr_id_forex.Visible = False
38660           .curr_id_lbl.ForeColor = WIN_CLR_DISF
38670           .curr_id_lbl_line.BorderColor = WIN_CLR_DISR
38680           .curr_id_lbl_dim_hi.Visible = True
38690           .curr_id_lbl_line_dim_hi.Visible = True
38700         End With
38710       End If
38720       With frm
38730         .transdate.Visible = False
38740         .transdate_lbl.ForeColor = WIN_CLR_DISF
38750         .transdate_lbl_line.BorderColor = WIN_CLR_DISR
38760         .transdate_lbl_dim_hi.Visible = True
38770         .transdate_lbl_line_dim_hi.Visible = True
38780         .journaltype.Visible = False
38790         .journaltype_lbl.ForeColor = WIN_CLR_DISF
38800         .journaltype_lbl_line.BorderColor = WIN_CLR_DISR
38810         .journaltype_lbl_dim_hi.Visible = True
38820         .journaltype_lbl_line_dim_hi.Visible = True
38830         .totdesc.Visible = False
38840         .totdesc_lbl.ForeColor = WIN_CLR_DISF
38850         .totdesc_lbl_line.BorderColor = WIN_CLR_DISR
38860         .totdesc_lbl_dim_hi.Visible = True
38870         .totdesc_lbl_line_dim_hi.Visible = True
38880         .shareface_0.Visible = False
38890         .shareface_1.Visible = False
38900         .shareface_2.Visible = False
38910         .shareface_3.Visible = False
38920         .shareface_4.Visible = False
38930         .shareface_box.Visible = False
38940         .shareface_lbl.ForeColor = WIN_CLR_DISF
38950         .shareface_lbl_line.BorderColor = WIN_CLR_DISR
38960         .shareface_lbl_dim_hi.Visible = True
38970         .shareface_lbl_line_dim_hi.Visible = True
38980         .icash_box.Visible = False
38990         .icash_str_lbl.ForeColor = WIN_CLR_DISF
39000         .icash_str_lbl_line.BorderColor = WIN_CLR_DISR
39010         .icash_str_lbl_dim_hi.Visible = True
39020         .icash_str_lbl_line_dim_hi.Visible = True
39030         .pcash_box.Visible = False
39040         .pcash_str_lbl.ForeColor = WIN_CLR_DISF
39050         .pcash_str_lbl_line.BorderColor = WIN_CLR_DISR
39060         .pcash_str_lbl_dim_hi.Visible = True
39070         .pcash_str_lbl_line_dim_hi.Visible = True
39080         .cost_box.Visible = False
39090         .cost_str_lbl.ForeColor = WIN_CLR_DISF
39100         .cost_str_lbl_line.BorderColor = WIN_CLR_DISR
39110         .cost_str_lbl_dim_hi.Visible = True
39120         .cost_str_lbl_line_dim_hi.Visible = True
39130         For lngX = 0& To (lngItems - 1&)
39140           .Controls(arr_varItem(I_CNAM, lngX)).Visible = False
39150         Next
39160         .ledger_HIDDEN.Visible = False
39170         .ledger_HIDDEN_lbl.ForeColor = WIN_CLR_DISF
39180         .ledger_HIDDEN_lbl_line.BorderColor = WIN_CLR_DISR
39190         .ledger_HIDDEN_lbl_dim_hi.Visible = True
39200         .ledger_HIDDEN_lbl_line_dim_hi.Visible = True
39210       End With
39220     End Select
39230     blnNoTrans = False

39240   End With

EXITP:
39250   Set ctl = Nothing
39260   Exit Sub

ERRH:
39270   Select Case ERR.Number
        Case Else
39280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
39290   End Select
39300   Resume EXITP

End Sub
