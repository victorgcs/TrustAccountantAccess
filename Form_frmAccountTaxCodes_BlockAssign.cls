VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountTaxCodes_BlockAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountTaxCodes_BlockAssign"

'VGC 10/02/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   ShowJournalNo()
' ##   IncludeCurrency()
' #######################################

' ** Array: arr_varCtl().
Private lngCtls As Long, arr_varCtl() As Variant
Private Const C_ELEMS As Integer = 9  ' ** Array's first-element UBound().
Private Const C_CNAM   As Integer = 0
Private Const C_LFT    As Integer = 1
Private Const C_TOP    As Integer = 2
Private Const C_L1_NAM As Integer = 3
Private Const C_L1_LFT As Integer = 4
Private Const C_L1_Top As Integer = 5
Private Const C_L2_NAM As Integer = 6
Private Const C_L2_LFT As Integer = 7
Private Const C_L2_TOP As Integer = 8
Private Const C_ISFRM  As Integer = 9

Private blnSelectAll_Focus As Boolean, blnSelectAll_MouseDown As Boolean
Private blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean

Private lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngJournalNo_Width As Long, lngCurrID_Width As Long, lngTaxCodes_Width As Long

Private strAccountNo As String
Private blnRecalcTots As Boolean, blnArchive As Boolean, blnFromOK As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim frm As Access.Form
        Dim blnRetVal As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then
150           intPos01 = InStr(strTmp01, "~")
160           If intPos01 > 0 Then

170             blnIsOpen = True

180             blnRecalcTots = False: blnArchive = False: blnFromOK = False

190             strCallingForm = Left(strTmp01, (intPos01 - 1))
200             strTmp01 = Mid(strTmp01, (intPos01 + 1))
210             intPos01 = InStr(strTmp01, "~")
220             strAccountNo = Left(strTmp01, (intPos01 - 1))
230             blnArchive = CBool(Mid(strTmp01, (intPos01 + 1)))

240             DoCmd.Hourglass True
250             DoEvents

260             .accountno = gstrAccountNo
270             .shortname = gstrAccountName

280             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
290             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

300             blnSelectAll_Focus = False: blnSelectAll_MouseDown = False
310             blnSelectNone_Focus = False: blnSelectNone_MouseDown = False

320             .cmdClear.Enabled = False
330             .cmdAssign.Enabled = False
340             .cmdOK.Enabled = False

350             .form_width_line.Left = 0&  ' ** Just get it out of the way.

360             Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form

370             lngForm_Width = .Width
380             lngTaxCodes_Width = .cmbTaxCodes.Width

390             CtlArrayLoad  ' ** Procedure: Below.

400             gblnHasForEx = HasForEx_All  ' ** Module Function: modCurrencyFuncs.
410             gblnHasForExThis = HasForEx_Acct(strAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.

420             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
430             DoEvents

440             Select Case gblnHasForEx
                Case True
450               Select Case gblnHasForExThis
                  Case True
460                 .chkIncludeCurrency.Enabled = True
470               Case False
480                 .chkIncludeCurrency.Enabled = False
490                 If .chkIncludeCurrency = True Then .chkIncludeCurrency = False
500               End Select
510             Case False
520               .chkIncludeCurrency.Enabled = False
530               If .chkIncludeCurrency = True Then .chkIncludeCurrency = False
540             End Select

                ' ** Variables are fed empty, then populated ByRef.
550             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

560             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

570             chkShowJournalNo_AfterUpdate  ' ** Procedure: Below.
580             chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
590             chkShowUnassignedOnly_AfterUpdate  ' ** Procedure: Below.
600             chkShowCheckedOnly_AfterUpdate  ' ** Procedure: Below.
610             cmbJournalTypes_AfterUpdate  ' ** Procedure: Below.

620             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
630             If lngMonitorCnt > 1& Then
640               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
650               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
660               If lngMonitorNum > 1& Then
670                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
680               End If
690             End If

700             .cmbTaxCodes.SetFocus

710             DoCmd.Hourglass False

720             blnIsOpen = False

730           Else
740             Cancel = -1
750           End If
760         Else
770           Cancel = -1
780         End If
790       Else
800         Cancel = -1
810       End If
820     End With

830     If Cancel = -1 Then
840       Beep
850       DoCmd.Hourglass False
860       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
870     End If

EXITP:
880     Set frm = Nothing
890     Exit Sub

ERRH:
900     DoCmd.Hourglass False
910     Select Case ERR.Number
        Case Else
920       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
930     End Select
940     Resume EXITP

End Sub

Private Sub Form_Load()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1010    With Me

1020      DoCmd.Hourglass True
1030      DoEvents

1040      .frmAccountTaxCodes_BlockAssign_Sub.Form.Requery

1050      .cmbTaxCodes.SetFocus

1060      DoCmd.Hourglass False

1070    End With

EXITP:
1080    Exit Sub

ERRH:
1090    DoCmd.Hourglass False
1100    Select Case ERR.Number
        Case Else
1110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1120    End Select
1130    Resume EXITP

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

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:           F5 {RecalcTots}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Rev Code:         D {cmbTaxCodes}
        ' **   Filter by JType:  J {cmbJournalTypes}
        ' **   Check This:       X {chkTaxChk on frmAccountTaxCodes_BlockAssign_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAccountTaxCodes_BlockAssign_Sub}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyF5
1270        With Me
1280          intRetVal = 0
1290          .frmAccountTaxCodes_BlockAssign_Sub.Form.RecalcTots  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
1300        End With
1310      Case vbKeyEscape
1320        intRetVal = 0
1330        cmdCancel_Click  ' ** Procedure: Below.
1340      End Select
1350    End If

        ' ** Alt keys.
1360    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1370      Select Case intRetVal
          Case vbKeyD
1380        With Me
1390          intRetVal = 0
1400          .cmbTaxCodes.SetFocus
1410        End With
1420      Case vbKeyJ
1430        With Me
1440          intRetVal = 0
1450          .cmbJournalTypes.SetFocus
1460        End With
1470      Case vbKeyX
1480        With Me
1490          intRetVal = 0
1500          lngRecsCur = .frmAccountTaxCodes_BlockAssign_Sub.Form.RecCnt  ' ** Form Function: frmAccountTaxCodes_BlockAssign_Sub.
1510          If lngRecsCur > 0& Then
1520            .frmAccountTaxCodes_BlockAssign_Sub.SetFocus
1530            .frmAccountTaxCodes_BlockAssign_Sub.Form.chkTaxChk.SetFocus
1540            .frmAccountTaxCodes_BlockAssign_Sub.Form.chkTaxChk = (Not .frmAccountTaxCodes_BlockAssign_Sub.Form.chkTaxChk)
1550            .frmAccountTaxCodes_BlockAssign_Sub.Form.chkTaxChk_AfterUpdate  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
1560          Else
1570            Beep
1580          End If
1590        End With
1600      End Select
1610    End If

        ' ** Ctrl keys.
1620    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1630      Select Case intRetVal
          Case vbKeyS
1640        With Me
1650          intRetVal = 0
1660          .frmAccountTaxCodes_BlockAssign_Sub.SetFocus
1670          .frmAccountTaxCodes_BlockAssign_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
1680        End With
1690      End Select
1700    End If

        ' ** Ctrl-Shift keys.
1710    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1720      Select Case intRetVal
          Case vbKeyF
1730        With Me
1740          intRetVal = 0
1750          .FocusHolder.SetFocus
1760        End With
1770      End Select
1780    End If

EXITP:
1790    KeyCode = intRetVal
1800    Exit Sub

ERRH:
1810    intRetVal = 0
1820    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1830    Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub Form_Timer()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1910    With Me
1920      .TimerInterval = 0&
1930      If gblnSetFocus = True Then
1940        gblnSetFocus = False
1950        strFormName = vbNullString
1960  On Error Resume Next
1970        Do Until strFormName = THIS_NAME
1980          DoCmd.SelectObject acForm, THIS_NAME, False
1990          strFormName = Screen.ActiveForm.Name
2000        Loop
2010  On Error GoTo ERRH
2020      End If
2030      If blnRecalcTots = True Then
2040        blnRecalcTots = False
2050        .frmAccountTaxCodes_BlockAssign_Sub.Form.RecalcTots  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
2060      End If
2070    End With

EXITP:
2080    Exit Sub

ERRH:
2090    Select Case ERR.Number
        Case Else
2100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2110    End Select
2120    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2210    If DataErr <> 0 Then
2220      Select Case DataErr
          Case 2115  ' ** The macro or function set to the BeforeUpdate or ValidationRule property for this field is preventing Microsoft Access from saving the data in the field.
            ' ** RecalcTots getting piled up, Form_Timer() calling here and subform.
2230        gblnSignal = False
2240        Response = acDataErrContinue
2250      Case Else
2260        zErrorHandler THIS_NAME, THIS_PROC, DataErr, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270      End Select
2280    End If

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
2290    Exit Sub

ERRH:
2300    Select Case ERR.Number
        Case Else
2310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2320    End Select
2330    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim frm As Access.Form
        Dim strJType As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

2410    With Me

2420      blnContinue = True
2430      DoCmd.Hourglass True
2440      DoEvents

2450      If blnFromOK = False Then
2460        strJType = .cmbJournalTypes
2470        If .cmbJournalTypes <> "All" Then
2480          .cmbJournalTypes = "All"
2490          cmbJournalTypes_AfterUpdate  ' ** Procedure: Below.
2500          DoEvents
2510        End If
2520        Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form
2530        frm.RecalcTots  ' ** Form procedure: frmAccountTaxCodes_BlockAssign_Sub.
2540        DoEvents
2550        lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountTaxCodes_BlockAssign_Sub.
2560        If lngRecsCur > 0& Then
2570          If .tcbas_changed_sum > 0 Then
2580            DoCmd.Hourglass False
2590            Beep
2600            msgResponse = MsgBox("You have made changes, but not confirmed them by clicking OK." & vbCrLf & _
                  "Are you sure you want to abandon your changes?", vbQuestion + vbYesNo + vbDefaultButton2, _
                  "Changes Not Saved")
2610            If msgResponse <> vbYes Then
2620              Cancel = -1
2630              blnContinue = False
2640            End If
2650          End If
2660        End If
2670      End If  ' ** blnFromOK

2680      If blnContinue = True Then

2690        If .cmbJournalTypes <> strJType Then
2700          .cmbJournalTypes = strJType
2710        End If

2720        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2730        gstrReturningForm = THIS_NAME
2740        gblnSetFocus = True
2750        If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Fuction: modFileUtilities.
2760          DoCmd.SelectObject acForm, strCallingForm, False
2770          Forms(strCallingForm).TimerInterval = 200&
2780        Else
2790          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Account"
2800        End If

2810      End If

2820    End With

EXITP:
2830    Set frm = Nothing
2840    Exit Sub

ERRH:
2850    DoCmd.Hourglass False
2860    Select Case ERR.Number
        Case Else
2870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2880    End Select
2890    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2910    With Me
2920      If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
2930        Select Case blnSelectAll_Focus
            Case True
2940          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
2950          .cmdSelectAll_raised_img.Visible = False
2960        Case False
2970          .cmdSelectAll_raised_img.Visible = True
2980          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
2990        End Select
3000        .cmdSelectAll_raised_focus_img.Visible = False
3010        .cmdSelectAll_raised_focus_dots_img.Visible = False
3020        .cmdSelectAll_sunken_focus_dots_img.Visible = False
3030        .cmdSelectAll_raised_img_dis.Visible = False
3040      End If
3050      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
3060        Select Case blnSelectNone_Focus
            Case True
3070          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
3080          .cmdSelectNone_raised_img.Visible = False
3090        Case False
3100          .cmdSelectNone_raised_img.Visible = True
3110          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
3120        End Select
3130        .cmdSelectNone_raised_focus_img.Visible = False
3140        .cmdSelectNone_raised_focus_dots_img.Visible = False
3150        .cmdSelectNone_sunken_focus_dots_img.Visible = False
3160        .cmdSelectNone_raised_img_dis.Visible = False
3170      End If
3180    End With

EXITP:
3190    Exit Sub

ERRH:
3200    Select Case ERR.Number
        Case Else
3210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3220    End Select
3230    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub cmdOK_Click()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, frm As Access.Form
        Dim lngRecsChanged As Long

3410    With Me

3420      DoCmd.Hourglass True
3430      DoEvents

3440      If .cmbJournalTypes <> "All" Then
3450        .cmbJournalTypes = "All"
3460        cmbJournalTypes_AfterUpdate  ' ** Procedure: Below.
3470        DoEvents
3480      End If

3490      DoCmd.Hourglass True
3500      DoEvents

3510      Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form

3520      lngRecsCur = frm.RecCnt  ' ** Form Function: frmAccountTaxCodes_BlockAssign_Sub.
3530      lngRecsChanged = .tcbas_changed_sum

3540      If lngRecsChanged > 0& Then

3550        Set dbs = CurrentDb

            ' ** Update qryAccountProfile_TaxCodes_Block_08_01 (Ledger, linked to
            ' ** tblTaxCode_BlockAssign_Staging, just tcbas_changed = True, with taxcode_new).
3560        Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_Block_08_02")
3570        qdf.Execute
3580        Set qdf = Nothing
3590        DoEvents

3600        If blnArchive = True Then
              ' ** Update qryAccountProfile_TaxCodes_Block_08_03 (LedgerArchive, linked to
              ' ** tblTaxCode_BlockAssign_Staging, just tcbas_changed = True, with taxcode_new).
3610          Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_Block_08_04")
3620          qdf.Execute
3630          Set qdf = Nothing
3640        End If
3650        DoEvents

            ' ** Update qryAccountProfile_TaxCodes_Block_08_05 (tmpEdit06, linked to
            ' ** tblTaxCode_BlockAssign_Staging, with taxcode_new, taxcode_DESC_new,
            ' ** taxcode_DESC_display_new).
3660        Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_Block_08_06")
3670        qdf.Execute
3680        Set qdf = Nothing
3690        DoEvents

3700        If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
3710          Forms(strCallingForm).frmAccountTaxCodes_Sub.Form.Requery
3720        End If
3730        DoEvents

3740        blnFromOK = True

3750        DoCmd.Close acForm, THIS_NAME

3760      Else
            ' ** Should absolutely, positively not happen!
3770        DoCmd.Hourglass False
3780        MsgBox "What?"
3790      End If
3800    End With

EXITP:
3810    Set frm = Nothing
3820    Set qdf = Nothing
3830    Set dbs = Nothing
3840    Exit Sub

ERRH:
3850    Select Case ERR.Number
        Case Else
3860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3870    End Select
3880    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3920    intShiftDown = (Shift And acShiftMask) > 0
3930    intAltDown = (Shift And acAltMask) > 0
3940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3960      Select Case intRetVal
          Case vbKeyTab
3970        With Me
3980          intRetVal = 0
3990          .cmdCancel.SetFocus
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab
4050        With Me
4060          intRetVal = 0
4070          If .cmdAssign.Enabled = True Then
4080            .cmdAssign.SetFocus
4090          ElseIf .cmdClear.Enabled = True Then
4100            .cmdClear.SetFocus
4110          Else
4120            .chkShowCheckedOnly.SetFocus
4130          End If
4140        End With
4150      End Select
4160    End If

EXITP:
4170    KeyCode = intRetVal
4180    Exit Sub

ERRH:
4190    intRetVal = 0
4200    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4210    Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

4310    With Me

4320      DoCmd.Hourglass True
4330      DoEvents

4340      DoCmd.Close acForm, THIS_NAME

4350    End With

EXITP:
4360    Exit Sub

ERRH:
4370    DoCmd.Hourglass False
4380    Select Case ERR.Number
        Case 2455  ' ** You entered an expression that has an invalid reference to the property |.
          ' ** Ignore.
4390    Case Else
4400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4410    End Select
4420    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
4590          .cmbTaxCodes.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab
4650        With Me
4660          intRetVal = 0
4670          If .cmdOK.Enabled = True Then
4680            .cmdOK.SetFocus
4690          ElseIf .cmdAssign.Enabled = True Then
4700            .cmdAssign.SetFocus
4710          ElseIf .cmdClear.Enabled = True Then
4720            .cmdClear.SetFocus
4730          Else
4740            .chkShowCheckedOnly.SetFocus
4750          End If
4760        End With
4770      End Select
4780    End If

EXITP:
4790    KeyCode = intRetVal
4800    Exit Sub

ERRH:
4810    intRetVal = 0
4820    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4830    Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub cmbTaxCodes_AfterUpdate()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_AfterUpdate"

        Dim varTmp00 As Variant

4910    With Me
4920      If IsNull(.cmbTaxCodes) = False Then
4930        varTmp00 = .frmAccountTaxCodes_BlockAssign_Sub.Form.chkTaxChk_sum
4940        If IsNull(varTmp00) = False Then
4950          If Abs(varTmp00) > 0 Then
4960            .cmdAssign.Enabled = True
4970          End If
4980        End If
4990      End If
5000    End With

EXITP:
5010    Exit Sub

ERRH:
5020    Select Case ERR.Number
        Case Else
5030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5040    End Select
5050    Resume EXITP

End Sub

Private Sub cmbTaxCodes_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_KeyDown"

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
5190          lngRecsCur = .frmAccountTaxCodes_BlockAssign_Sub.Form.RecCnt  ' ** Form Function: frmAccountTaxCodes_BlockAssign_Sub.
5200          If lngRecsCur > 0& Then
5210            .frmAccountTaxCodes_BlockAssign_Sub.SetFocus
5220            .frmAccountTaxCodes_BlockAssign_Sub.Form.FocusHolder.SetFocus
5230          Else
5240            .cmbJournalTypes.SetFocus
5250          End If
5260        End With
5270      End Select
5280    End If

        ' ** Shift keys.
5290    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5300      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5310        With Me
5320          intRetVal = 0
5330          .cmdCancel.SetFocus
5340        End With
5350      End Select
5360    End If

EXITP:
5370    KeyCode = intRetVal
5380    Exit Sub

ERRH:
5390    intRetVal = 0
5400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5410    Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

5510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Private Sub chkTaxChk_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxChk_box_MouseMove"

5610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub chkTaxChk_sum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkTaxChk_sum_MouseMove"

5710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5720    Exit Sub

ERRH:
5730    Select Case ERR.Number
        Case Else
5740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5750    End Select
5760    Resume EXITP

End Sub

Private Sub cmdSelectAll_GotFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_GotFocus"

5810    With Me
5820      blnSelectAll_Focus = True
5830      .cmdSelectAll_raised_semifocus_dots_img.Visible = True
5840      .cmdSelectAll_raised_img.Visible = False
5850      .cmdSelectAll_raised_focus_img.Visible = False
5860      .cmdSelectAll_raised_focus_dots_img.Visible = False
5870      .cmdSelectAll_sunken_focus_dots_img.Visible = False
5880      .cmdSelectAll_raised_img_dis.Visible = False
5890    End With

EXITP:
5900    Exit Sub

ERRH:
5910    Select Case ERR.Number
        Case Else
5920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5930    End Select
5940    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseDown"

6010    With Me
6020      blnSelectAll_MouseDown = True
6030      .cmdSelectAll_sunken_focus_dots_img.Visible = True
6040      .cmdSelectAll_raised_img.Visible = False
6050      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
6060      .cmdSelectAll_raised_focus_img.Visible = False
6070      .cmdSelectAll_raised_focus_dots_img.Visible = False
6080      .cmdSelectAll_raised_img_dis.Visible = False
6090    End With

EXITP:
6100    Exit Sub

ERRH:
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Private Sub cmdSelectAll_Click()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_Click"

        Dim rst As DAO.Recordset, frm As Access.Form
        Dim lngRecs As Long
        Dim lngX As Long

6210    With Me
6220      DoCmd.Hourglass True
6230      DoEvents
6240      Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form
6250      Set rst = frm.RecordsetClone
6260      With rst
6270        If .BOF = True And .EOF = True Then
6280          Beep
6290        Else
6300          .MoveLast
6310          lngRecs = .RecordCount
6320          .MoveFirst
6330          For lngX = 1& To lngRecs
6340            .Edit
6350            ![chkTaxChk] = True
6360            ![tcbas_datemodified] = Now()
6370            .Update
6380            If lngX < lngRecs Then .MoveNext
6390          Next
6400        End If
6410        .Close
6420      End With
6430      Set rst = Nothing
6440      .cmdClear.Enabled = True
6450      If IsNull(.cmbTaxCodes) = False Then
6460        .cmdAssign.Enabled = True
6470      Else
6480        .cmdAssign.Enabled = False
6490      End If
6500      frm.Refresh
6510      DoCmd.Hourglass False
6520      Beep
6530      blnRecalcTots = True
6540      .TimerInterval = 500&
6550    End With

EXITP:
6560    Set frm = Nothing
6570    Set rst = Nothing
6580    Exit Sub

ERRH:
6590    DoCmd.Hourglass False
6600    Select Case ERR.Number
        Case Else
6610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6620    End Select
6630    Resume EXITP

End Sub

Private Sub cmdSelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_KeyDown"

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
          Case vbKeyTab
6770        With Me
6780          intRetVal = 0
6790          .cmdSelectNone.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab
6850        With Me
6860          intRetVal = 0
6870          lngRecsCur = .frmAccountTaxCodes_BlockAssign_Sub.Form.RecCnt  ' ** Form Function: frmAccountTaxCodes_BlockAssign_Sub.
6880          If lngRecsCur > 0& Then
6890            .frmAccountTaxCodes_BlockAssign_Sub.SetFocus
6900          Else
6910            .cmbTaxCodes.SetFocus
6920          End If
6930        End With
6940      End Select
6950    End If

EXITP:
6960    KeyCode = intRetVal
6970    Exit Sub

ERRH:
6980    intRetVal = 0
6990    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7000    Case Else
7010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7020    End Select
7030    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseMove"

7110    With Me
7120      If blnSelectAll_MouseDown = False Then
7130        Select Case blnSelectAll_Focus
            Case True
7140          .cmdSelectAll_raised_focus_dots_img.Visible = True
7150          .cmdSelectAll_raised_focus_img.Visible = False
7160        Case False
7170          .cmdSelectAll_raised_focus_img.Visible = True
7180          .cmdSelectAll_raised_focus_dots_img.Visible = False
7190        End Select
7200        .cmdSelectAll_raised_img.Visible = False
7210        .cmdSelectAll_raised_semifocus_dots_img.Visible = False
7220        .cmdSelectAll_sunken_focus_dots_img.Visible = False
7230        .cmdSelectAll_raised_img_dis.Visible = False
7240      End If
7250    End With

EXITP:
7260    Exit Sub

ERRH:
7270    Select Case ERR.Number
        Case Else
7280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7290    End Select
7300    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseUp"

7410    With Me
7420      .cmdSelectAll_raised_focus_dots_img.Visible = True
7430      .cmdSelectAll_raised_img.Visible = False
7440      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
7450      .cmdSelectAll_raised_focus_img.Visible = False
7460      .cmdSelectAll_sunken_focus_dots_img.Visible = False
7470      .cmdSelectAll_raised_img_dis.Visible = False
7480      blnSelectAll_MouseDown = False
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

Private Sub cmdSelectAll_LostFocus()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_LostFocus"

7610    With Me
7620      .cmdSelectAll_raised_img.Visible = True
7630      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
7640      .cmdSelectAll_raised_focus_img.Visible = False
7650      .cmdSelectAll_raised_focus_dots_img.Visible = False
7660      .cmdSelectAll_sunken_focus_dots_img.Visible = False
7670      .cmdSelectAll_raised_img_dis.Visible = False
7680      blnSelectAll_Focus = False
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

Private Sub cmdSelectNone_GotFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_GotFocus"

7810    With Me
7820      blnSelectNone_Focus = True
7830      .cmdSelectNone_raised_semifocus_dots_img.Visible = True
7840      .cmdSelectNone_raised_img.Visible = False
7850      .cmdSelectNone_raised_focus_img.Visible = False
7860      .cmdSelectNone_raised_focus_dots_img.Visible = False
7870      .cmdSelectNone_sunken_focus_dots_img.Visible = False
7880      .cmdSelectNone_raised_img_dis.Visible = False
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

Private Sub cmdSelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseDown"

8010    With Me
8020      blnSelectNone_MouseDown = True
8030      .cmdSelectNone_sunken_focus_dots_img.Visible = True
8040      .cmdSelectNone_raised_img.Visible = False
8050      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
8060      .cmdSelectNone_raised_focus_img.Visible = False
8070      .cmdSelectNone_raised_focus_dots_img.Visible = False
8080      .cmdSelectNone_raised_img_dis.Visible = False
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

Private Sub cmdSelectNone_Click()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_Click"

        Dim rst As DAO.Recordset, frm As Access.Form
        Dim lngRecs As Long
        Dim lngX As Long

8210    With Me
8220      DoCmd.Hourglass True
8230      DoEvents
8240      Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form
8250      Set rst = frm.RecordsetClone
8260      With rst
8270        If .BOF = True And .EOF = True Then
8280          Beep
8290        Else
8300          .MoveLast
8310          lngRecs = .RecordCount
8320          .MoveFirst
8330          For lngX = 1& To lngRecs
8340            .Edit
8350            ![chkTaxChk] = False
8360            ![tcbas_datemodified] = Now()
8370            .Update
8380            If lngX < lngRecs Then .MoveNext
8390          Next
8400        End If
8410        .Close
8420      End With
8430      Set rst = Nothing
8440      frm.Refresh
8450      .cmdClear.Enabled = False
8460      .cmdAssign.Enabled = False
8470      .cmdOK.Enabled = False
8480      DoCmd.Hourglass False
8490      Beep
8500      blnRecalcTots = True
8510      .TimerInterval = 500&
8520    End With

EXITP:
8530    Set frm = Nothing
8540    Set rst = Nothing
8550    Exit Sub

ERRH:
8560    DoCmd.Hourglass False
8570    Select Case ERR.Number
        Case Else
8580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8590    End Select
8600    Resume EXITP

End Sub

Private Sub cmdSelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_KeyDown"

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
          Case vbKeyTab
8770        With Me
8780          intRetVal = 0
8790          .cmbJournalTypes.SetFocus
8800        End With
8810      End Select
8820    End If

        ' ** Shift keys.
8830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8840      Select Case intRetVal
          Case vbKeyTab
8850        With Me
8860          intRetVal = 0
8870          .cmdSelectAll.SetFocus
8880        End With
8890      End Select
8900    End If

EXITP:
8910    KeyCode = intRetVal
8920    Exit Sub

ERRH:
8930    intRetVal = 0
8940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8950    Case Else
8960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8970    End Select
8980    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseMove"

9010    With Me
9020      If blnSelectNone_MouseDown = False Then
9030        Select Case blnSelectNone_Focus
            Case True
9040          .cmdSelectNone_raised_focus_dots_img.Visible = True
9050          .cmdSelectNone_raised_focus_img.Visible = False
9060        Case False
9070          .cmdSelectNone_raised_focus_img.Visible = True
9080          .cmdSelectNone_raised_focus_dots_img.Visible = False
9090        End Select
9100        .cmdSelectNone_raised_img.Visible = False
9110        .cmdSelectNone_raised_semifocus_dots_img.Visible = False
9120        .cmdSelectNone_sunken_focus_dots_img.Visible = False
9130        .cmdSelectNone_raised_img_dis.Visible = False
9140      End If
9150    End With

EXITP:
9160    Exit Sub

ERRH:
9170    Select Case ERR.Number
        Case Else
9180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9190    End Select
9200    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseUp"

9310    With Me
9320      .cmdSelectNone_raised_focus_dots_img.Visible = True
9330      .cmdSelectNone_raised_img.Visible = False
9340      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
9350      .cmdSelectNone_raised_focus_img.Visible = False
9360      .cmdSelectNone_sunken_focus_dots_img.Visible = False
9370      .cmdSelectNone_raised_img_dis.Visible = False
9380      blnSelectNone_MouseDown = False
9390    End With

EXITP:
9400    Exit Sub

ERRH:
9410    Select Case ERR.Number
        Case Else
9420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9430    End Select
9440    Resume EXITP

End Sub

Private Sub cmdSelectNone_LostFocus()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_LostFocus"

9510    With Me
9520      .cmdSelectNone_raised_img.Visible = True
9530      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
9540      .cmdSelectNone_raised_focus_img.Visible = False
9550      .cmdSelectNone_raised_focus_dots_img.Visible = False
9560      .cmdSelectNone_sunken_focus_dots_img.Visible = False
9570      .cmdSelectNone_raised_img_dis.Visible = False
9580      blnSelectNone_Focus = False
9590    End With

EXITP:
9600    Exit Sub

ERRH:
9610    Select Case ERR.Number
        Case Else
9620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9630    End Select
9640    Resume EXITP

End Sub

Private Sub cmbJournalTypes_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalTypes_box_MouseMove"

9710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9720    Exit Sub

ERRH:
9730    Select Case ERR.Number
        Case Else
9740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9750    End Select
9760    Resume EXITP

End Sub

Private Sub cmbJournalTypes_DblClick(Cancel As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalTypes_DblClick"

9810    With Me
9820      .cmbJournalTypes = "All"
9830      cmbJournalTypes_AfterUpdate  ' ** Procedure: Below.
9840    End With

EXITP:
9850    Exit Sub

ERRH:
9860    Select Case ERR.Number
        Case Else
9870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9880    End Select
9890    Resume EXITP

End Sub

Private Sub cmbJournalTypes_AfterUpdate()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalTypes_AfterUpdate"

        Dim strFilter As String, blnFilterOn As Boolean
        Dim strTmp01 As String

9910    With Me

9920      strFilter = .frmAccountTaxCodes_BlockAssign_Sub.Form.Filter
9930      blnFilterOn = .frmAccountTaxCodes_BlockAssign_Sub.Form.FilterOn

9940      If IsNull(.cmbJournalTypes) = False Then
9950        strTmp01 = .cmbJournalTypes
9960        If strTmp01 = "All" Then
9970          Select Case .chkShowUnassignedOnly
              Case True
9980            Select Case .chkShowCheckedOnly
                Case True
9990              strFilter = "[taxcode] <= 2 And [chkTaxChk] = True"
10000           Case False
10010             strFilter = "[taxcode] <= 2"
10020           End Select
10030         Case False
10040           Select Case .chkShowCheckedOnly
                Case True
10050             strFilter = "[chkTaxChk] = True"
10060           Case False
10070             strFilter = vbNullString
10080           End Select
10090         End Select
10100         .cmbJournalTypes_box.BackColor = MY_CLR_LTBGE
10110         .cmbJournalTypes_lbl.ForeColor = CLR_VDKGRY
10120         .cmbJournalTypes_lbl2.ForeColor = CLR_VDKGRY
10130       Else
10140         strTmp01 = "[journaltypex] = '" & strTmp01 & "'"
10150         Select Case .chkShowUnassignedOnly
              Case True
10160           Select Case .chkShowCheckedOnly
                Case True
10170             strFilter = strTmp01 & " And [taxcode] <= 2 And [chkTaxChk] = True"
10180           Case False
10190             strFilter = strTmp01 & " And [taxcode] <= 2"
10200           End Select
10210         Case False
10220           Select Case .chkShowCheckedOnly
                Case True
10230             strFilter = "[chkTaxChk] = True"
10240           Case False
10250             strFilter = strTmp01
10260           End Select
10270         End Select
10280         .cmbJournalTypes_box.BackColor = CLR_VLTYEL
10290         .cmbJournalTypes_lbl.ForeColor = CLR_BRN
10300         .cmbJournalTypes_lbl2.ForeColor = CLR_BRN
10310       End If
10320     Else
10330       strTmp01 = "All"
10340       .cmbJournalTypes = strTmp01
10350       Select Case .chkShowUnassignedOnly
            Case True
10360         Select Case .chkShowCheckedOnly
              Case True
10370           strFilter = "[taxcode] <= 2 And [chkTaxChk] = True"
10380         Case False
10390           strFilter = "[taxcode] <= 2"
10400         End Select
10410       Case False
10420         Select Case .chkShowCheckedOnly
              Case True
10430           strFilter = "[chkTaxChk] = True"
10440         Case False
10450           strFilter = vbNullString
10460         End Select
10470       End Select
10480       .cmbJournalTypes_box.BackColor = MY_CLR_LTBGE
10490       .cmbJournalTypes_lbl.ForeColor = CLR_VDKGRY
10500       .cmbJournalTypes_lbl2.ForeColor = CLR_VDKGRY
10510     End If
10520     .frmAccountTaxCodes_BlockAssign_Sub.Form.Filter = strFilter
10530     If strFilter <> vbNullString Then
10540       .frmAccountTaxCodes_BlockAssign_Sub.Form.FilterOn = True
10550     Else
10560       .frmAccountTaxCodes_BlockAssign_Sub.Form.FilterOn = False
10570     End If

10580     lngRecsCur = .frmAccountTaxCodes_BlockAssign_Sub.Form.RecCnt  ' ** Form Function: frmAccountTaxCodes_BlockAssign_Sub.
10590     If lngRecsCur > 0& Then
10600       .frmAccountTaxCodes_BlockAssign_Sub.SetFocus
10610     End If

10620     blnRecalcTots = True
10630     .TimerInterval = 500&

10640   End With

EXITP:
10650   Exit Sub

ERRH:
10660   Select Case ERR.Number
        Case Else
10670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10680   End Select
10690   Resume EXITP

End Sub

Private Sub cmbJournalTypes_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalTypes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         intRetVal = 0
10790         .chkShowJournalNo.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Shift keys.
10830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10850       With Me
10860         intRetVal = 0
10870         .cmdSelectNone.SetFocus
10880       End With
10890     End Select
10900   End If

EXITP:
10910   KeyCode = intRetVal
10920   Exit Sub

ERRH:
10930   intRetVal = 0
10940   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
10950   Case Else
10960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10970   End Select
10980   Resume EXITP

End Sub

Private Sub cmbJournalTypes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbJournalTypes_MouseMove"

11010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11020   Exit Sub

ERRH:
11030   Select Case ERR.Number
        Case Else
11040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11050   End Select
11060   Resume EXITP

End Sub

Private Sub chkShowJournalNo_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowJournalNo_AfterUpdate"

11110   With Me
11120     Select Case .chkShowJournalNo
          Case True
11130       .chkShowJournalNo_lbl.FontBold = True
11140       .frmAccountTaxCodes_BlockAssign_Sub.Form.ShowJournalNo_Sub True  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
11150       ShowJournalNo True  ' ** Procedure: Below.
11160     Case False
11170       .chkShowJournalNo_lbl.FontBold = False
11180       .frmAccountTaxCodes_BlockAssign_Sub.Form.ShowJournalNo_Sub False  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
11190       ShowJournalNo False  ' ** Procedure: Below.
11200     End Select
11210     blnRecalcTots = True
11220     .TimerInterval = 500&
11230   End With

EXITP:
11240   Exit Sub

ERRH:
11250   Select Case ERR.Number
        Case Else
11260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11270   End Select
11280   Resume EXITP

End Sub

Private Sub chkShowJournalNo_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowJournalNo_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
11370       With Me
11380         intRetVal = 0
11390         If .chkIncludeCurrency.Enabled = True Then
11400           .chkIncludeCurrency.SetFocus
11410         Else
11420           .chkShowUnassignedOnly.SetFocus
11430         End If
11440       End With
11450     End Select
11460   End If

        ' ** Shift keys.
11470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11490       With Me
11500         intRetVal = 0
11510         .cmbJournalTypes.SetFocus
11520       End With
11530     End Select
11540   End If

EXITP:
11550   KeyCode = intRetVal
11560   Exit Sub

ERRH:
11570   intRetVal = 0
11580   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
11590   Case Else
11600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11610   End Select
11620   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

11710   With Me
11720     Select Case .chkIncludeCurrency
          Case True
11730       .chkIncludeCurrency_lbl.FontBold = True
11740       .frmAccountTaxCodes_BlockAssign_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
11750       IncludeCurrency True  ' ** Procedure: Below.
11760     Case False
11770       .chkIncludeCurrency_lbl.FontBold = False
11780       .frmAccountTaxCodes_BlockAssign_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
11790       IncludeCurrency False  ' ** Procedure: Below.
11800     End Select
11810     blnRecalcTots = True
11820     .TimerInterval = 500&
11830   End With

EXITP:
11840   Exit Sub

ERRH:
11850   Select Case ERR.Number
        Case Else
11860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11870   End Select
11880   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11920   intShiftDown = (Shift And acShiftMask) > 0
11930   intAltDown = (Shift And acAltMask) > 0
11940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11970       With Me
11980         intRetVal = 0
11990         .chkShowUnassignedOnly.SetFocus
12000       End With
12010     End Select
12020   End If

        ' ** Shift keys.
12030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12050       With Me
12060         intRetVal = 0
12070         .chkShowJournalNo.SetFocus
12080       End With
12090     End Select
12100   End If

EXITP:
12110   KeyCode = intRetVal
12120   Exit Sub

ERRH:
12130   intRetVal = 0
12140   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
12150   Case Else
12160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12170   End Select
12180   Resume EXITP

End Sub

Private Sub chkShowUnassignedOnly_AfterUpdate()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowUnassignedOnly_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String, blnFilterOn As Boolean
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String

12210   With Me
          ' ** Filters always seem to get huge!
12220     DoCmd.Hourglass True
12230     DoEvents
12240     Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form
12250     Select Case .chkShowUnassignedOnly
          Case True
12260       .chkShowUnassignedOnly_lbl.FontBold = True
12270       strTmp01 = "[taxcode] <= 2"
12280       strTmp02 = frm.Filter
12290       blnFilterOn = frm.FilterOn
12300       Select Case blnFilterOn
            Case True
12310         intPos01 = InStr(strTmp02, "[taxcode]")
12320         If intPos01 > 0 Then
12330           intPos02 = InStr(intPos01, strTmp02, " And ")
12340           If intPos02 > 0 Then
12350             strTmp03 = Trim(Mid(strTmp02, intPos02))  ' ** Starts with 'And '.
12360             If intPos01 > 1 Then
12370               strTmp02 = Trim(Left(strTmp02, (intPos01 - 1)))  ' ** Should end with ' And'.
12380               strFilter = strTmp02 & " " & strTmp01 & " " & strTmp03
12390             Else
12400               strTmp03 = Trim(Mid(strTmp03, 4))
12410               strFilter = strTmp03 & " And " & strTmp01
12420             End If
12430           Else
12440             If intPos01 > 1 Then
12450               strTmp02 = Trim(Left(strTmp02, (intPos01 - 1)))  ' ** Should end with ' And'.
12460               strFilter = strTmp02 & " " & strTmp01
12470             Else
12480               strFilter = strTmp01
12490             End If
12500           End If
12510         Else
12520           strFilter = strTmp02 & " And " & strTmp01
12530         End If
12540       Case False
              ' ** If the filter's off, throw out anything in it.
12550         strFilter = strTmp01
12560       End Select
12570       frm.Filter = strFilter
12580       If strFilter <> vbNullString Then
12590         frm.FilterOn = True
12600       Else
12610         frm.FilterOn = False
12620       End If
12630     Case False
12640       .chkShowUnassignedOnly_lbl.FontBold = False
12650       strTmp01 = vbNullString
12660       strTmp02 = frm.Filter
12670       blnFilterOn = frm.FilterOn
12680       Select Case blnFilterOn
            Case True
12690         intPos01 = InStr(strTmp02, "[taxcode]")
12700         If intPos01 > 0 Then
12710           intPos02 = InStr(intPos01, strTmp02, " And ")
12720           If intPos02 > 0 Then
12730             strTmp03 = Trim(Mid(strTmp02, intPos02))  ' ** Starts with 'And '.
12740             If intPos01 > 1 Then
12750               strTmp02 = Trim(Left(strTmp02, (intPos01 - 1)))  ' ** Should end with ' And'.
12760               strFilter = StringReplace((strTmp02 & " " & strTmp03), " And And ", " And ")  ' ** Module Function: modStringFuncs.
12770             Else
12780               strTmp03 = Trim(Mid(strTmp03, 4))  ' ** Strip of 'And '.
12790               strFilter = strTmp03
12800             End If
12810           Else
12820             If intPos01 > 1 Then
12830               strTmp02 = Trim(Left(strTmp02, (intPos01 - 1)))
12840               If Right(strTmp02, 4) = " And" Then strTmp02 = Trim(Left(strTmp02, (Len(strTmp02) - 4)))
12850               strFilter = strTmp02
12860             Else
12870               strFilter = vbNullString
12880             End If
12890           End If
12900         Else
12910           strFilter = Trim(strTmp02)
12920         End If
12930       Case False
              ' ** If the filter's off, throw out anything in it.
12940         strFilter = vbNullString
12950       End Select
12960       frm.Filter = strFilter
12970       If strFilter <> vbNullString Then
12980         frm.FilterOn = True
12990       Else
13000         frm.FilterOn = False
13010       End If
13020     End Select
13030     DoCmd.Hourglass False
13040     blnRecalcTots = True
13050     .TimerInterval = 500&
13060   End With

EXITP:
13070   Set frm = Nothing
13080   Exit Sub

ERRH:
13090   DoCmd.Hourglass False
13100   Select Case ERR.Number
        Case Else
13110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13120   End Select
13130   Resume EXITP

End Sub

Private Sub chkShowUnassignedOnly_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowUnassignedOnly_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
13270       With Me
13280         intRetVal = 0
13290         .chkShowCheckedOnly.SetFocus
13300       End With
13310     End Select
13320   End If

        ' ** Shift keys.
13330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13350       With Me
13360         intRetVal = 0
13370         If .chkIncludeCurrency.Enabled = True Then
13380           .chkIncludeCurrency.SetFocus
13390         Else
13400           .chkShowJournalNo.SetFocus
13410         End If
13420       End With
13430     End Select
13440   End If

EXITP:
13450   KeyCode = intRetVal
13460   Exit Sub

ERRH:
13470   intRetVal = 0
13480   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
13490   Case Else
13500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13510   End Select
13520   Resume EXITP

End Sub

Private Sub chkShowCheckedOnly_AfterUpdate()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowCheckedOnly_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String, blnFilterOn As Boolean
        Dim intPos01 As Integer, intPos02 As Integer
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String

13610   With Me
13620     varTmp00 = Nz(.chkTaxChk_sum, 0)
13630     If varTmp00 > 0 Then

            ' ** Filters always seem to get gigantic!
13640       DoCmd.Hourglass True
13650       DoEvents
13660       Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form
13670       Select Case .chkShowCheckedOnly
            Case True
13680         .chkShowCheckedOnly_lbl.FontBold = True
13690         strTmp01 = "[chkTaxChk] = True"
13700         strTmp02 = frm.Filter
13710         blnFilterOn = frm.FilterOn
13720         Select Case blnFilterOn
              Case True
13730           If Trim(strTmp02) <> vbNullString Then
13740             intPos01 = InStr(strTmp02, "[chkTaxChk]")
13750             If intPos01 > 0 Then
                    ' ** That's odd.
13760               intPos02 = InStr(intPos01, strTmp02, " And ")
13770               If intPos02 > 0 Then
13780                 strTmp03 = Trim(Mid(strTmp02, intPos02))
13790                 If intPos01 > 1 Then
13800                   strTmp02 = Trim(Left(strTmp02, (intPos01 - 1)))
13810                   strFilter = strTmp02 & " " & strTmp01 & " " & strTmp03
13820                 Else
13830                   strTmp02 = vbNullString
13840                   strFilter = strTmp01 & " " & strTmp03
13850                 End If
13860               Else
13870                 If intPos01 > 1 Then
13880                   strTmp02 = Trim(Left(strTmp02, (intPos01 - 1)))
13890                   strFilter = strTmp02 & " And " & strTmp01
13900                 Else
13910                   strTmp02 = vbNullString
13920                   strFilter = strTmp01
13930                 End If
13940               End If
13950             Else
13960               strFilter = strTmp02 & " And " & strTmp01
13970             End If
13980           Else
13990             strFilter = strTmp01
14000           End If
14010         Case False
                ' ** If the filter's off, throw out anything in it.
14020           strFilter = strTmp01
14030         End Select
              ' ** Somewhere up above we got too many And's!
14040         strFilter = StringReplace(strFilter, " And And ", " And ")
14050         If Left(Trim(strFilter), 4) = "And " Then strFilter = Trim(Mid(strFilter, 4))
14060         frm.Filter = strFilter
14070         If strFilter <> vbNullString Then
14080           frm.FilterOn = True
14090         Else
14100           frm.FilterOn = False
14110         End If
14120       Case False
14130         .chkShowCheckedOnly_lbl.FontBold = False
14140         strTmp01 = vbNullString
14150         strTmp02 = frm.Filter
14160         blnFilterOn = frm.FilterOn
14170         Select Case blnFilterOn
              Case True
14180           intPos01 = InStr(strTmp02, "[chkTaxChk]")
14190           If intPos01 > 0 Then
14200             intPos02 = InStr(intPos01, strTmp02, " And ")
14210             If intPos02 > 0 Then
14220               strTmp03 = Trim(Mid(strTmp02, intPos02))  ' ** Includes 'And'.
14230               If intPos01 = 1 Then
14240                 strTmp03 = Trim(Mid(strTmp03, 4))  ' ** Without 'And'.
14250                 strFilter = strTmp03
14260               Else
14270                 strTmp02 = Trim(Left(strTmp02, (intPos01 - 1)))  ' ** Should end with 'And'.
14280                 strFilter = StringReplace((strTmp02 & " " & strTmp03), " And And ", " And ")  ' ** Module Functions: modStringFuncs.
14290               End If
14300             Else
14310               If intPos01 = 1 Then
14320                 strFilter = vbNullString
14330               Else
14340                 strTmp02 = Trim(Left(strTmp02, (intPos01 - 1)))
14350                 If Right(strTmp02, 4) = " And" Then strTmp02 = Trim(Left(strTmp02, (Len(strTmp02) - 4)))
14360                 strFilter = strTmp02
14370               End If
14380             End If
14390           Else
                  ' ** That's odd.
14400             strFilter = strTmp02
14410           End If
14420         Case False
                ' ** If the filter's off, throw out anything in it.
14430           strFilter = vbNullString
14440         End Select
14450         frm.Filter = strFilter
14460         If strFilter <> vbNullString Then
14470           frm.FilterOn = True
14480         Else
14490           frm.FilterOn = False
14500         End If
14510       End Select
14520       DoCmd.Hourglass False
14530       blnRecalcTots = True
14540       .TimerInterval = 500&

14550     Else
14560       Beep
14570       .chkShowCheckedOnly = False
14580       .chkShowCheckedOnly_lbl.FontBold = False
14590     End If
14600   End With

EXITP:
14610   Set frm = Nothing
14620   Exit Sub

ERRH:
14630   DoCmd.Hourglass False
14640   Select Case ERR.Number
        Case Else
14650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14660   End Select
14670   Resume EXITP

End Sub

Private Sub chkShowCheckedOnly_KeyDown(KeyCode As Integer, Shift As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowCheckedOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14720   intShiftDown = (Shift And acShiftMask) > 0
14730   intAltDown = (Shift And acAltMask) > 0
14740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14770       With Me
14780         intRetVal = 0
14790         If .cmdClear.Enabled = True Then
14800           .cmdClear.SetFocus
14810         ElseIf .cmdAssign.Enabled = True Then
14820           .cmdAssign.SetFocus
14830         ElseIf .cmdOK.Enabled = True Then
14840           .cmdOK.SetFocus
14850         Else
14860           .cmdCancel.SetFocus
14870         End If
14880       End With
14890     End Select
14900   End If

        ' ** Shift keys.
14910   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14920     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14930       With Me
14940         intRetVal = 0
14950         .chkShowUnassignedOnly.SetFocus
14960       End With
14970     End Select
14980   End If

EXITP:
14990   KeyCode = intRetVal
15000   Exit Sub

ERRH:
15010   intRetVal = 0
15020   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
15030   Case Else
15040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15050   End Select
15060   Resume EXITP

End Sub

Private Sub cmdClear_Click()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClear_Click"

        Dim rst As DAO.Recordset, frm As Access.Form
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long, lngRecsChanged As Long
        Dim blnRetVal As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long

15110   With Me
15120     DoCmd.Hourglass True
15130     DoEvents
15140     blnRetVal = True
15150     Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form
15160     varTmp00 = frm.chkTaxChk_sum
15170     If IsNull(varTmp00) = True Then
            ' ** Actually, they couldn't have gotten here if it was Zero.
15180       frm.RecalcTots  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
15190       DoEvents
15200     End If
15210     If Abs(varTmp00) > 0 Then
15220       DoCmd.Hourglass False
15230       Beep
15240       msgResponse = MsgBox("This will remove the current Tax Code" & vbCrLf & _
              "from all checked entries." & vbCrLf & vbCrLf & "Proceed?", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Clear Tax Codes")
15250       If msgResponse = vbYes Then

15260         DoCmd.Hourglass True
15270         DoEvents

15280         lngRecsChanged = 0&

15290         Set rst = frm.RecordsetClone
15300         With rst
15310           If .BOF = True And .EOF = True Then
                  ' ** Oh-oh.
15320             blnRetVal = False
15330             DoCmd.Hourglass False
15340             MsgBox "Expected records were not found, and Trust Accountant cannot continue.", vbCritical + vbOKOnly, "Records Not Found"
15350           Else
15360             .MoveLast
15370             lngRecs = .RecordCount
15380             .MoveFirst
15390             For lngX = 1& To lngRecs
15400               If ![chkTaxChk] = True Then  ' ** Make sure we only do this to checked items!
15410                 Select Case ![taxcode_type]
                      Case TAXTYP_INC
15420                   .Edit
15430                   ![taxcode] = TAXID_INC
15440                   ![tcbas_changed] = True
15450                   ![tcbas_datemodified] = Now()
15460                   .Update
15470                   lngRecsChanged = lngRecsChanged + 1&
15480                 Case TAXTYP_DED
15490                   .Edit
15500                   ![taxcode] = TAXID_DED
15510                   ![tcbas_changed] = True
15520                   ![tcbas_datemodified] = Now()
15530                   .Update
15540                   lngRecsChanged = lngRecsChanged + 1&
15550                 End Select
15560               End If
15570               If lngX < lngRecs Then .MoveNext
15580             Next
15590           End If
15600           .Close
15610         End With
15620         Set rst = Nothing

15630         If blnRetVal = True Then

15640           .cmdOK.Enabled = True

15650           DoCmd.Hourglass False
15660           Beep
15670           MsgBox CStr(lngRecsChanged) & " transaction" & IIf(lngRecsChanged = 1&, " was ", "s were ") & _
                  "reset to Unspecified." & vbCrLf & vbCrLf & "Changes will not be permanent until OK is pressed.", _
                  vbInformation + vbOKOnly, "Entries Cleared"

15680         End If

15690       End If

15700     Else
15710       DoCmd.Hourglass False
15720       MsgBox "There are no checked items to clear.", vbInformation + vbOKOnly, "Nothing To Do"
15730     End If
15740     blnRecalcTots = True
15750     .TimerInterval = 500&
15760   End With

EXITP:
15770   Set frm = Nothing
15780   Set rst = Nothing
15790   Exit Sub

ERRH:
15800   DoCmd.Hourglass False
15810   Select Case ERR.Number
        Case Else
15820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15830   End Select
15840   Resume EXITP

End Sub

Private Sub cmdClear_KeyDown(KeyCode As Integer, Shift As Integer)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClear_KeyDown"

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
15990         If .cmdAssign.Enabled = True Then
16000           .cmdAssign.SetFocus
16010         ElseIf .cmdOK.Enabled = True Then
16020           .cmdOK.SetFocus
16030         Else
16040           .cmdCancel.SetFocus
16050         End If
16060       End With
16070     End Select
16080   End If

        ' ** Shift keys.
16090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16100     Select Case intRetVal
          Case vbKeyTab
16110       With Me
16120         intRetVal = 0
16130         .chkShowCheckedOnly.SetFocus
16140       End With
16150     End Select
16160   End If

EXITP:
16170   KeyCode = intRetVal
16180   Exit Sub

ERRH:
16190   intRetVal = 0
16200   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
16210   Case Else
16220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16230   End Select
16240   Resume EXITP

End Sub

Private Sub cmdAssign_Click()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssign_Click"

        Dim frm As Access.Form, rst As DAO.Recordset
        Dim lngTaxcode As Long, lngTaxcodeType As Long
        Dim strMsg As String, strTaxcodeDesc As String
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long, lngRecsChanged As Long
        Dim blnRetVal As Boolean
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long

        'TAXID_INC
        'TAXID_DED
        'TAXTYP_INC
        'TAXTYP_DED
16310   With Me
16320     DoCmd.Hourglass True
16330     DoEvents
16340     If IsNull(.cmbTaxCodes) = False Then
16350       blnRetVal = True
16360       lngTaxcode = .cmbTaxCodes.Column(0)
16370       strTaxcodeDesc = .cmbTaxCodes.Column(1)
16380       lngTaxcodeType = .cmbTaxCodes.Column(2)
16390       Set frm = .frmAccountTaxCodes_BlockAssign_Sub.Form
16400       varTmp00 = frm.chkTaxChk_sum
16410       If IsNull(varTmp00) = True Then
              ' ** Actually, they couldn't have gotten here if it was Zero.
16420         frm.RecalcTots  ' ** Form Procedure: frmAccountTaxCodes_BlockAssign_Sub.
16430         DoEvents
16440       End If
16450       If Abs(varTmp00) > 0 Then
16460         varTmp00 = Abs(varTmp00)
16470         Select Case lngTaxcodeType
              Case TAXTYP_INC
16480           strTmp01 = "Income"
16490         Case TAXTYP_DED
16500           strTmp01 = "Deduction"
16510         End Select
16520         strMsg = "This will assign the chosen " & strTmp01 & " code" & vbCrLf & "to the " & CStr(varTmp00) & " " & _
                "checked transaction" & IIf(varTmp00 = 1, " ", "s ") & "in the list." & vbCrLf & vbCrLf & _
                "Proper income vs. deduction type will be checked," & vbCrLf & _
                "so final tally may differ." & vbCrLf & vbCrLf & "Proceed?"
16530         DoCmd.Hourglass False
16540         Beep
16550         msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Assign Tax Codes")
16560         If msgResponse = vbYes Then

16570           DoCmd.Hourglass True
16580           DoEvents

16590           Set rst = frm.RecordsetClone
16600           With rst
16610             If .BOF = True And .EOF = True Then
                    ' ** Oh-oh.
16620               blnRetVal = False
16630               DoCmd.Hourglass False
16640               MsgBox "Expected records were not found, and Trust Accountant cannot continue.", vbCritical + vbOKOnly, "Records Not Found"
16650             Else
16660               .MoveLast
16670               lngRecs = .RecordCount
16680               .MoveFirst

16690               lngRecsChanged = 0&
16700               For lngX = 1& To lngRecs
16710                 If ![chkTaxChk] = True Then  ' ** Make sure we only do this to checked items!
16720                   If ![taxcode_type] = lngTaxcodeType Then
16730                     .Edit
16740                     ![taxcode] = lngTaxcode
16750                     ![taxcode_description] = strTaxcodeDesc
16760                     ![tcbas_changed] = True
16770                     ![tcbas_datemodified] = Now()
16780                     .Update
16790                     lngRecsChanged = lngRecsChanged + 1&
16800                   End If
16810                 End If
16820                 If lngX < lngRecs Then .MoveNext
16830               Next

16840             End If
16850             .Close
16860           End With
16870           Set rst = Nothing

16880           If blnRetVal = True Then

16890             .cmdOK.Enabled = True

16900             DoCmd.Hourglass False
16910             Beep
16920             MsgBox CStr(lngRecsChanged) & " transaction" & IIf(lngRecsChanged = 1&, " was ", "s were ") & _
                    "updated." & vbCrLf & vbCrLf & "Changes will not be permanent until OK is pressed.", _
                    vbInformation + vbOKOnly, "Tax Codes Changed"

16930           End If

16940         End If
16950       Else
16960         DoCmd.Hourglass False
16970         MsgBox "There are no checked items to assign.", vbInformation + vbOKOnly, "Nothing To Do"
16980       End If
16990     Else
17000       DoCmd.Hourglass False
17010       MsgBox "An Income or Deduction Code must be chosen to assign.", vbInformation + vbOKOnly, "Nothing To Do"
17020     End If
17030     blnRecalcTots = True
17040     .TimerInterval = 500&
17050   End With

EXITP:
17060   Set frm = Nothing
17070   Set rst = Nothing
17080   Exit Sub

ERRH:
17090   Select Case ERR.Number
        Case Else
17100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17110   End Select
17120   Resume EXITP

End Sub

Private Sub cmdAssign_KeyDown(KeyCode As Integer, Shift As Integer)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssign_KeyDown"

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
17290         If .cmdOK.Enabled = True Then
17300           .cmdOK.SetFocus
17310         Else
17320           .cmdCancel.SetFocus
17330         End If
17340       End With
17350     End Select
17360   End If

        ' ** Shift keys.
17370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17380     Select Case intRetVal
          Case vbKeyTab
17390       With Me
17400         intRetVal = 0
17410         If .cmdClear.Enabled = True Then
17420           .cmdClear.SetFocus
17430         Else
17440           .chkShowCheckedOnly.SetFocus
17450         End If
17460       End With
17470     End Select
17480   End If

EXITP:
17490   KeyCode = intRetVal
17500   Exit Sub

ERRH:
17510   intRetVal = 0
17520   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
17530   Case Else
17540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17550   End Select
17560   Resume EXITP

End Sub

Private Sub ShowJournalNo(blnShow As Boolean)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo"

        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long

17610   With Me

17620     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
17630       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
17640     End If

17650     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
17660     lngMonitorNum = 1&: lngTmp02 = 0&
17670     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
17680     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

17690     Select Case blnShow
          Case True

17700       Select Case .chkIncludeCurrency
            Case True
17710         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
17720         DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, lngFrm_Height  'lngFrm_Top
17730         If lngMonitorNum > 1& Then
17740           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
17750         End If
17760         .Width = lngForm_Width
17770       Case False
17780         lngTmp01 = (lngCurrID_Width / 2&)
17790         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
17800         DoCmd.MoveSize (lngFrm_Left + lngTmp01), lngTmp02, (lngFrm_Width - lngCurrID_Width), lngFrm_Height  'lngFrm_Top
17810         If lngMonitorNum > 1& Then
17820           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
17830         End If
17840         .Width = (lngForm_Width - lngCurrID_Width)
17850       End Select

17860       For lngX = (lngCtls - 1&) To 0& Step -1&
17870         If arr_varCtl(C_ISFRM, lngX) = False Then
17880           .Controls(arr_varCtl(C_CNAM, lngX)).Left = arr_varCtl(C_LFT, lngX)
17890           If IsNull(arr_varCtl(C_L1_NAM, lngX)) = False Then
17900             .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = arr_varCtl(C_L1_LFT, lngX)
17910           End If
17920           If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
17930             .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = arr_varCtl(C_L2_LFT, lngX)
17940           End If
17950           If Left(arr_varCtl(C_CNAM, lngX), 9) = "cmdSelect" Then
17960             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_img").Left = arr_varCtl(C_LFT, lngX)
17970             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_semifocus_dots_img").Left = arr_varCtl(C_LFT, lngX)
17980             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_focus_img").Left = arr_varCtl(C_LFT, lngX)
17990             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_focus_dots_img").Left = arr_varCtl(C_LFT, lngX)
18000             .Controls(arr_varCtl(C_CNAM, lngX) & "_sunken_focus_dots_img").Left = arr_varCtl(C_LFT, lngX)
18010             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_img_dis").Left = arr_varCtl(C_LFT, lngX)
18020           End If
18030         End If
18040       Next

18050       Select Case .chkIncludeCurrency
            Case True
18060         lngTmp01 = 0&
18070       Case False
18080         lngTmp01 = lngCurrID_Width
18090       End Select

18100       For lngX = (lngCtls - 1&) To 0& Step -1&
18110         If arr_varCtl(C_ISFRM, lngX) = True Then
18120           If Left(arr_varCtl(C_CNAM, lngX), 2) = "cm" Then
                  ' ** Left adjustment.
18130             .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) - lngTmp01)
18140             If IsNull(arr_varCtl(C_L1_NAM, lngX)) = False Then
18150               .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = (arr_varCtl(C_L1_LFT, lngX) - lngTmp01)
18160             End If
18170             If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
18180               .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = (arr_varCtl(C_L2_LFT, lngX) - lngTmp01)
18190             End If
18200             If arr_varCtl(C_CNAM, lngX) = "cmbTaxCodes" Then
18210               .cmbTaxCodes_lbl3.Left = .cmbTaxCodes_lbl2.Left
                    'lngTmp02 = (.cmbTaxCodes_lbl3.Left + .cmbTaxCodes_lbl3.Width)
                    'lngTmp02 = (lngTmp02 - .cmbTaxCodes_lbl4.Width)
                    '.cmbTaxCodes_lbl4.Left = lngTmp02
                    '.cmbTaxCodes_lbl5.Left = lngTmp02
                    '.cmbTaxCodes_lbl6.Left = lngTmp02
18220               .cmbTaxCodes.Width = lngTaxCodes_Width
18230             End If
18240           Else
                  ' ** Width adjustment.
18250             If Left(arr_varCtl(C_CNAM, lngX), 3) = "frm" Then  ' ** Note!
18260               .Controls(arr_varCtl(C_CNAM, lngX)).Width = (arr_varCtl(C_LFT, lngX) - lngTmp01)
18270             Else
                    'DOES ANYTHING HIT THIS?
18280               .Controls(arr_varCtl(C_CNAM, lngX)).Width = (lngFrm_Width - lngTmp01)
18290             End If
18300           End If
18310         End If
18320       Next

            '.cmbTaxCodes_lbl4.Visible = False
            '.cmbTaxCodes_lbl5.Visible = False
            '.cmbTaxCodes_lbl6.Visible = False
18330       .cmbTaxCodes_lbl.Visible = True
18340       .cmbTaxCodes_lbl2.Visible = True
18350       .cmbTaxCodes_lbl3.Visible = True

18360       Select Case .chkIncludeCurrency
            Case True
18370         lngTmp01 = lngForm_Width
18380       Case False
18390         lngTmp01 = (lngForm_Width - lngCurrID_Width)
18400       End Select

18410       .Header_vline01.Left = lngTmp01
18420       .Header_vline02.Left = lngTmp01
18430       .Header_hline01.Width = lngTmp01
18440       .Header_hline02.Width = lngTmp01
18450       .Detail_vline01.Left = lngTmp01
18460       .Detail_vline02.Left = lngTmp01
18470       .Detail_hline01.Width = lngTmp01
18480       .Detail_hline02.Width = lngTmp01
18490       .Footer_vline01.Left = lngTmp01
18500       .Footer_vline02.Left = lngTmp01
18510       .Footer_hline01.Width = lngTmp01
18520       .Footer_hline02.Width = lngTmp01

18530       If .chkIncludeCurrency = False Then
18540         .Width = (lngForm_Width - lngCurrID_Width)
18550       End If

18560     Case False

18570       For lngX = 0& To (lngCtls - 1&)
18580         If arr_varCtl(C_ISFRM, lngX) = False Then
18590           .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) - lngJournalNo_Width)
18600           If IsNull(arr_varCtl(C_L1_NAM, lngX)) = False Then
18610             .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = (arr_varCtl(C_L1_LFT, lngX) - lngJournalNo_Width)
18620           End If
18630           If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
18640             .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = (arr_varCtl(C_L2_LFT, lngX) - lngJournalNo_Width)
18650           End If
18660           If Left(arr_varCtl(C_CNAM, lngX), 9) = "cmdSelect" Then
18670             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_img").Left = (arr_varCtl(C_LFT, lngX) - lngJournalNo_Width)
18680             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_semifocus_dots_img").Left = (arr_varCtl(C_LFT, lngX) - lngJournalNo_Width)
18690             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_focus_img").Left = (arr_varCtl(C_LFT, lngX) - lngJournalNo_Width)
18700             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_focus_dots_img").Left = (arr_varCtl(C_LFT, lngX) - lngJournalNo_Width)
18710             .Controls(arr_varCtl(C_CNAM, lngX) & "_sunken_focus_dots_img").Left = (arr_varCtl(C_LFT, lngX) - lngJournalNo_Width)
18720             .Controls(arr_varCtl(C_CNAM, lngX) & "_raised_img_dis").Left = (arr_varCtl(C_LFT, lngX) - lngJournalNo_Width)
18730           End If
18740         End If
18750       Next

18760       Select Case .chkIncludeCurrency
            Case True
18770         lngTmp01 = 0&
18780       Case False
18790         lngTmp01 = lngCurrID_Width
18800       End Select

18810       For lngX = (lngCtls - 1&) To 0& Step -1&
18820         If arr_varCtl(C_ISFRM, lngX) = True Then
18830           If Left(arr_varCtl(C_CNAM, lngX), 2) = "cm" Then
                  ' ** Left adjustment.
18840             .Controls(arr_varCtl(C_CNAM, lngX)).Left = ((arr_varCtl(C_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
18850             If IsNull(arr_varCtl(C_L1_NAM, lngX)) = False Then
18860               .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = ((arr_varCtl(C_L1_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
18870             End If
18880             If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
18890               .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = ((arr_varCtl(C_L2_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
18900             End If
18910             If arr_varCtl(C_CNAM, lngX) = "cmbTaxCodes" Then
18920               .cmbTaxCodes_lbl3.Left = .cmbTaxCodes_lbl2.Left
                    'lngTmp02 = (.cmbTaxCodes_lbl3.Left + .cmbTaxCodes_lbl3.Width)
                    'lngTmp02 = (lngTmp02 - .cmbTaxCodes_lbl4.Width)
                    '.cmbTaxCodes_lbl4.Left = lngTmp02
                    '.cmbTaxCodes_lbl5.Left = lngTmp02
                    '.cmbTaxCodes_lbl6.Left = lngTmp02
18930               .cmbTaxCodes.Width = lngTaxCodes_Width
18940             End If
18950           Else
                  ' ** Width adjustment.
18960             If Left(arr_varCtl(C_CNAM, lngX), 3) = "frm" Then  ' ** Note!
18970               .Controls(arr_varCtl(C_CNAM, lngX)).Width = ((arr_varCtl(C_LFT, lngX) - lngJournalNo_Width) - lngTmp01)
18980             Else
                    'DOES ANYTHING HIT THIS?
18990               .Controls(arr_varCtl(C_CNAM, lngX)).Width = ((lngFrm_Width - lngJournalNo_Width) - lngTmp01)
19000             End If
19010           End If
19020         End If
19030       Next

19040       Select Case .chkIncludeCurrency
            Case True
              '.cmbTaxCodes_lbl4.Visible = False
              '.cmbTaxCodes_lbl5.Visible = False
              '.cmbTaxCodes_lbl6.Visible = False
19050         .cmbTaxCodes_lbl.Visible = True
19060         .cmbTaxCodes_lbl2.Visible = True
19070         .cmbTaxCodes_lbl3.Visible = True
19080         .cmbTaxCodes.Width = lngTaxCodes_Width
19090       Case False
19100         .cmbTaxCodes_lbl.Visible = True
19110         .cmbTaxCodes_lbl2.Visible = True
19120         .cmbTaxCodes_lbl3.Visible = True
              '.cmbTaxCodes_lbl4.Visible = True
              '.cmbTaxCodes_lbl5.Visible = True
              '.cmbTaxCodes_lbl6.Visible = True
              '.cmbTaxcodes.Width = (lngTaxCodes_Width - (4& * lngTpp))
              '.cmbTaxcodes.Left = (.cmbTaxcodes.Left + (4& * lngTpp))
              '.cmbTaxCodes_lbl4.Left = (.cmbTaxCodes_lbl4.Left + (4& * lngTpp))
              '.cmbTaxCodes_lbl5.Left = (.cmbTaxCodes_lbl5.Left + (4& * lngTpp))
              '.cmbTaxCodes_lbl6.Left = (.cmbTaxCodes_lbl6.Left + (4& * lngTpp))
19130       End Select

19140       Select Case .chkIncludeCurrency
            Case True
19150         lngTmp01 = (lngForm_Width - lngJournalNo_Width)
19160       Case False
19170         lngTmp01 = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
19180       End Select

19190       .Header_vline01.Left = lngTmp01
19200       .Header_vline02.Left = lngTmp01
19210       .Header_hline01.Width = lngTmp01
19220       .Header_hline02.Width = lngTmp01
19230       .Detail_vline01.Left = lngTmp01
19240       .Detail_vline02.Left = lngTmp01
19250       .Detail_hline01.Width = lngTmp01
19260       .Detail_hline02.Width = lngTmp01
19270       .Footer_vline01.Left = lngTmp01
19280       .Footer_vline02.Left = lngTmp01
19290       .Footer_hline01.Width = lngTmp01
19300       .Footer_hline02.Width = lngTmp01

19310       Select Case .chkIncludeCurrency
            Case True
19320         lngTmp01 = (lngJournalNo_Width / 2&)
19330         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
19340         DoCmd.MoveSize (lngFrm_Left + lngTmp01), lngTmp02, (lngFrm_Width - lngJournalNo_Width), lngFrm_Height  'lngFrm_Top
19350         If lngMonitorNum > 1& Then
19360           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
19370         End If
19380         .Width = (lngForm_Width - lngJournalNo_Width)
19390       Case False
19400         lngTmp01 = ((lngCurrID_Width + lngJournalNo_Width) / 2&)
19410         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
19420         DoCmd.MoveSize (lngFrm_Left + lngTmp01), lngTmp02, ((lngFrm_Width - lngJournalNo_Width) - lngCurrID_Width), lngFrm_Height  'lngFrm_Top
19430         If lngMonitorNum > 1& Then
19440           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
19450         End If
19460         .Width = ((lngForm_Width - lngJournalNo_Width) - lngCurrID_Width)
19470       End Select

19480     End Select

19490   End With

EXITP:
19500   Exit Sub

ERRH:
19510   Select Case ERR.Number
        Case Else
19520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19530   End Select
19540   Resume EXITP

End Sub

Private Sub IncludeCurrency(blnInclude As Boolean)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long

19610   With Me

19620     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
19630       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
19640     End If

19650     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
19660     lngMonitorNum = 1&: lngTmp02 = 0&
19670     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
19680     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

19690     Select Case blnInclude
          Case True

19700       Select Case .chkShowJournalNo
            Case True
19710         lngTmp01 = 0&
19720         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
19730         DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, lngFrm_Height  'lngFrm_Top
19740         If lngMonitorNum > 1& Then
19750           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
19760         End If
19770       Case False
19780         lngTmp01 = (lngJournalNo_Width / 2&)
19790         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
19800         DoCmd.MoveSize (lngFrm_Left + lngTmp01), lngTmp02, (lngFrm_Width - lngJournalNo_Width), lngFrm_Height  'lngFrm_Top
19810         If lngMonitorNum > 1& Then
19820           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
19830         End If
19840       End Select

19850       Select Case .chkShowJournalNo
            Case True
19860         lngTmp01 = 0&
19870       Case False
19880         lngTmp01 = lngJournalNo_Width
19890       End Select

19900       .Width = (lngForm_Width - lngTmp01)

19910       For lngX = 0& To (lngCtls - 1&)
19920         If arr_varCtl(C_ISFRM, lngX) = True Then
19930           If Left(arr_varCtl(C_CNAM, lngX), 2) = "cm" Then
                  ' ** Left adjustment.
19940             .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) - lngTmp01)
19950             If IsNull(arr_varCtl(C_L1_NAM, lngX)) = False Then
19960               .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = (arr_varCtl(C_L1_LFT, lngX) - lngTmp01)
19970             End If
19980             If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
19990               .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = (arr_varCtl(C_L2_LFT, lngX) - lngTmp01)
20000             End If
20010             If arr_varCtl(C_CNAM, lngX) = "cmbTaxCodes" Then
20020               .cmbTaxCodes_lbl3.Left = .cmbTaxCodes_lbl2.Left
20030               .cmbTaxCodes.Width = lngTaxCodes_Width
20040             End If
20050           Else
                  ' ** Width adjustment.
20060             If Left(arr_varCtl(C_CNAM, lngX), 3) = "frm" Then  ' ** Note!
20070               .Controls(arr_varCtl(C_CNAM, lngX)).Width = (arr_varCtl(C_LFT, lngX) - lngTmp01)
20080             Else
                    ' ** There shouldn't be anything else.
20090             End If
20100           End If
20110         End If
20120       Next

20130       .cmbTaxCodes_lbl.Visible = True
20140       .cmbTaxCodes_lbl2.Visible = True
20150       .cmbTaxCodes_lbl3.Visible = True

20160       Select Case .chkShowJournalNo
            Case True
20170         lngTmp01 = lngForm_Width
20180       Case False
20190         lngTmp01 = (lngForm_Width - lngJournalNo_Width)
20200       End Select

20210       .Header_vline01.Left = lngTmp01
20220       .Header_vline02.Left = lngTmp01
20230       .Header_hline01.Width = lngTmp01
20240       .Header_hline02.Width = lngTmp01
20250       .Detail_vline01.Left = lngTmp01
20260       .Detail_vline02.Left = lngTmp01
20270       .Detail_hline01.Width = lngTmp01
20280       .Detail_hline02.Width = lngTmp01
20290       .Footer_vline01.Left = lngTmp01
20300       .Footer_vline02.Left = lngTmp01
20310       .Footer_hline01.Width = lngTmp01
20320       .Footer_hline02.Width = lngTmp01

20330     Case False

20340       Select Case .chkShowJournalNo
            Case True
20350         lngTmp01 = lngCurrID_Width
20360       Case False
20370         lngTmp01 = (lngCurrID_Width + lngJournalNo_Width)
20380       End Select

20390       For lngX = 0& To (lngCtls - 1&)
20400         If arr_varCtl(C_ISFRM, lngX) = True Then
20410           If Left(arr_varCtl(C_CNAM, lngX), 2) = "cm" Then
                  ' ** Left adjustment.
20420             .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) - lngTmp01)
20430             If IsNull(arr_varCtl(C_L1_NAM, lngX)) = False Then
20440               .Controls(arr_varCtl(C_L1_NAM, lngX)).Left = (arr_varCtl(C_L1_LFT, lngX) - lngTmp01)
20450             End If
20460             If IsNull(arr_varCtl(C_L2_NAM, lngX)) = False Then
20470               .Controls(arr_varCtl(C_L2_NAM, lngX)).Left = (arr_varCtl(C_L2_LFT, lngX) - lngTmp01)
20480             End If
20490             If arr_varCtl(C_CNAM, lngX) = "cmbTaxCodes" Then
20500               .cmbTaxCodes_lbl3.Left = .cmbTaxCodes_lbl2.Left
20510               .cmbTaxCodes.Width = lngTaxCodes_Width
20520             End If
20530           Else
                  ' ** Width adjustment.
20540             If Left(arr_varCtl(C_CNAM, lngX), 3) = "frm" Then  ' ** Note!
20550               .Controls(arr_varCtl(C_CNAM, lngX)).Width = (arr_varCtl(C_LFT, lngX) - lngTmp01)
20560             Else
                    ' ** There shouldn't be anything else.
20570             End If
20580           End If
20590         End If
20600       Next

20610       Select Case .chkShowJournalNo
            Case True
20620         .cmbTaxCodes_lbl.Visible = True
20630         .cmbTaxCodes_lbl2.Visible = True
20640         .cmbTaxCodes_lbl3.Visible = True
20650       Case False
20660         .cmbTaxCodes_lbl.Visible = True
20670         .cmbTaxCodes_lbl2.Visible = True
20680         .cmbTaxCodes_lbl3.Visible = True
20690       End Select

20700       Select Case .chkShowJournalNo
            Case True
20710         lngTmp01 = (lngForm_Width - lngCurrID_Width)
20720       Case False
20730         lngTmp01 = ((lngForm_Width - lngCurrID_Width) - lngJournalNo_Width)
20740       End Select

20750       .Header_vline01.Left = lngTmp01
20760       .Header_vline02.Left = lngTmp01
20770       .Header_hline01.Width = lngTmp01
20780       .Header_hline02.Width = lngTmp01
20790       .Detail_vline01.Left = lngTmp01
20800       .Detail_vline02.Left = lngTmp01
20810       .Detail_hline01.Width = lngTmp01
20820       .Detail_hline02.Width = lngTmp01
20830       .Footer_vline01.Left = lngTmp01
20840       .Footer_vline02.Left = lngTmp01
20850       .Footer_hline01.Width = lngTmp01
20860       .Footer_hline02.Width = lngTmp01

20870       .Width = lngTmp01

20880       Select Case .chkShowJournalNo
            Case True
20890         lngTmp01 = (lngCurrID_Width / 2&)
20900         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
20910         DoCmd.MoveSize (lngFrm_Left + lngTmp01), lngTmp02, (lngFrm_Width - lngCurrID_Width), lngFrm_Height  'lngFrm_Top
20920         If lngMonitorNum > 1& Then
20930           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
20940         End If
20950       Case False
20960         lngTmp01 = ((lngCurrID_Width + lngJournalNo_Width) / 2&)
20970         If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
20980         DoCmd.MoveSize (lngFrm_Left + lngTmp01), lngTmp02, ((lngFrm_Width - lngCurrID_Width) - lngJournalNo_Width), lngFrm_Height  'lngFrm_Top
20990         If lngMonitorNum > 1& Then
21000           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
21010         End If
21020       End Select

21030     End Select

21040   End With

EXITP:
21050   Exit Sub

ERRH:
21060   Select Case ERR.Number
        Case Else
21070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21080   End Select
21090   Resume EXITP

End Sub

Private Sub CtlArrayLoad()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "CtlArrayLoad"

        Dim lngE As Long

21110   With Me
21120     If lngCtls = 0& Or IsEmpty(arr_varCtl) = True Then

21130       lngCtls = 0&
21140       ReDim arr_varCtl(C_ELEMS, 0)

            ' **********************************************
            ' ** Array: arr_varCtl()
            ' **
            ' **   Field  Element  Name         Constant
            ' **   =====  =======  ===========  ==========
            ' **     1       0     ctl name     C_CNAM
            ' **     2       1     clt left     C_LFT
            ' **     3       2     clt top      C_TOP
            ' **     4       3     lbl1 name    C_L1_NAM
            ' **     5       4     lbl1 left    C_L1_LFT
            ' **     6       5     lbl1 top     C_L1_TOP
            ' **     7       6     lbl2 name    C_L2_NAM
            ' **     8       7     lbl2 left    C_L2_LFT
            ' **     9       8     lbl2 top     C_L2_TOP
            ' **    10       9     whole frm    C_ISFRM
            ' **
            ' **********************************************

21150       lngCtls = lngCtls + 1&
21160       lngE = lngCtls - 1&
21170       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
21180       arr_varCtl(C_CNAM, lngE) = chkTaxChk_box.Name
21190       arr_varCtl(C_LFT, lngE) = .chkTaxChk_box.Left
21200       arr_varCtl(C_TOP, lngE) = .chkTaxChk_box.Top
21210       arr_varCtl(C_L1_NAM, lngE) = Null
21220       arr_varCtl(C_L1_LFT, lngE) = Null
21230       arr_varCtl(C_L1_Top, lngE) = Null
21240       arr_varCtl(C_L2_NAM, lngE) = Null
21250       arr_varCtl(C_L2_LFT, lngE) = Null
21260       arr_varCtl(C_L2_TOP, lngE) = Null
21270       arr_varCtl(C_ISFRM, lngE) = CBool(False)

21280       lngCtls = lngCtls + 1&
21290       lngE = lngCtls - 1&
21300       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
21310       arr_varCtl(C_CNAM, lngE) = .chkTaxChk_sum.Name
21320       arr_varCtl(C_LFT, lngE) = .chkTaxChk_sum.Left
21330       arr_varCtl(C_TOP, lngE) = .chkTaxChk_sum.Top
21340       arr_varCtl(C_L1_NAM, lngE) = .chkTaxChk_sum_lbl.Name
21350       arr_varCtl(C_L1_LFT, lngE) = .chkTaxChk_sum_lbl.Left
21360       arr_varCtl(C_L1_Top, lngE) = .chkTaxChk_sum_lbl.Top
21370       arr_varCtl(C_L2_NAM, lngE) = Null
21380       arr_varCtl(C_L2_LFT, lngE) = Null
21390       arr_varCtl(C_L2_TOP, lngE) = Null
21400       arr_varCtl(C_ISFRM, lngE) = CBool(False)

21410       lngCtls = lngCtls + 1&
21420       lngE = lngCtls - 1&
21430       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
21440       arr_varCtl(C_CNAM, lngE) = .cmdSelectAll.Name
21450       arr_varCtl(C_LFT, lngE) = .cmdSelectAll.Left
21460       arr_varCtl(C_TOP, lngE) = .cmdSelectAll.Top
21470       arr_varCtl(C_L1_NAM, lngE) = .cmdSelectAll_lbl.Name
21480       arr_varCtl(C_L1_LFT, lngE) = .cmdSelectAll_lbl.Left
21490       arr_varCtl(C_L1_Top, lngE) = .cmdSelectAll_lbl.Top
21500       arr_varCtl(C_L2_NAM, lngE) = Null
21510       arr_varCtl(C_L2_LFT, lngE) = Null
21520       arr_varCtl(C_L2_TOP, lngE) = Null
21530       arr_varCtl(C_ISFRM, lngE) = CBool(False)

21540       lngCtls = lngCtls + 1&
21550       lngE = lngCtls - 1&
21560       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
21570       arr_varCtl(C_CNAM, lngE) = .cmdSelectNone.Name
21580       arr_varCtl(C_LFT, lngE) = .cmdSelectNone.Left
21590       arr_varCtl(C_TOP, lngE) = .cmdSelectNone.Top
21600       arr_varCtl(C_L1_NAM, lngE) = .cmdSelectNone_lbl.Name
21610       arr_varCtl(C_L1_LFT, lngE) = .cmdSelectNone_lbl.Left
21620       arr_varCtl(C_L1_Top, lngE) = .cmdSelectNone_lbl.Top
21630       arr_varCtl(C_L2_NAM, lngE) = Null
21640       arr_varCtl(C_L2_LFT, lngE) = Null
21650       arr_varCtl(C_L2_TOP, lngE) = Null
21660       arr_varCtl(C_ISFRM, lngE) = CBool(False)

21670       lngCtls = lngCtls + 1&
21680       lngE = lngCtls - 1&
21690       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
21700       arr_varCtl(C_CNAM, lngE) = .cmbJournalTypes_box.Name
21710       arr_varCtl(C_LFT, lngE) = .cmbJournalTypes_box.Left
21720       arr_varCtl(C_TOP, lngE) = .cmbJournalTypes_box.Top
21730       arr_varCtl(C_L1_NAM, lngE) = Null
21740       arr_varCtl(C_L1_LFT, lngE) = Null
21750       arr_varCtl(C_L1_Top, lngE) = Null
21760       arr_varCtl(C_L2_NAM, lngE) = Null
21770       arr_varCtl(C_L2_LFT, lngE) = Null
21780       arr_varCtl(C_L2_TOP, lngE) = Null
21790       arr_varCtl(C_ISFRM, lngE) = CBool(False)

21800       lngCtls = lngCtls + 1&
21810       lngE = lngCtls - 1&
21820       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
21830       arr_varCtl(C_CNAM, lngE) = .cmbJournalTypes.Name
21840       arr_varCtl(C_LFT, lngE) = .cmbJournalTypes.Left
21850       arr_varCtl(C_TOP, lngE) = .cmbJournalTypes.Top
21860       arr_varCtl(C_L1_NAM, lngE) = .cmbJournalTypes_lbl.Name
21870       arr_varCtl(C_L1_LFT, lngE) = .cmbJournalTypes_lbl.Left
21880       arr_varCtl(C_L1_Top, lngE) = .cmbJournalTypes_lbl.Top
21890       arr_varCtl(C_L2_NAM, lngE) = .cmbJournalTypes_lbl2.Name
21900       arr_varCtl(C_L2_LFT, lngE) = .cmbJournalTypes_lbl2.Left
21910       arr_varCtl(C_L2_TOP, lngE) = .cmbJournalTypes_lbl2.Top
21920       arr_varCtl(C_ISFRM, lngE) = CBool(False)

21930       lngCtls = lngCtls + 1&
21940       lngE = lngCtls - 1&
21950       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
21960       arr_varCtl(C_CNAM, lngE) = .chkShowJournalNo_box.Name
21970       arr_varCtl(C_LFT, lngE) = .chkShowJournalNo_box.Left
21980       arr_varCtl(C_TOP, lngE) = .chkShowJournalNo_box.Top
21990       arr_varCtl(C_L1_NAM, lngE) = .chkShowJournalNo_vline01.Name
22000       arr_varCtl(C_L1_LFT, lngE) = .chkShowJournalNo_vline01.Left
22010       arr_varCtl(C_L1_Top, lngE) = .chkShowJournalNo_vline01.Top
22020       arr_varCtl(C_L2_NAM, lngE) = .chkShowJournalNo_vline02.Name
22030       arr_varCtl(C_L2_LFT, lngE) = .chkShowJournalNo_vline02.Left
22040       arr_varCtl(C_L2_TOP, lngE) = .chkShowJournalNo_vline02.Top
22050       arr_varCtl(C_ISFRM, lngE) = CBool(False)

22060       lngCtls = lngCtls + 1&
22070       lngE = lngCtls - 1&
22080       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
22090       arr_varCtl(C_CNAM, lngE) = .chkShowJournalNo.Name
22100       arr_varCtl(C_LFT, lngE) = .chkShowJournalNo.Left
22110       arr_varCtl(C_TOP, lngE) = .chkShowJournalNo.Top
22120       arr_varCtl(C_L1_NAM, lngE) = .chkShowJournalNo_lbl.Name
22130       arr_varCtl(C_L1_LFT, lngE) = .chkShowJournalNo_lbl.Left
22140       arr_varCtl(C_L1_Top, lngE) = .chkShowJournalNo_lbl.Top
22150       arr_varCtl(C_L2_NAM, lngE) = Null
22160       arr_varCtl(C_L2_LFT, lngE) = Null
22170       arr_varCtl(C_L2_TOP, lngE) = Null
22180       arr_varCtl(C_ISFRM, lngE) = CBool(False)

22190       lngCtls = lngCtls + 1&
22200       lngE = lngCtls - 1&
22210       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
22220       arr_varCtl(C_CNAM, lngE) = .chkIncludeCurrency.Name
22230       arr_varCtl(C_LFT, lngE) = .chkIncludeCurrency.Left
22240       arr_varCtl(C_TOP, lngE) = .chkIncludeCurrency.Top
22250       arr_varCtl(C_L1_NAM, lngE) = .chkIncludeCurrency_lbl.Name
22260       arr_varCtl(C_L1_LFT, lngE) = .chkIncludeCurrency_lbl.Left
22270       arr_varCtl(C_L1_Top, lngE) = .chkIncludeCurrency_lbl.Top
22280       arr_varCtl(C_L2_NAM, lngE) = Null
22290       arr_varCtl(C_L2_LFT, lngE) = Null
22300       arr_varCtl(C_L2_TOP, lngE) = Null
22310       arr_varCtl(C_ISFRM, lngE) = CBool(False)

22320       lngCtls = lngCtls + 1&
22330       lngE = lngCtls - 1&
22340       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
22350       arr_varCtl(C_CNAM, lngE) = .chkShowUnassignedOnly.Name
22360       arr_varCtl(C_LFT, lngE) = .chkShowUnassignedOnly.Left
22370       arr_varCtl(C_TOP, lngE) = .chkShowUnassignedOnly.Top
22380       arr_varCtl(C_L1_NAM, lngE) = .chkShowUnassignedOnly_lbl.Name
22390       arr_varCtl(C_L1_LFT, lngE) = .chkShowUnassignedOnly_lbl.Left
22400       arr_varCtl(C_L1_Top, lngE) = .chkShowUnassignedOnly_lbl.Top
22410       arr_varCtl(C_L2_NAM, lngE) = Null
22420       arr_varCtl(C_L2_LFT, lngE) = Null
22430       arr_varCtl(C_L2_TOP, lngE) = Null
22440       arr_varCtl(C_ISFRM, lngE) = CBool(False)

22450       lngCtls = lngCtls + 1&
22460       lngE = lngCtls - 1&
22470       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
22480       arr_varCtl(C_CNAM, lngE) = .chkShowCheckedOnly.Name
22490       arr_varCtl(C_LFT, lngE) = .chkShowCheckedOnly.Left
22500       arr_varCtl(C_TOP, lngE) = .chkShowCheckedOnly.Top
22510       arr_varCtl(C_L1_NAM, lngE) = .chkShowCheckedOnly_lbl.Name
22520       arr_varCtl(C_L1_LFT, lngE) = .chkShowCheckedOnly_lbl.Left
22530       arr_varCtl(C_L1_Top, lngE) = .chkShowCheckedOnly_lbl.Top
22540       arr_varCtl(C_L2_NAM, lngE) = Null
22550       arr_varCtl(C_L2_LFT, lngE) = Null
22560       arr_varCtl(C_L2_TOP, lngE) = Null
22570       arr_varCtl(C_ISFRM, lngE) = CBool(False)

22580       lngCtls = lngCtls + 1&
22590       lngE = lngCtls - 1&
22600       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
22610       arr_varCtl(C_CNAM, lngE) = .cmbTaxCodes.Name
22620       arr_varCtl(C_LFT, lngE) = .cmbTaxCodes.Left
22630       arr_varCtl(C_TOP, lngE) = Null
22640       arr_varCtl(C_L1_NAM, lngE) = .cmbTaxCodes_lbl.Name
22650       arr_varCtl(C_L1_LFT, lngE) = .cmbTaxCodes_lbl.Left
22660       arr_varCtl(C_L1_Top, lngE) = Null
22670       arr_varCtl(C_L2_NAM, lngE) = .cmbTaxCodes_lbl2.Name
22680       arr_varCtl(C_L2_LFT, lngE) = .cmbTaxCodes_lbl2.Left
22690       arr_varCtl(C_L2_TOP, lngE) = Null
22700       arr_varCtl(C_ISFRM, lngE) = CBool(True)

22710       lngCtls = lngCtls + 1&
22720       lngE = lngCtls - 1&
22730       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
22740       arr_varCtl(C_CNAM, lngE) = .cmdClear.Name
22750       arr_varCtl(C_LFT, lngE) = .cmdClear.Left
22760       arr_varCtl(C_TOP, lngE) = Null
22770       arr_varCtl(C_L1_NAM, lngE) = Null
22780       arr_varCtl(C_L1_LFT, lngE) = Null
22790       arr_varCtl(C_L1_Top, lngE) = Null
22800       arr_varCtl(C_L2_NAM, lngE) = Null
22810       arr_varCtl(C_L2_LFT, lngE) = Null
22820       arr_varCtl(C_L2_TOP, lngE) = Null
22830       arr_varCtl(C_ISFRM, lngE) = CBool(True)

22840       lngCtls = lngCtls + 1&
22850       lngE = lngCtls - 1&
22860       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
22870       arr_varCtl(C_CNAM, lngE) = .cmdAssign.Name
22880       arr_varCtl(C_LFT, lngE) = .cmdAssign.Left
22890       arr_varCtl(C_TOP, lngE) = Null
22900       arr_varCtl(C_L1_NAM, lngE) = Null
22910       arr_varCtl(C_L1_LFT, lngE) = Null
22920       arr_varCtl(C_L1_Top, lngE) = Null
22930       arr_varCtl(C_L2_NAM, lngE) = Null
22940       arr_varCtl(C_L2_LFT, lngE) = Null
22950       arr_varCtl(C_L2_TOP, lngE) = Null
22960       arr_varCtl(C_ISFRM, lngE) = CBool(True)

22970       lngCtls = lngCtls + 1&
22980       lngE = lngCtls - 1&
22990       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
23000       arr_varCtl(C_CNAM, lngE) = .cmdOK.Name
23010       arr_varCtl(C_LFT, lngE) = .cmdOK.Left
23020       arr_varCtl(C_TOP, lngE) = Null
23030       arr_varCtl(C_L1_NAM, lngE) = Null
23040       arr_varCtl(C_L1_LFT, lngE) = Null
23050       arr_varCtl(C_L1_Top, lngE) = Null
23060       arr_varCtl(C_L2_NAM, lngE) = Null
23070       arr_varCtl(C_L2_LFT, lngE) = Null
23080       arr_varCtl(C_L2_TOP, lngE) = Null
23090       arr_varCtl(C_ISFRM, lngE) = CBool(True)

23100       lngCtls = lngCtls + 1&
23110       lngE = lngCtls - 1&
23120       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
23130       arr_varCtl(C_CNAM, lngE) = .cmdCancel.Name
23140       arr_varCtl(C_LFT, lngE) = .cmdCancel.Left
23150       arr_varCtl(C_TOP, lngE) = Null
23160       arr_varCtl(C_L1_NAM, lngE) = Null
23170       arr_varCtl(C_L1_LFT, lngE) = Null
23180       arr_varCtl(C_L1_Top, lngE) = Null
23190       arr_varCtl(C_L2_NAM, lngE) = Null
23200       arr_varCtl(C_L2_LFT, lngE) = Null
23210       arr_varCtl(C_L2_TOP, lngE) = Null
23220       arr_varCtl(C_ISFRM, lngE) = CBool(True)

23230       lngCtls = lngCtls + 1&
23240       lngE = lngCtls - 1&
23250       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
23260       arr_varCtl(C_CNAM, lngE) = .frmAccountTaxCodes_BlockAssign_Sub.Name
23270       arr_varCtl(C_LFT, lngE) = .frmAccountTaxCodes_BlockAssign_Sub.Width  ' ** Note!
23280       arr_varCtl(C_TOP, lngE) = .frmAccountTaxCodes_BlockAssign_Sub.Height  ' ** Note!
23290       arr_varCtl(C_L1_NAM, lngE) = Null
23300       arr_varCtl(C_L1_LFT, lngE) = Null
23310       arr_varCtl(C_L1_Top, lngE) = Null
23320       arr_varCtl(C_L2_NAM, lngE) = Null
23330       arr_varCtl(C_L2_LFT, lngE) = Null
23340       arr_varCtl(C_L2_TOP, lngE) = Null
23350       arr_varCtl(C_ISFRM, lngE) = CBool(True)

23360       lngCtls = lngCtls + 1&
23370       lngE = lngCtls - 1&
23380       ReDim Preserve arr_varCtl(C_ELEMS, lngE)
23390       arr_varCtl(C_CNAM, lngE) = .frmAccountTaxCodes_BlockAssign_Sub_box.Name
23400       arr_varCtl(C_LFT, lngE) = .frmAccountTaxCodes_BlockAssign_Sub_box.Width  ' ** Note!
23410       arr_varCtl(C_TOP, lngE) = .frmAccountTaxCodes_BlockAssign_Sub_box.Height  ' ** Note!
23420       arr_varCtl(C_L1_NAM, lngE) = Null
23430       arr_varCtl(C_L1_LFT, lngE) = Null
23440       arr_varCtl(C_L1_Top, lngE) = Null
23450       arr_varCtl(C_L2_NAM, lngE) = Null
23460       arr_varCtl(C_L2_LFT, lngE) = Null
23470       arr_varCtl(C_L2_TOP, lngE) = Null
23480       arr_varCtl(C_ISFRM, lngE) = CBool(True)

            'chkTaxChk_box
            'chkTaxChk_sum
            'cmdSelectAll
            'cmdSelectNone
            'cmbJournalTypes_box
            'cmbJournalTypes
            'chkShowJournalNo_box
            'chkShowJournalNo
            'chkIncludeCurrency
            'chkShowUnassignedOnly
            'chkShowCheckedOnly
            'frmAccountTaxCodes_BlockAssign_Sub
            'frmAccountTaxCodes_BlockAssign_Sub_box
            'Header_
            'Detail_
            'Footer_

23490     End If
23500   End With

EXITP:
23510   Exit Sub

ERRH:
23520   Select Case ERR.Number
        Case Else
23530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23540   End Select
23550   Resume EXITP

End Sub

Public Sub SetJNoWidth(lngWidth As Long)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "SetJNoWidth"

23610   lngJournalNo_Width = lngWidth

EXITP:
23620   Exit Sub

ERRH:
23630   Select Case ERR.Number
        Case Else
23640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23650   End Select
23660   Resume EXITP

End Sub

Public Sub SetCurrWidth(lngWidth As Long)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "SetCurrWidth"

23710   lngCurrID_Width = lngWidth

EXITP:
23720   Exit Sub

ERRH:
23730   Select Case ERR.Number
        Case Else
23740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23750   End Select
23760   Resume EXITP

End Sub
