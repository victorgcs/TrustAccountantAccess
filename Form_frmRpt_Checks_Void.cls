VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Void"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Void"

'VGC 07/08/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Bank Accot Num:   A {chkbank_acctnum on frmRpt_Checks_Void_Sub
' **   Account Number:   B {accountno on frmRpt_Checks_Void_Sub
' **   Copy:             C {cmdCopy}
' **   Bank Name:        E {chkbank_name on frmRpt_Checks_Void_Sub
' **   Check Number:     K {chkvoid_chknum on frmRpt_Checks_Void_Sub
' **   Account Name:     M {shortname on frmRpt_Checks_Void_Sub
' **   Select None:      N {cmdSelectNone}
' **   OK:               O {cmdOK}
' **   Payee:            P {chkvoid_payee on frmRpt_Checks_Void_Sub
' **   Select All:       S {cmdSelectAll}
' **   Void:             V {chkvoid_void on frmRpt_Checks_Void_Sub

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:           F5 {RecalcTots}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmRpt_Checks_Void_Sub}

Private blnCopy_Focus As Boolean, blnCopy_MouseDown As Boolean
Private blnSelectAll_Focus As Boolean, blnSelectAll_MouseDown As Boolean
Private blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean

Private lngRecsCur As Long
Private strAccountNo As String, blnAll As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then

150           intPos01 = InStr(strTmp01, "~")
160           strCallingForm = Left(strTmp01, (intPos01 - 1))
170           strTmp01 = Mid(strTmp01, (intPos01 + 1))
180           intPos01 = InStr(strTmp01, "~")
190           strAccountNo = Left(strTmp01, (intPos01 - 1))
200           blnAll = CBool(Mid(strTmp01, (intPos01 + 1)))

210           .PictureData = .Form_Rpt_Checks_Void_img

220           blnCopy_Focus = False: blnCopy_MouseDown = False
230           blnSelectAll_Focus = False: blnSelectAll_MouseDown = False
240           blnSelectNone_Focus = False: blnSelectNone_MouseDown = False

250           .frmRpt_Checks_Void_Sub.Form.RecalcTots  ' ** Form Procedure: frmRpt_Checks_Void_Sub.

260           .frmRpt_Checks_Void_Sub.SetFocus

270         Else
280           Cancel = -1
290         End If
300       Else
310         Cancel = -1
320       End If
330     End With

340     DoCmd.Hourglass False

350     If Cancel = -1 Then
360       Beep
370       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
380     End If

EXITP:
390     Exit Sub

ERRH:
400     DoCmd.Hourglass False
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Account Number:   A {chkbank_acctnum on frmRpt_Checks_Void_Sub}
        ' **   Account Number:   B {accountno on frmRpt_Checks_Void_Sub}
        ' **   Bank Name:        E {chkbank_name on frmRpt_Checks_Void_Sub}
        ' **   Check Number:     K {chkvoid_chknum on frmRpt_Checks_Void_Sub}
        ' **   Account Name:     M {shortname on frmRpt_Checks_Void_Sub}
        ' **   Payee:            P {chkvoid_payee on frmRpt_Checks_Void_Sub}
        ' **   Void:             V {chkvoid_void on frmRpt_Checks_Void_Sub}

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:           F5 {RecalcTots on frmRpt_Checks_Void_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmRpt_Checks_Void_Sub}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyF5
570         With Me
580           intRetVal = 0
590           .frmRpt_Checks_Void_Sub.Form.RecalcTots  ' ** Form Procedure: frmRpt_Checks_Void_Sub.
600         End With
610       Case vbKeyEscape
            ' ** Do not allow Escape (what would it mean?).
620         intRetVal = 0
630       End Select
640     End If

        ' ** Alt keys.
650     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyA
670         With Me
680           intRetVal = 0
690           .frmRpt_Checks_Void_Sub.SetFocus
700           .frmRpt_Checks_Void_Sub.Form.chkbank_acctnum.SetFocus
710         End With
720       Case vbKeyB
730         With Me
740           intRetVal = 0
750           .frmRpt_Checks_Void_Sub.SetFocus
760           .frmRpt_Checks_Void_Sub.Form.accountno.SetFocus
770         End With
780       Case vbKeyE
790         With Me
800           intRetVal = 0
810           .frmRpt_Checks_Void_Sub.SetFocus
820           .frmRpt_Checks_Void_Sub.Form.chkbank_name.SetFocus
830         End With
840       Case vbKeyK
850         With Me
860           intRetVal = 0
870           .frmRpt_Checks_Void_Sub.SetFocus
880           .frmRpt_Checks_Void_Sub.Form.chkvoid_chknum.SetFocus
890         End With
900       Case vbKeyM
910         With Me
920           intRetVal = 0
930           .frmRpt_Checks_Void_Sub.SetFocus
940           .frmRpt_Checks_Void_Sub.Form.shortname.SetFocus
950         End With
960       Case vbKeyP
970         With Me
980           intRetVal = 0
990           .frmRpt_Checks_Void_Sub.SetFocus
1000          .frmRpt_Checks_Void_Sub.Form.chkvoid_payee.SetFocus
1010        End With
1020      Case vbKeyV
1030        With Me
1040          intRetVal = 0
1050          .frmRpt_Checks_Void_Sub.SetFocus
1060          .frmRpt_Checks_Void_Sub.Form.chkvoid_void.SetFocus
1070          .frmRpt_Checks_Void_Sub.Form.chkvoid_void = (Not .frmRpt_Checks_Void_Sub.Form.chkvoid_void)
1080          .frmRpt_Checks_Void_Sub.Form.chkvoid_void_AfterUpdate  ' ** Form Procedure frmRpt_Checks_Void_Sub.
1090        End With
1100      End Select
1110    End If

        ' ** Ctrl keys.
1120    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1130      Select Case intRetVal
          Case vbKeyS
1140        With Me
1150          intRetVal = 0
1160          .frmRpt_Checks_Void_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmRpt_Checks_Void_Sub.
1170        End With
1180      End Select
1190    End If

        ' ** Ctrl-Shift keys.
1200    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1210      Select Case intRetVal
          Case vbKeyF
1220        With Me
1230          intRetVal = 0
1240          .FocusHolder.SetFocus
1250        End With
1260      End Select
1270    End If

EXITP:
1280    KeyCode = intRetVal
1290    Exit Sub

ERRH:
1300    intRetVal = 0
1310    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1320    Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1340    End Select
1350    Resume EXITP

End Sub

Private Sub Form_Timer()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1410    With Me
1420      .TimerInterval = 0&

1430    End With

EXITP:
1440    Exit Sub

ERRH:
1450    Select Case ERR.Number
        Case Else
1460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1470    End Select
1480    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1510    DoCmd.Hourglass True
1520    DoEvents
1530    gstrReturningForm = THIS_NAME
1540    gblnSetFocus = True
1550    Select Case strCallingForm
        Case "frmRpt_Checks"
1560      If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
1570        DoCmd.SelectObject acForm, strCallingForm, False
1580      Else
1590        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
1600      End If
1610    End Select

EXITP:
1620    Exit Sub

ERRH:
1630    DoCmd.Hourglass False
1640    Select Case ERR.Number
        Case Else
1650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1660    End Select
1670    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1710    With Me
1720      If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
1730        Select Case blnSelectAll_Focus
            Case True
1740          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
1750          .cmdSelectAll_raised_img.Visible = False
1760        Case False
1770          .cmdSelectAll_raised_img.Visible = True
1780          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
1790        End Select
1800        .cmdSelectAll_raised_focus_img.Visible = False
1810        .cmdSelectAll_raised_focus_dots_img.Visible = False
1820        .cmdSelectAll_sunken_focus_dots_img.Visible = False
1830        .cmdSelectAll_raised_img_dis.Visible = False
1840      End If
1850      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
1860        Select Case blnSelectNone_Focus
            Case True
1870          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
1880          .cmdSelectNone_raised_img.Visible = False
1890        Case False
1900          .cmdSelectNone_raised_img.Visible = True
1910          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
1920        End Select
1930        .cmdSelectNone_raised_focus_img.Visible = False
1940        .cmdSelectNone_raised_focus_dots_img.Visible = False
1950        .cmdSelectNone_sunken_focus_dots_img.Visible = False
1960        .cmdSelectNone_raised_img_dis.Visible = False
1970      End If
1980      If .cmdCopy_raised_focus_dots_img.Visible = True Or .cmdCopy_raised_focus_img.Visible = True Then
1990        Select Case blnCopy_Focus
            Case True
2000          .cmdCopy_raised_semifocus_dots_img.Visible = True
2010          .cmdCopy_raised_img.Visible = False
2020        Case False
2030          .cmdCopy_raised_img.Visible = True
2040          .cmdCopy_raised_semifocus_dots_img.Visible = False
2050        End Select
2060        .cmdCopy_raised_focus_img.Visible = False
2070        .cmdCopy_raised_focus_dots_img.Visible = False
2080        .cmdCopy_sunken_focus_dots_img.Visible = False
2090        .cmdCopy_raised_img_dis.Visible = False
2100      End If
2110    End With

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub cmdOK_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngItems As Long, arr_varItem() As Variant
        Dim lngVoids As Long, lngRecs As Long
        Dim blnContinue As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim lngX As Long, lngE As Long
        Dim blnSkip As Boolean

        ' ** Array: arr_varItem().
        Const I_ELEMS As Integer = 3
        Const I_VID    As Integer = 0
        Const I_CHKNUM As Integer = 1
        Const I_BANK   As Integer = 2
        Const I_ACCT   As Integer = 3

2210    With Me

2220      blnContinue = True

2230      DoCmd.Hourglass True
2240      DoEvents

2250      lngVoids = 0&
2260      lngItems = 0&
2270      ReDim arr_varItem(I_ELEMS, 0)

          ' ** Check for presence of banks and account numbers.
2280      Set rst = .frmRpt_Checks_Void_Sub.Form.RecordsetClone
2290      If rst.BOF = True And rst.EOF = True Then
            ' ** No records!
2300        rst.Close
2310        Set rst = Nothing
2320      Else

2330        rst.MoveLast
2340        lngRecs = rst.RecordCount
2350        rst.MoveFirst

2360        For lngX = 1& To lngRecs
2370          If rst![chkvoid_void] = True Then
2380            lngVoids = lngVoids + 1&
2390            If IsNull(rst![chkbank_name]) = True Or IsNull(rst![chkbank_acctnum]) = True Then
2400              lngItems = lngItems + 1&
2410              lngE = lngItems - 1&
2420              ReDim Preserve arr_varItem(I_ELEMS, lngE)
2430              arr_varItem(I_VID, lngE) = rst![chkvoids_id]
2440              arr_varItem(I_CHKNUM, lngE) = rst![chkvoid_chknum]
2450              arr_varItem(I_BANK, lngE) = Not (IsNull(rst![chkbank_name]))     ' ** True if present.
2460              arr_varItem(I_ACCT, lngE) = Not (IsNull(rst![chkbank_acctnum]))  ' ** True if present.
2470            End If
2480          End If
2490          If lngX < lngRecs Then rst.MoveNext
2500        Next

2510        If lngItems > 0& Then

2520          blnContinue = False
2530          rst.Close
2540          Set rst = Nothing
2550          DoEvents

2560          Beep
2570          DoCmd.Hourglass False
2580          MsgBox CStr(lngItems) & " check" & IIf(lngItems = 1&, " needs ", "s need ") & "bank information." & vbCrLf & vbCrLf & _
                "Please enter the data to continue.", vbInformation + vbOKOnly, "Entry Required"

2590        ElseIf lngVoids = 0& Then

2600          rst.Close
2610          Set rst = Nothing

2620          Beep
2630          DoCmd.Hourglass False
2640          msgResponse = MsgBox("You have not chosen any checks to void." & vbCrLf & _
                "Does this mean that all checks printed correctly?" & vbCrLf & vbCrLf & _
                "If you click 'Yes' there will not be another opportunity" & vbCrLf & "    to mark one of these as void." & vbCrLf & _
                "Click 'No' if you would like to think about it a little longer.", vbQuestion + vbYesNo, "No Checks To Void")

2650          DoCmd.Hourglass True
2660          DoEvents

2670          Select Case msgResponse
              Case vbYes
                ' ** Continue without making any updates.
2680            gblnMessage = False
2690          Case Else
                ' ** Stay here a little longer.
2700            blnContinue = False
2710          End Select

2720        Else
              ' ** Continue.

2730          rst.Close
2740          Set rst = Nothing

2750          Set dbs = CurrentDb

              ' ******************************************
              ' ** Update tblCheckBank with any changes.
              ' ******************************************

              ' ** Update qryPrintChecks_05_22 (tblCheckBank, with DLookups() to
              ' ** qryPrintChecks_5_21 (qryPrintChecks_05_20 (tblCheckVoid_Staging,
              ' ** just changed = True, InChkBnk = False), linked to tblCheckBank,
              ' ** with chkbank_name_new, chkbank_acctnum_new)).
2760          Set qdf = dbs.QueryDefs("qryPrintChecks_05_23")
2770          qdf.Execute
2780          Set qdf = Nothing
2790          DoEvents

              ' *************************************
              ' ** Add new records to tblCheckBank.
              ' *************************************

              ' ** Append qryPrintChecks_05_24 (qryPrintChecks_05_20 (tblCheckVoid_Staging,
              ' ** just changed = True, InChkBnk = False), not in tblCheckBank) to tblCheckBank.
2800          Set qdf = dbs.QueryDefs("qryPrintChecks_05_25")
2810          qdf.Execute
2820          Set qdf = Nothing
2830          DoEvents

              ' *******************************
              ' ** Add voids to tblCheckVoid.
              ' *******************************

2840          glngTaxCode_Distribution = .chkvoid_void_sum  ' ** Borrowing this variable.

              ' ** Append qryPrintChecks_05_26 (tblCheckVoid_Staging, linked to tblCheckBank,
              ' ** just chkvoid_void = True, all new voids) to tblCheckVoid.
2850          Set qdf = dbs.QueryDefs("qryPrintChecks_05_27")
2860          qdf.Execute
2870          Set qdf = Nothing
2880          DoEvents

2890          dbs.Close

2900        End If

2910      End If

2920      If blnContinue = True Then
2930        DoCmd.Close acForm, THIS_NAME
2940      Else
2950        DoCmd.Hourglass False
2960      End If

2970    End With

EXITP:
2980    Set rst = Nothing
2990    Set qdf = Nothing
3000    Set dbs = Nothing
3010    Exit Sub

ERRH:
3020    DoCmd.Hourglass False
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
3170        With Me
3180          intRetVal = 0
3190          .frmRpt_Checks_Void_Sub.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab
3250        With Me
3260          intRetVal = 0
3270          If .cmdCopy.Enabled = True Then
3280            .cmdCopy.SetFocus
3290          ElseIf .cmdSelectNone.Enabled = True Then
3300            .cmdSelectNone.SetFocus
3310          Else
3320            .frmRpt_Checks_Void_Sub.SetFocus
3330          End If
3340        End With
3350      End Select
3360    End If

EXITP:
3370    KeyCode = intRetVal
3380    Exit Sub

ERRH:
3390    intRetVal = 0
3400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3410    Case Else
3420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3430    End Select
3440    Resume EXITP

End Sub

Private Sub frmRpt_Checks_Void_Sub_Enter()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "frmRpt_Checks_Void_Sub_Enter"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

3510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub MsgSubhead_lbl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "MsgSubhead_lbl1_MouseMove"

3610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub MsgSubhead_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "MsgSubhead_lbl2_MouseMove"

3710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub cmdSelectAll_GotFocus()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_GotFocus"

3810    With Me
3820      blnSelectAll_Focus = True
3830      .cmdSelectAll_raised_semifocus_dots_img.Visible = True
3840      .cmdSelectAll_raised_img.Visible = False
3850      .cmdSelectAll_raised_focus_img.Visible = False
3860      .cmdSelectAll_raised_focus_dots_img.Visible = False
3870      .cmdSelectAll_sunken_focus_dots_img.Visible = False
3880      .cmdSelectAll_raised_img_dis.Visible = False
3890    End With

EXITP:
3900    Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseDown"

4010    With Me
4020      blnSelectAll_MouseDown = True
4030      .cmdSelectAll_sunken_focus_dots_img.Visible = True
4040      .cmdSelectAll_raised_img.Visible = False
4050      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
4060      .cmdSelectAll_raised_focus_img.Visible = False
4070      .cmdSelectAll_raised_focus_dots_img.Visible = False
4080      .cmdSelectAll_raised_img_dis.Visible = False
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

Private Sub cmdSelectAll_Click()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_Click"

        Dim rst As DAO.Recordset, frm As Access.Form
        Dim lngRecs As Long
        Dim lngX As Long

4210    With Me
4220      DoCmd.Hourglass True
4230      DoEvents
4240      Set frm = .frmRpt_Checks_Void_Sub.Form
4250      lngRecsCur = frm.RecCnt  ' ** Form Function: frmRpt_Checks_Void_Sub.
4260      If lngRecsCur > 0& Then
4270        frm.cmdSave_Click  ' ** Form Procedure: frmRpt_Checks_Void_Sub.
4280        Set rst = frm.RecordsetClone
4290        With rst
4300          .MoveLast
4310          lngRecs = .RecordCount
4320          .MoveFirst
4330          For lngX = 1& To lngRecs
4340            If ![chkvoid_void] = False Then
4350              .Edit
4360              ![chkvoid_void] = True
4370              ![chkvoids_datemodified] = Now()
4380              .Update
4390            End If
4400            If lngX < lngRecs Then .MoveNext
4410          Next
4420          .Close
4430        End With
4440        frm.Refresh
4450        DoEvents
4460        frm.RecalcTots  ' ** Form Procedure: frmRpt_Checks_Void_Sub.
4470        .frmRpt_Checks_Void_Sub.SetFocus
4480      Else
4490        Beep
4500      End If
4510      DoCmd.Hourglass False
4520    End With

EXITP:
4530    Set frm = Nothing
4540    Set rst = Nothing
4550    Exit Sub

ERRH:
4560    DoCmd.Hourglass False
4570    Select Case ERR.Number
        Case Else
4580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4590    End Select
4600    Resume EXITP

End Sub

Private Sub cmdSelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_KeyDown"

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
4790          .cmdSelectNone.SetFocus
4800        End With
4810      End Select
4820    End If

        ' ** Shift keys.
4830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4840      Select Case intRetVal
          Case vbKeyTab
4850        With Me
4860          intRetVal = 0
4870          .frmRpt_Checks_Void_Sub.SetFocus
4880        End With
4890      End Select
4900    End If

EXITP:
4910    KeyCode = intRetVal
4920    Exit Sub

ERRH:
4930    intRetVal = 0
4940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4950    Case Else
4960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4970    End Select
4980    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseMove"

5010    With Me
5020      If blnSelectAll_MouseDown = False Then
5030        Select Case blnSelectAll_Focus
            Case True
5040          .cmdSelectAll_raised_focus_dots_img.Visible = True
5050          .cmdSelectAll_raised_focus_img.Visible = False
5060        Case False
5070          .cmdSelectAll_raised_focus_img.Visible = True
5080          .cmdSelectAll_raised_focus_dots_img.Visible = False
5090        End Select
5100        .cmdSelectAll_raised_img.Visible = False
5110        .cmdSelectAll_raised_semifocus_dots_img.Visible = False
5120        .cmdSelectAll_sunken_focus_dots_img.Visible = False
5130        .cmdSelectAll_raised_img_dis.Visible = False
5140      End If
5150    End With

EXITP:
5160    Exit Sub

ERRH:
5170    Select Case ERR.Number
        Case Else
5180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5190    End Select
5200    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseUp"

5310    With Me
5320      .cmdSelectAll_raised_focus_dots_img.Visible = True
5330      .cmdSelectAll_raised_img.Visible = False
5340      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
5350      .cmdSelectAll_raised_focus_img.Visible = False
5360      .cmdSelectAll_sunken_focus_dots_img.Visible = False
5370      .cmdSelectAll_raised_img_dis.Visible = False
5380      blnSelectAll_MouseDown = False
5390    End With

EXITP:
5400    Exit Sub

ERRH:
5410    Select Case ERR.Number
        Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Sub

Private Sub cmdSelectAll_LostFocus()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_LostFocus"

5510    With Me
5520      .cmdSelectAll_raised_img.Visible = True
5530      .cmdSelectAll_raised_semifocus_dots_img.Visible = False
5540      .cmdSelectAll_raised_focus_img.Visible = False
5550      .cmdSelectAll_raised_focus_dots_img.Visible = False
5560      .cmdSelectAll_sunken_focus_dots_img.Visible = False
5570      .cmdSelectAll_raised_img_dis.Visible = False
5580      blnSelectAll_Focus = False
5590    End With

EXITP:
5600    Exit Sub

ERRH:
5610    Select Case ERR.Number
        Case Else
5620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5630    End Select
5640    Resume EXITP

End Sub

Private Sub cmdSelectNone_GotFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_GotFocus"

5710    With Me
5720      blnSelectNone_Focus = True
5730      .cmdSelectNone_raised_semifocus_dots_img.Visible = True
5740      .cmdSelectNone_raised_img.Visible = False
5750      .cmdSelectNone_raised_focus_img.Visible = False
5760      .cmdSelectNone_raised_focus_dots_img.Visible = False
5770      .cmdSelectNone_sunken_focus_dots_img.Visible = False
5780      .cmdSelectNone_raised_img_dis.Visible = False
5790    End With

EXITP:
5800    Exit Sub

ERRH:
5810    Select Case ERR.Number
        Case Else
5820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5830    End Select
5840    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseDown"

5910    With Me
5920      blnSelectNone_MouseDown = True
5930      .cmdSelectNone_sunken_focus_dots_img.Visible = True
5940      .cmdSelectNone_raised_img.Visible = False
5950      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
5960      .cmdSelectNone_raised_focus_img.Visible = False
5970      .cmdSelectNone_raised_focus_dots_img.Visible = False
5980      .cmdSelectNone_raised_img_dis.Visible = False
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

Private Sub cmdSelectNone_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_Click"

        Dim rst As DAO.Recordset, frm As Access.Form
        Dim lngRecs As Long
        Dim lngX As Long

6110    With Me
6120      DoCmd.Hourglass True
6130      DoEvents
6140      Set frm = .frmRpt_Checks_Void_Sub.Form
6150      lngRecsCur = frm.RecCnt  ' ** Form Function: frmRpt_Checks_Void_Sub.
6160      If lngRecsCur > 0& Then
6170        frm.cmdSave_Click  ' ** Form Procedure: frmRpt_Checks_Void_Sub.
6180        Set rst = frm.RecordsetClone
6190        With rst
6200          .MoveLast
6210          lngRecs = .RecordCount
6220          .MoveFirst
6230          For lngX = 1& To lngRecs
6240            If ![chkvoid_void] = True Then
6250              .Edit
6260              ![chkvoid_void] = False
6270              ![chkvoids_datemodified] = Now()
6280              .Update
6290            End If
6300            If lngX < lngRecs Then .MoveNext
6310          Next
6320          .Close
6330        End With
6340        frm.Refresh
6350        DoEvents
6360        frm.RecalcTots  ' ** Form Procedure: frmRpt_Checks_Void_Sub.
6370        .frmRpt_Checks_Void_Sub.SetFocus
6380      Else
6390        Beep
6400      End If
6410      DoCmd.Hourglass False
6420    End With

EXITP:
6430    Set frm = Nothing
6440    Set rst = Nothing
6450    Exit Sub

ERRH:
6460    DoCmd.Hourglass False
6470    Select Case ERR.Number
        Case Else
6480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6490    End Select
6500    Resume EXITP

End Sub

Private Sub cmdSelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_KeyDown"

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
6690          If .cmdCopy.Enabled = True Then
6700            .cmdCopy.SetFocus
6710          Else
6720            .cmdOK.SetFocus
6730          End If
6740        End With
6750      End Select
6760    End If

        ' ** Shift keys.
6770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6780      Select Case intRetVal
          Case vbKeyTab
6790        With Me
6800          intRetVal = 0
6810          .cmdSelectAll.SetFocus
6820        End With
6830      End Select
6840    End If

EXITP:
6850    KeyCode = intRetVal
6860    Exit Sub

ERRH:
6870    intRetVal = 0
6880    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6890    Case Else
6900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6910    End Select
6920    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseMove"

7010    With Me
7020      If blnSelectNone_MouseDown = False Then
7030        Select Case blnSelectNone_Focus
            Case True
7040          .cmdSelectNone_raised_focus_dots_img.Visible = True
7050          .cmdSelectNone_raised_focus_img.Visible = False
7060        Case False
7070          .cmdSelectNone_raised_focus_img.Visible = True
7080          .cmdSelectNone_raised_focus_dots_img.Visible = False
7090        End Select
7100        .cmdSelectNone_raised_img.Visible = False
7110        .cmdSelectNone_raised_semifocus_dots_img.Visible = False
7120        .cmdSelectNone_sunken_focus_dots_img.Visible = False
7130        .cmdSelectNone_raised_img_dis.Visible = False
7140      End If
7150    End With

EXITP:
7160    Exit Sub

ERRH:
7170    Select Case ERR.Number
        Case Else
7180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7190    End Select
7200    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseUp"

7310    With Me
7320      .cmdSelectNone_raised_focus_dots_img.Visible = True
7330      .cmdSelectNone_raised_img.Visible = False
7340      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
7350      .cmdSelectNone_raised_focus_img.Visible = False
7360      .cmdSelectNone_sunken_focus_dots_img.Visible = False
7370      .cmdSelectNone_raised_img_dis.Visible = False
7380      blnSelectNone_MouseDown = False
7390    End With

EXITP:
7400    Exit Sub

ERRH:
7410    Select Case ERR.Number
        Case Else
7420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7430    End Select
7440    Resume EXITP

End Sub

Private Sub cmdSelectNone_LostFocus()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_LostFocus"

7510    With Me
7520      .cmdSelectNone_raised_img.Visible = True
7530      .cmdSelectNone_raised_semifocus_dots_img.Visible = False
7540      .cmdSelectNone_raised_focus_img.Visible = False
7550      .cmdSelectNone_raised_focus_dots_img.Visible = False
7560      .cmdSelectNone_sunken_focus_dots_img.Visible = False
7570      .cmdSelectNone_raised_img_dis.Visible = False
7580      blnSelectNone_Focus = False
7590    End With

EXITP:
7600    Exit Sub

ERRH:
7610    Select Case ERR.Number
        Case Else
7620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7630    End Select
7640    Resume EXITP

End Sub

Private Sub cmdCopy_GotFocus()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_GotFocus"

7710    With Me
7720      blnCopy_Focus = True
7730      .cmdCopy_raised_semifocus_dots_img.Visible = True
7740      .cmdCopy_raised_img.Visible = False
7750      .cmdCopy_raised_focus_img.Visible = False
7760      .cmdCopy_raised_focus_dots_img.Visible = False
7770      .cmdCopy_sunken_focus_dots_img.Visible = False
7780      .cmdCopy_raised_img_dis.Visible = False
7790    End With

EXITP:
7800    Exit Sub

ERRH:
7810    Select Case ERR.Number
        Case Else
7820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7830    End Select
7840    Resume EXITP

End Sub

Private Sub cmdCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_MouseDown"

7910    With Me
7920      blnCopy_MouseDown = True
7930      .cmdCopy_sunken_focus_dots_img.Visible = True
7940      .cmdCopy_raised_img.Visible = False
7950      .cmdCopy_raised_semifocus_dots_img.Visible = False
7960      .cmdCopy_raised_focus_img.Visible = False
7970      .cmdCopy_raised_focus_dots_img.Visible = False
7980      .cmdCopy_raised_img_dis.Visible = False
7990    End With

EXITP:
8000    Exit Sub

ERRH:
8010    Select Case ERR.Number
        Case Else
8020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8030    End Select
8040    Resume EXITP

End Sub

Private Sub cmdCopy_Click()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_Click"

        Dim frm As Access.Form, rst As DAO.Recordset
        Dim strBankName As String, strBankAcctNum As String
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long, lngChkVoidID As Long
        Dim lngX As Long

8110    With Me
8120      Set frm = .frmRpt_Checks_Void_Sub.Form
8130      With frm
8140        lngChkVoidID = .chkvoids_id
8150        strBankName = .chkbank_name
8160        strBankAcctNum = .chkbank_acctnum
8170        msgResponse = MsgBox("Copy bank information from this record to" & vbCrLf & _
              "other voided checks that are missing bank info." & vbCrLf & vbCrLf & _
              String(10, Chr(160)) & "Bank:  " & strBankName & vbCrLf & String(4, Chr(160)) & "Account:  " & strBankAcctNum & vbCrLf & vbCrLf & _
              "Continue?", vbQuestion + vbOKCancel + vbDefaultButton1, "Copy Bank Information")
8180        If msgResponse = vbOK Then
8190          DoCmd.Hourglass True
8200          DoEvents
8210          If IsNull(.chkbank_name) = False And IsNull(.chkbank_acctnum) = False Then
8220            Set rst = .RecordsetClone
8230            With rst
8240              .MoveLast
8250              lngRecs = .RecordCount
8260              .MoveFirst
8270              For lngX = 1& To lngRecs
8280                If ![chkvoid_void] = True And ![chkvoids_id] <> lngChkVoidID Then
8290                  .Edit
8300                  ![chkbank_name] = strBankName
8310                  ![chkbank_acctnum] = strBankAcctNum
8320                  ![chkvoids_datemodified] = Now()
8330                  .Update
8340                End If
8350                If lngX < lngRecs Then .MoveNext
8360              Next
8370              .Close
8380            End With
8390            Set rst = Nothing
8400            .Refresh
8410            Beep
8420            DoCmd.Hourglass False
8430          Else
8440            DoCmd.Hourglass False
8450            MsgBox "Bank data is missing.", vbInformation + vbOKOnly, "Nothing To Do"
8460          End If
8470        End If  ' ** msgResponse.
8480      End With  ' ** frm.
8490    End With

EXITP:
8500    Set rst = Nothing
8510    Set frm = Nothing
8520    Exit Sub

ERRH:
8530    DoCmd.Hourglass False
8540    Select Case ERR.Number
        Case Else
8550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8560    End Select
8570    Resume EXITP

End Sub

Private Sub cmdCopy_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8620    intShiftDown = (Shift And acShiftMask) > 0
8630    intAltDown = (Shift And acAltMask) > 0
8640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8660      Select Case intRetVal
          Case vbKeyTab
8670        With Me
8680          intRetVal = 0
8690          .cmdOK.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab
8750        With Me
8760          intRetVal = 0
8770          .cmdSelectAll.SetFocus
8780        End With
8790      End Select
8800    End If

EXITP:
8810    KeyCode = intRetVal
8820    Exit Sub

ERRH:
8830    intRetVal = 0
8840    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8850    Case Else
8860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8870    End Select
8880    Resume EXITP

End Sub

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_MouseMove"

8910    With Me
8920      If blnCopy_MouseDown = False Then
8930        Select Case blnCopy_Focus
            Case True
8940          .cmdCopy_raised_focus_dots_img.Visible = True
8950          .cmdCopy_raised_focus_img.Visible = False
8960        Case False
8970          .cmdCopy_raised_focus_img.Visible = True
8980          .cmdCopy_raised_focus_dots_img.Visible = False
8990        End Select
9000        .cmdCopy_raised_img.Visible = False
9010        .cmdCopy_raised_semifocus_dots_img.Visible = False
9020        .cmdCopy_sunken_focus_dots_img.Visible = False
9030        .cmdCopy_raised_img_dis.Visible = False
9040      End If
9050    End With

EXITP:
9060    Exit Sub

ERRH:
9070    Select Case ERR.Number
        Case Else
9080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9090    End Select
9100    Resume EXITP

End Sub

Private Sub cmdCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_MouseUp"

9210    With Me
9220      .cmdCopy_raised_focus_dots_img.Visible = True
9230      .cmdCopy_raised_img.Visible = False
9240      .cmdCopy_raised_semifocus_dots_img.Visible = False
9250      .cmdCopy_raised_focus_img.Visible = False
9260      .cmdCopy_sunken_focus_dots_img.Visible = False
9270      .cmdCopy_raised_img_dis.Visible = False
9280      blnCopy_MouseDown = False
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    Select Case ERR.Number
        Case Else
9320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9330    End Select
9340    Resume EXITP

End Sub

Private Sub cmdCopy_LostFocus()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopy_LostFocus"

9410    With Me
9420      .cmdCopy_raised_img.Visible = True
9430      .cmdCopy_raised_semifocus_dots_img.Visible = False
9440      .cmdCopy_raised_focus_img.Visible = False
9450      .cmdCopy_raised_focus_dots_img.Visible = False
9460      .cmdCopy_sunken_focus_dots_img.Visible = False
9470      .cmdCopy_raised_img_dis.Visible = False
9480      blnCopy_Focus = False
9490    End With

EXITP:
9500    Exit Sub

ERRH:
9510    Select Case ERR.Number
        Case Else
9520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9530    End Select
9540    Resume EXITP

End Sub
