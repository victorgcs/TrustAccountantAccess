VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile_Add"

'VGC 04/15/2017: CHANGES!

' ** For GTR, shows arrow, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Fee Annually:     A {opgFeeFreq_optAnnual on frmAccountProfile_Add_Sub}
' **   AccountNo:        B {accountno on frmAccountProfile_Add_Sub}
' **   Cancel:           C {cmdCancel}
' **   Notes:            E {tabContacts_pgNotes on frmAccountProfile_Add_Sub}
' **   Bank Other:       H {opgLoc_optOther on frmAccountProfile_Sub}
' **   Statement Freq:   K {cmdStatementFreq on frmAccountProfile_Add_Sub}
' **   Review Freq:      L {cmdReviewFreq on frmAccountProfile_Add_Sub}
' **   Fee Monthly:      M {opgFeeFreq_optMonth on frmAccountProfile_Add_Sub}
' **   Short Name:       N {shortname on frmAccountProfile_Add_Sub}
' **   Short Name PQ:    N {shortname_pq on frmAccountProfile_Add_Sub} (Printer's Quotes)
' **   OK:               O {cmdOK}
' **   Choose RA:          (cmdChoose on frmAccountProfile_Add_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Fee Quarterly:    Q {opgFeeFreq_optQuarter on frmAccountProfile_Add_Sub}
' **   Related Accounts: R {related_accountno on frmAccountProfile_Add_Sub} (Disabled)
' **   Fee SemiAnnually: S {opgFeeFreq_optSemi on frmAccountProfile_Add_Sub}
' **   Account Type:     T {accounttype on frmAccountProfile_Add_Sub}
' **   Bank USA:         U {opgLoc_optUSA}
' **   Contact Name 1:   1 {tabContacts_pgContact1 on frmAccountProfile_Add_Sub}
' **   Contact Name 2:   2 {tabContacts_pgContact2 on frmAccountProfile_Add_Sub}

' ** Shortcut Ctrl keys responsive from this form:
' **   Contact1 Other:   H {opgLocContact1_optOther on frmAccountProfile_Sub}
' **   Save:             S {cmdSave on frmAccountProfile_Add_Sub}
' **   Contact1 USA:     U {opgLocContact1_optUSA on frmAccountProfile_Sub}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Contact2 Other:   H {opgLocContact2_optOther on frmAccountProfile_Sub}
' **   Contact2 USA:     U {opgLocContact2_optUSA on frmAccountProfile_Sub}

Private lngTpp As Long, lngSizable_Offset As Long
Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngSub_Width As Long, lngSub_Height As Long, lngSubBox_Width As Long, lngSubBox_Height As Long
Private lngCancel_Left As Long, lngOK_Offset As Long, lngDetail_Height As Long, lngSubHLine_Width As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim usr As DAO.User, grp As DAO.Group
        Dim blnRetVal As Boolean, blnFound As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.
160           StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

170           gstrFormQuerySpec = THIS_NAME

180           gstrFormQuerySpec = THIS_NAME
190           gblnClosing = False
200           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

              ' ** Check user status, and disable add/delete if view only.
210           blnFound = False
220           For Each grp In DBEngine.Workspaces(0).Groups
230             If grp.Name = "ViewOnly" Then
240               For Each usr In grp.Users
250                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
260                   blnFound = True
270                   Exit For
280                 End If
290               Next
300             End If
310           Next

320           .cmdPrintReport.Enabled = False
330           .cmdPrintReport_raised_img_dis.Visible = True
340           .cmdPrintReport_raised_img.Visible = False
350           .cmdPrintReport_raised_semifocus_dots_img.Visible = False
360           .cmdPrintReport_raised_focus_img.Visible = False
370           .cmdPrintReport_raised_focus_dots_img.Visible = False
380           .cmdPrintReport_sunken_focus_dots_img.Visible = False

              ' ** Variables are fed empty, then populated ByRef.
390           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
400           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
410           lngCancel_Left = .cmdCancel.Left
420           lngOK_Offset = (.cmdCancel.Left - .cmdOK.Left)
430           lngSub_Width = .frmAccountProfile_Add_Sub.Width
440           lngSub_Height = .frmAccountProfile_Add_Sub.Height
450           lngSubBox_Width = .frmAccountProfile_Add_Sub_box.Width
460           lngSubBox_Height = .frmAccountProfile_Add_Sub_box.Height
470           lngDetail_Height = .frmAccountProfile_Add_Sub.Form.Detail.Height
480           lngSubHLine_Width = .frmAccountProfile_Add_Sub.Form.Footer_hline01.Width
490           lngSizable_Offset = (.Sizable_lbl1.Left - .cmdCancel.Left)

500           .cmdOK.Enabled = False

510           DoCmd.Hourglass False

520         Else
530           Cancel = -1
540         End If
550       Else
560         Cancel = -1
570       End If
580     End With

590     If Cancel = -1 Then
600       Beep
610       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
620     End If

EXITP:
630     Set usr = Nothing
640     Set grp = Nothing
650     Exit Sub

ERRH:
660     DoCmd.Hourglass False
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality.
        ' **   Notes:            E {tabContacts_pgNotes on frmAccountProfile_Sub}
        ' **   Short Name:       N {shortname on frmAccountProfile_Sub}
        ' **   Short Name PQ:    N {shortname_pq on frmAccountProfile_Sub}
        ' **   Contact Name 1:   1 {tabContacts_pgContact1 on frmAccountProfile_Sub}
        ' **   Contact Name 2:   2 {tabContacts_pgContact2 on frmAccountProfile_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality.
        ' **   Contact1 Other:   H {opgLocContact1_optOther on frmAccountProfile_Sub}
        ' **   Save:             S {cmdSave}
        ' **   Contact1 USA:     U {opgLocContact1_optUSA on frmAccountProfile_Sub}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality.
        ' **   Contact2 Other:   H {opgLocContact2_optOther on frmAccountProfile_Sub}
        ' **   Contact2 USA:     U {opgLocContact2_optUSA on frmAccountProfile_Sub}

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         With Me
880           intRetVal = 0
890           gblnClosing = True
900           cmdCancel_Click  ' ** Procedure: Below.
910         End With
920       End Select
930     End If

        ' ** Alt keys.
940     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
950       Select Case intRetVal
          Case vbKeyE
960         With Me
970           intRetVal = 0
980           .frmAccountProfile_Add_Sub.SetFocus
990           .frmAccountProfile_Add_Sub.Form.FocusHolder.SetFocus
1000          .frmAccountProfile_Add_Sub.Form.tabContacts = .frmAccountProfile_Add_Sub.Form.tabContacts_pgNotes.PageIndex
1010          .frmAccountProfile_Add_Sub.Form.Notes.SetFocus
1020          fSetScrollBarPosVT .frmAccountProfile_Add_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
1030        End With
1040      Case vbKeyN
1050        With Me
1060          intRetVal = 0
1070          .frmAccountProfile_Add_Sub.SetFocus
1080          If .frmAccountProfile_Add_Sub.Form.shortname.Visible = True And .frmAccountProfile_Add_Sub.Form.shortname.Enabled = True Then
1090            .frmAccountProfile_Add_Sub.Form.shortname.SetFocus
1100          ElseIf .frmAccountProfile_Add_Sub.Form.shortname_pq.Visible = True And .frmAccountProfile_Add_Sub.Form.shortname_pq.Enabled = True Then
1110            .frmAccountProfile_Add_Sub.Form.shortname_pq.SetFocus
1120          End If
1130        End With
1140      Case vbKey1
1150        With Me
1160          intRetVal = 0
1170          .frmAccountProfile_Add_Sub.SetFocus
1180          .frmAccountProfile_Add_Sub.Form.FocusHolder.SetFocus
1190          .frmAccountProfile_Add_Sub.Form.tabContacts = .frmAccountProfile_Add_Sub.Form.tabContacts_pgContact1.PageIndex
1200          .frmAccountProfile_Add_Sub.Form.Contact1_Name.SetFocus
1210          fSetScrollBarPosVT .frmAccountProfile_Add_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
1220        End With
1230      Case vbKey2
1240        With Me
1250          intRetVal = 0
1260          .frmAccountProfile_Add_Sub.SetFocus
1270          .frmAccountProfile_Add_Sub.Form.FocusHolder.SetFocus
1280          .frmAccountProfile_Add_Sub.Form.tabContacts = .frmAccountProfile_Add_Sub.Form.tabContacts_pgContact2.PageIndex
1290          .frmAccountProfile_Add_Sub.Form.Contact2_Name.SetFocus
1300          fSetScrollBarPosVT .frmAccountProfile_Add_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
1310        End With
1320      End Select
1330    End If

        ' ** Ctrl keys.
1340    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1350      Select Case intRetVal
          Case vbKeyH
1360        With Me
1370          If .frmAccountProfile_Add_Sub.Form.opgLocContact1 <> .frmAccountProfile_Add_Sub.Form.opgLocContact1_optOther.OptionValue Then
1380            .frmAccountProfile_Add_Sub.SetFocus
1390            .frmAccountProfile_Add_Sub.Form.opgLocContact1 = .frmAccountProfile_Add_Sub.Form.opgLocContact1_optOther.OptionValue
1400            .frmAccountProfile_Add_Sub.Form.opgLocContact1_AfterUpdate  ' ** Form Procedure: frmAccountProfile_Add_Sub.
1410          End If
1420        End With
1430      Case vbKeyS
1440        With Me
1450          intRetVal = 0
1460          .frmAccountProfile_Add_Sub.SetFocus
1470          .frmAccountProfile_Add_Sub.Form.cmdSave_Click ' ** Form Procedure: frmAccountProfile_Sub.
1480        End With
1490      Case vbKeyU
1500        With Me
1510          If .frmAccountProfile_Add_Sub.Form.opgLocContact1 <> .frmAccountProfile_Add_Sub.Form.opgLocContact1_optUSA.OptionValue Then
1520            .frmAccountProfile_Add_Sub.SetFocus
1530            .frmAccountProfile_Add_Sub.Form.opgLocContact1 = .frmAccountProfile_Add_Sub.Form.opgLocContact1_optUSA.OptionValue
1540            .frmAccountProfile_Add_Sub.Form.opgLocContact1_AfterUpdate  ' ** Form Procedure: frmAccountProfile_Add_Sub.
1550          End If
1560        End With
1570      End Select
1580    End If

        ' ** Ctrl-Shift keys.
1590    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1600      Select Case intRetVal
          Case vbKeyH
1610        With Me
1620          If .frmAccountProfile_Add_Sub.Form.opgLocContact2 <> .frmAccountProfile_Add_Sub.Form.opgLocContact2_optOther.OptionValue Then
1630            .frmAccountProfile_Add_Sub.SetFocus
1640            .frmAccountProfile_Add_Sub.Form.opgLocContact2 = .frmAccountProfile_Add_Sub.Form.opgLocContact2_optOther.OptionValue
1650            .frmAccountProfile_Add_Sub.Form.opgLocContact2_AfterUpdate  ' ** Form Procedure: frmAccountProfile_Add_Sub.
1660          End If
1670        End With
1680      Case vbKeyU
1690        With Me
1700          If .frmAccountProfile_Add_Sub.Form.opgLocContact2 <> .frmAccountProfile_Add_Sub.Form.opgLocContact2_optUSA.OptionValue Then
1710            .frmAccountProfile_Add_Sub.SetFocus
1720            .frmAccountProfile_Add_Sub.Form.opgLocContact2 = .frmAccountProfile_Add_Sub.Form.opgLocContact2_optUSA.OptionValue
1730            .frmAccountProfile_Add_Sub.Form.opgLocContact2_AfterUpdate  ' ** Form Procedure: frmAccountProfile_Add_Sub.
1740          End If
1750        End With
1760      End Select
1770    End If

EXITP:
1780    KeyCode = intRetVal
1790    Exit Sub

ERRH:
1800    intRetVal = 0
1810    Select Case ERR.Number
        Case Else
1820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1830    End Select
1840    Resume EXITP

End Sub

Private Sub Form_Resize()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngSubHeight As Long
        Dim lngTmp01 As Long

1910    With Me

1920      GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.

1930      If lngWidth <> lngFrm_Width Or lngHeight <> lngFrm_Height Then

1940        If lngWidth <= lngFrm_Width Then
              ' ** Don't let it shrink narrower than it's opening width.
1950          .cmdCancel.Left = lngCancel_Left - (lngFrm_Width - lngWidth)
1960          .cmdOK.Left = .cmdCancel.Left - lngOK_Offset
1970          .Sizable_lbl1.Left = (.cmdCancel.Left + lngSizable_Offset)
1980          .Sizable_lbl2.Left = (.cmdCancel.Left + lngSizable_Offset)
1990        Else
2000          .cmdCancel.Left = lngCancel_Left + (lngWidth - lngFrm_Width)
2010          .cmdOK.Left = .cmdCancel.Left - lngOK_Offset
2020          .Sizable_lbl1.Left = (.cmdCancel.Left + lngSizable_Offset)
2030          .Sizable_lbl2.Left = (.cmdCancel.Left + lngSizable_Offset)
2040          .frmAccountProfile_Add_Sub.Width = lngSub_Width + (lngWidth - lngFrm_Width)
2050          .frmAccountProfile_Add_Sub_box.Width = (lngSub_Width + (lngWidth - lngFrm_Width)) + (2& * lngTpp)
2060          lngTmp01 = ((.cmdCancel.Left + .cmdCancel.Width) + (9& * lngTpp))
2070          .Header_hline01.Width = lngTmp01
2080          .Header_hline02.Width = lngTmp01
2090          .Header_vline01.Left = lngTmp01
2100          .Header_vline02.Left = lngTmp01
2110          .Footer_hline01.Width = lngTmp01
2120          .Footer_hline02.Width = lngTmp01
2130          .Footer_vline01.Left = lngTmp01
2140          .Footer_vline02.Left = lngTmp01
2150          lngTmp01 = lngSubHLine_Width + (lngWidth - lngFrm_Width)
2160          .frmAccountProfile_Add_Sub.Form.Footer_hline01.Width = lngTmp01
2170          .frmAccountProfile_Add_Sub.Form.Footer_hline02.Width = lngTmp01
2180          .frmAccountProfile_Add_Sub.Form.Footer_vline01.Left = lngTmp01
2190          .frmAccountProfile_Add_Sub.Form.Footer_vline02.Left = lngTmp01
2200        End If

2210        lngSubHeight = .frmAccountProfile_Add_Sub.Height
2220        If lngHeight > lngFrm_Height Then
2230          .Detail.Height = lngDetail_Height + (lngHeight - lngFrm_Height)
2240          .frmAccountProfile_Add_Sub.Height = lngSub_Height + (lngHeight - lngFrm_Height)
2250          .frmAccountProfile_Add_Sub_box.Height = (lngSub_Height + (lngHeight - lngFrm_Height)) + (2& * lngTpp)
2260          .Detail.Height = (.frmAccountProfile_Add_Sub.Top + .frmAccountProfile_Add_Sub.Height) + (32& * lngTpp)
2270          .detail_height_line.Top = (.Detail.Height - (24& * lngTpp))
2280        Else
2290          If (lngSub_Height - (lngFrm_Height - lngHeight)) > (.frmAccountProfile_Add_Sub.Form.Detail.Height + (19& * lngTpp)) Then
2300            .frmAccountProfile_Add_Sub.Height = (lngSub_Height - (lngFrm_Height - lngHeight))
2310            .frmAccountProfile_Add_Sub_box.Height = (lngSub_Height - (lngFrm_Height - lngHeight)) + (2& * lngTpp)
2320            .detail_height_line.Top = ((lngDetail_Height - (lngFrm_Height - lngHeight)) - (24& * lngTpp))
2330            .Detail.Height = (.frmAccountProfile_Add_Sub.Top + .frmAccountProfile_Add_Sub.Height) + (32& * lngTpp)
2340            .detail_height_line.Top = (.Detail.Height - (24& * lngTpp))
2350          End If
2360        End If

2370      End If

2380    End With

EXITP:
2390    Exit Sub

ERRH:
2400    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
2410    Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub Form_Timer()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2510    With Me
2520      .TimerInterval = 0&
2530      If gblnSetFocus = True Then
2540        gblnSetFocus = False
2550        strFormName = vbNullString
2560  On Error Resume Next
2570        Do Until strFormName = THIS_NAME
2580          DoCmd.SelectObject acForm, THIS_NAME, False
2590          strFormName = Screen.ActiveForm.Name
2600        Loop
2610  On Error GoTo ERRH
2620        .frmAccountProfile_Add_Sub.SetFocus
2630        DoCmd.Hourglass False
2640      End If
2650      If gblnGoToReport = True Then
2660        Select Case blnGoingToReport
            Case True
              ' ** Just sit there, and we're done.
2670          DoCmd.Hourglass False  ' ** It should already be off.
2680          gblnGoToReport = False
2690          blnGoingToReport = False
2700          .GoToReport_arw_pr_img.Visible = False
2710          .cmdCancel.SetFocus
2720          DoEvents
2730          .cmdPrintReport.Enabled = False
2740          .cmdPrintReport_raised_img_dis.Visible = True
2750          .cmdPrintReport_raised_img.Visible = False
2760          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2770          .cmdPrintReport_raised_focus_img.Visible = False
2780          .cmdPrintReport_raised_focus_dots_img.Visible = False
2790          .cmdPrintReport_sunken_focus_dots_img.Visible = False
2800          DoEvents
2810        Case False
2820          blnGoingToReport = True
2830          .GoToReport_arw_pr_img.Visible = True
2840          .cmdPrintReport.Enabled = True
2850          .cmdPrintReport.SetFocus
2860          DoEvents
2870          Beep
2880          DoCmd.Hourglass False
2890          .TimerInterval = (GTR_WAIT * 2&)
2900        End Select
2910      End If
2920    End With

EXITP:
2930    Exit Sub

ERRH:
2940    DoCmd.Hourglass False
2950    Select Case ERR.Number
        Case Else
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3010    With Me

3020      DoCmd.Hourglass True
3030      DoEvents

3040      Pref_Save THIS_NAME & "_Sub"  ' ** Module Procedure: modPreferenceFuncs.

          ' ** Make sure all reports are closed.
3050      If Reports.Count > 0 Then
3060        Do While Reports.Count > 0
3070          DoCmd.Close acReport, Reports(0).Name
3080        Loop
3090      End If

3100      DoCmd.Hourglass False
3110      gstrFormQuerySpec = vbNullString
3120      gblnSetFocus = True
3130      gstrReturningForm = THIS_NAME
3140      Select Case strCallingForm
          Case "frmMenu_Account"
3150        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
3160      End Select

3170    End With

EXITP:
3180    Exit Sub

ERRH:
3190    DoCmd.Hourglass False
3200    Select Case ERR.Number
        Case Else
3210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3220    End Select
3230    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3310    With Me
3320      If gblnGoToReport = False Then
3330        If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3340          Select Case blnPrintReport_Focus
              Case True
3350            .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3360            .cmdPrintReport_raised_img.Visible = False
3370          Case False
3380            .cmdPrintReport_raised_img.Visible = True
3390            .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3400          End Select
3410          .cmdPrintReport_raised_focus_img.Visible = False
3420          .cmdPrintReport_raised_focus_dots_img.Visible = False
3430          .cmdPrintReport_sunken_focus_dots_img.Visible = False
3440          .cmdPrintReport_raised_img_dis.Visible = False
3450        End If
3460      End If
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

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

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

Private Sub cmdOK_Click()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, frm As Access.Form
        Dim strControl As String
        Dim blnContinue As Boolean, blnSuppress As Boolean

3710    blnContinue = True

3720    With Me

3730      DoCmd.Hourglass True
3740      DoEvents

3750      strControl = vbNullString

3760      Set frm = .frmAccountProfile_Add_Sub.Form
3770      With frm

3780        If .Dirty = True Then
3790          .cmdSave_Click  ' ** Form Procedure: frmAccountProfile_Add_Sub.
3800        End If

3810        If IsNull(.accountno) = True Then
3820          blnContinue = False
3830          DoCmd.Hourglass False
3840          MsgBox "Please enter a valid Account Number.", vbInformation + vbOKOnly, "Entry Required"
3850          strControl = "accountno"
3860        Else
3870          If Trim(.accountno) = vbNullString Then
3880            blnContinue = False
3890            DoCmd.Hourglass False
3900            MsgBox "Please enter a valid Account Number.", vbInformation + vbOKOnly, "Entry Required"
3910            strControl = "accountno"
3920          Else
3930            If IsNull(.shortname) = True Then
3940              blnContinue = False
3950              DoCmd.Hourglass False
3960              MsgBox "Please enter a Short Name to continue", vbInformation + vbOKOnly, "Entry Required"
3970              strControl = "shortname"
3980            Else
3990              If Trim(.shortname) = vbNullString Then
4000                blnContinue = False
4010                DoCmd.Hourglass False
4020                MsgBox "Please enter a Short Name to continue", vbInformation + vbOKOnly, "Entry Required"
4030                strControl = "shortname"
4040              Else
4050                If IsNull(.accounttype) = True Then
4060                  blnContinue = False
4070                  DoCmd.Hourglass False
4080                  MsgBox "Please enter an Account Type to continue", vbInformation + vbOKOnly, "Entry Required"
4090                  strControl = "accounttype"
4100                Else
                      ' ** Proceed to the next step.
4110                End If
4120              End If
4130            End If
4140          End If
4150        End If

4160        If blnContinue = True Then

              ' ** For now, I'm using an unused field for the default asset.
4170          If IsNull(.taxlot) = True Then
4180            .taxlot = "0"
4190          Else
4200            If IsNumeric(.taxlot) = False Then
4210              .taxlot = "0"
4220            End If
4230          End If

4240          gstrActNo = .accountno

4250          Set dbs = CurrentDb
4260          With dbs

                ' ** Append tmpAccount to Account, by specified [actno].
4270            Set qdf = .QueryDefs("qryAccountProfile_16")
4280            With qdf.Parameters
4290              ![actno] = gstrActNo
4300            End With
4310            qdf.Execute dbFailOnError
4320            Set qdf = Nothing

                ' ** Append new account to Balance, by specified [actno].
4330            Set qdf = .QueryDefs("qryAccountProfile_17")
4340            With qdf.Parameters
4350              ![actno] = gstrActNo
4360            End With
4370            qdf.Execute dbFailOnError
4380            Set qdf = Nothing
4390            DoEvents

4400            blnSuppress = Pref_Suppress  ' ** Module Function: modPreferenceFuncs.

                ' ** Append qryAccountProfile_Add_02 (New tblCurrency_Account
                ' ** record, by specified [actno], [supr]) to tblCurrency_Account.
4410            Set qdf = .QueryDefs("qryAccountProfile_Add_03")
4420            With qdf.Parameters
4430              ![actno] = gstrActNo
4440              ![supr] = blnSuppress
4450            End With
4460            qdf.Execute dbFailOnError
4470            Set qdf = Nothing
4480            DoEvents

                ' ** Append qryAccountProfile_01c (qryAccountProfile_01b (Account, just accountno), with
                ' ** qryAccountProfile_01a (tblForm_Graphics, just frmAccountProfile_Sub); Cartesian) to tmpForm_Graphics_Account.
4490            Set qdf = .QueryDefs("qryAccountProfile_01d")
4500            qdf.Execute
4510            Set qdf = Nothing

4520            .Close
4530          End With
4540          Set dbs = Nothing
4550          DoEvents

              ' ** Reload currency stuff.
4560          gblnHasForEx = HasForEx_All  ' ** Module Function: modCurrencyFuncs.
4570          HasForEx_Load  ' ** Module Procedure: modCurrencyFuncs.
4580          DoEvents

4590        End If  ' ** blnContinue.

4600        If blnContinue = False And strControl <> vbNullString Then
4610          .Controls(strControl).SetFocus
4620        End If

4630      End With

4640      If blnContinue = True Then
4650        DoCmd.Close acForm, THIS_NAME
4660      End If

4670    End With

EXITP:
4680    Set frm = Nothing
4690    Set qdf = Nothing
4700    Set dbs = Nothing
4710    Exit Sub

ERRH:
4720    Select Case ERR.Number
        Case Else
4730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4740    End Select
4750    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
4890          .cmdCancel.SetFocus
4900        End With
4910      End Select
4920    End If

        ' ** Shift keys.
4930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4940      Select Case intRetVal
          Case vbKeyTab
4950        With Me
4960          intRetVal = 0
4970          If .cmdPrintReport.Enabled = True Then
4980            .cmdPrintReport.SetFocus
4990          Else
5000            .frmAccountProfile_Add_Sub.SetFocus
5010            .frmAccountProfile_Add_Sub.Form.FocusHolder.SetFocus
5020            .frmAccountProfile_Add_Sub.Form.tabContacts = .frmAccountProfile_Add_Sub.Form.tabContacts_pgContact1.PageIndex
5030            .frmAccountProfile_Add_Sub.Form.Contact1_Email.SetFocus
5040            fSetScrollBarPosVT .frmAccountProfile_Add_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
5050          End If
5060        End With
5070      End Select
5080    End If

EXITP:
5090    KeyCode = intRetVal
5100    Exit Sub

ERRH:
5110    intRetVal = 0
5120    Select Case ERR.Number
        Case Else
5130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5140    End Select
5150    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

5210    With Me

5220      .TimerInterval = 0&

5230      DoCmd.Close acForm, THIS_NAME

5240    End With

EXITP:
5250    Exit Sub

ERRH:
5260    Select Case ERR.Number
        Case Else
5270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5280    End Select
5290    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
5370        With Me
5380          intRetVal = 0
5390          .frmAccountProfile_Add_Sub.SetFocus
5400          .frmAccountProfile_Add_Sub.Form.accountno.SetFocus
5410          fSetScrollBarPosVT .frmAccountProfile_Add_Sub.Form, 1&  ' ** Module Function: modScrollBarFuncs.
5420        End With
5430      End Select
5440    End If

        ' ** Shift keys.
5450    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5460      Select Case intRetVal
          Case vbKeyTab
5470        With Me
5480          intRetVal = 0
5490          If .cmdOK.Enabled = True Then
5500            .cmdOK.SetFocus
5510          ElseIf .cmdPrintReport.Enabled = True Then
5520            .cmdPrintReport.SetFocus
5530          Else
5540            .frmAccountProfile_Add_Sub.SetFocus
5550            .frmAccountProfile_Add_Sub.Form.FocusHolder.SetFocus
5560            .frmAccountProfile_Add_Sub.Form.tabContacts = .frmAccountProfile_Add_Sub.Form.tabContacts_pgContact1.PageIndex
5570            .frmAccountProfile_Add_Sub.Form.Contact1_Email.SetFocus
5580            fSetScrollBarPosVT .frmAccountProfile_Add_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
5590          End If
5600        End With
5610      End Select
5620    End If

EXITP:
5630    KeyCode = intRetVal
5640    Exit Sub

ERRH:
5650    intRetVal = 0
5660    Select Case ERR.Number
        Case Else
5670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5680    End Select
5690    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

5710    With Me
5720      blnPrintReport_Focus = True
5730      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
5740      .cmdPrintReport_raised_img.Visible = False
5750      .cmdPrintReport_raised_focus_img.Visible = False
5760      .cmdPrintReport_raised_focus_dots_img.Visible = False
5770      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5780      .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

5910    With Me
5920      blnPrintReport_MouseDown = True
5930      .cmdPrintReport_sunken_focus_dots_img.Visible = True
5940      .cmdPrintReport_raised_img.Visible = False
5950      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5960      .cmdPrintReport_raised_focus_img.Visible = False
5970      .cmdPrintReport_raised_focus_dots_img.Visible = False
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

Private Sub cmdPrintReport_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String
        Dim intMode As Integer

6110    With Me
6120      gstrFormQuerySpec = THIS_NAME
6130      Select Case gblnSeparateCheckingAccounts
          Case True
6140        strDocName = "rptAccountProfile_SepChkAcct"
6150      Case False
6160        strDocName = "rptAccountProfile_SngChkAcct"
6170      End Select
6180      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
6190        intMode = acViewPreview
6200      Else
6210        intMode = acViewNormal
6220      End If
6230      gstrAccountNo = .frmAccountProfile_Add_Sub.Form.accountno
          '##GTR_Ref: rptAccountProfile_SepChkAcct
          '##GTR_Ref: rptAccountProfile_SngChkAcct
6240      DoCmd.OpenReport strDocName, intMode
6250      If intMode = acViewPreview Then
6260        DoCmd.Maximize
6270        DoCmd.RunCommand acCmdFitToWindow
6280      End If
6290    End With

EXITP:
6300    Exit Sub

ERRH:
6310    Select Case ERR.Number
        Case Else
6320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6330    End Select
6340    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
6490          .cmdOK.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab
6550        With Me
6560          intRetVal = 0
6570          .frmAccountProfile_Add_Sub.SetFocus
6580        End With
6590      End Select
6600    End If

EXITP:
6610    KeyCode = intRetVal
6620    Exit Sub

ERRH:
6630    intRetVal = 0
6640    Select Case ERR.Number
        Case Else
6650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6660    End Select
6670    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

6710    With Me
6720      If blnPrintReport_MouseDown = False Then
6730        Select Case blnPrintReport_Focus
            Case True
6740          .cmdPrintReport_raised_focus_dots_img.Visible = True
6750          .cmdPrintReport_raised_focus_img.Visible = False
6760        Case False
6770          .cmdPrintReport_raised_focus_img.Visible = True
6780          .cmdPrintReport_raised_focus_dots_img.Visible = False
6790        End Select
6800        .cmdPrintReport_raised_img.Visible = False
6810        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
6830        .cmdPrintReport_raised_img_dis.Visible = False
6840      End If
6850    End With

EXITP:
6860    Exit Sub

ERRH:
6870    Select Case ERR.Number
        Case Else
6880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6890    End Select
6900    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

7010    With Me
7020      .cmdPrintReport_raised_focus_dots_img.Visible = True
7030      .cmdPrintReport_raised_img.Visible = False
7040      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7050      .cmdPrintReport_raised_focus_img.Visible = False
7060      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7070      .cmdPrintReport_raised_img_dis.Visible = False
7080      blnPrintReport_MouseDown = False
7090    End With

EXITP:
7100    Exit Sub

ERRH:
7110    Select Case ERR.Number
        Case Else
7120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7130    End Select
7140    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

7210    With Me
7220      .cmdPrintReport_raised_img.Visible = True
7230      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7240      .cmdPrintReport_raised_focus_img.Visible = False
7250      .cmdPrintReport_raised_focus_dots_img.Visible = False
7260      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7270      .cmdPrintReport_raised_img_dis.Visible = False
7280      blnPrintReport_Focus = False
7290    End With

EXITP:
7300    Exit Sub

ERRH:
7310    Select Case ERR.Number
        Case Else
7320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7330    End Select
7340    Resume EXITP

End Sub

Private Sub frmAccountProfile_Add_Sub_Enter()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "frmAccountProfile_Add_Sub_Enter"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

7410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7420    Exit Sub

ERRH:
7430    Select Case ERR.Number
        Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub

Private Sub frmAccountProfile_Add_Sub_Exit(Cancel As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "frmAccountProfile_Add_Sub_Exit"

7510    With Me
7520      If .cmdOK.Enabled = False Then
            ' ** They may be trying to close.
7530        If IsNull(.frmAccountProfile_Add_Sub.Form.accountno) = True Or _
                IsNull(.frmAccountProfile_Add_Sub.Form.shortname) = True Then
7540          gblnClosing = True
7550        End If
7560      End If
7570    End With

EXITP:
7580    Exit Sub

ERRH:
7590    Select Case ERR.Number
        Case Else
7600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7610    End Select
7620    Resume EXITP

End Sub
