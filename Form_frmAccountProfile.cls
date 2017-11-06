VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile"

'VGC 10/09/2017: CHANGES!

' ** For GTR, shows arrow, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Fee Annually:     A {opgFeeFreq_optAnnual on frmAccountProfile_Sub}
' **   AccountNo:        B {accountno on frmAccountProfile_Sub}
' **   Close:            C {cmdClose}
' **   Notes:            E {tabContacts_pgNotes on frmAccountProfile_Sub}
' **   Bank Other:       H {opgLoc_optOther on frmAccountProfile_Sub}
' **   Statement Freq:   K {cmdStatementFreq on frmAccountProfile_Sub}
' **   Review Freq:      L {cmdReviewFreq on frmAccountProfile_Sub}
' **   Fee Monthly:      M {opgFeeFreq_optMonth on frmAccountProfile_Sub}
' **   Short Name:       N {shortname on frmAccountProfile_Sub}
' **   Short Name PQ:    N {shortname_pq on frmAccountProfile_Sub} (Printer's Quotes)
' **   Choose RA:        O {cmdChoose on frmAccountProfile_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Fee Quarterly:    Q {opgFeeFreq_optQuarter on frmAccountProfile_Sub}
' **   Related Accounts: R {related_accountno on frmAccountProfile_Sub}
' **   Fee SemiAnnually: S {opgFeeFreq_optSemi on frmAccountProfile_Sub}
' **   Account Type:     T {accounttype on frmAccountProfile_Sub}
' **   Bank USA:         U {opgLoc_optUSA on frmAccountProfile_Sub}
' **   Exit:             X {cmdClose}
' **   Contact Name 1:   1 {tabContacts_pgContact1 on frmAccountProfile_Sub}
' **   Contact Name 2:   2 {tabContacts_pgContact2 on frmAccountProfile_Sub}

' ** Shortcut Ctrl keys responsive from this form:
' **   Contact1 Other:   H {opgLocContact1_optOther on frmAccountProfile_Sub}
' **   Next Record:      N {MoveRec on frmAccountProfile_Sub}
' **   Previous Record:  P {MoveRec on frmAccountProfile_Sub}
' **   Save:             S {cmdSave on frmAccountProfile_Sub}
' **   Contact1 USA:     U {opgLocContact1_optUSA on frmAccountProfile_Sub}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Contact2 Other:   H {opgLocContact2_optOther on frmAccountProfile_Sub}
' **   Close Timer:      T {Form_KeyDown}
' **   Contact2 USA:     U {opgLocContact2_optUSA on frmAccountProfile_Sub}

Private lngTpp As Long, lngNavOffset_Top As Long, lngShortMenuOffset_Top As Long, lngSizable_Offset As Long
Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngSub_Width As Long, lngSub_Height As Long, lngSubBox_Width As Long, lngSubBox_Height As Long
Private lngClose_Left As Long, lngDetail_Height As Long, lngSubHLine_Width As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnSysAcct As Boolean, blnIncludeClosed As Boolean, blnRelAcctsCalled As Boolean
Private strThisAccountNo As String, lngMaxAccounts As Long
Private blnExit As Boolean, blnEscPressed As Boolean
Private blnAccountHasChanged As Boolean, blnNewHasSaved As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim usr As DAO.User, grp As DAO.Group, ctl As Access.Control
        Dim blnRetVal As Boolean, blnFound As Boolean

110     blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.

120     With Me
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then

150           strCallingForm = Trim(.OpenArgs)
160           blnIncludeClosed = CBool(Mid(strCallingForm, (InStr(strCallingForm, "~") + 1)))
170           strCallingForm = Left(strCallingForm, (InStr(strCallingForm, "~") - 1))
180           .chkIncludeClosed = blnIncludeClosed

190           gstrFormQuerySpec = THIS_NAME

200           blnEscPressed = False: blnRelAcctsCalled = False
210           blnAccountHasChanged = False: blnNewHasSaved = False
220           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
230           blnGTR_Emblem = False: lngGTR_Stat = 0&

240           If gblnGoToReport = True Then
250             Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                Case True
260               .GoToReport_arw_pr_img.Visible = True
270             Case False
280               .GoToReport_Emblem_01_img.Visible = True
290             End Select
300             DoEvents
310           End If

320           lngMaxAccounts = 0&
330           ChkAccountLimit  ' ** Procedure: Below.
340           StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

              ' ** Check user status, and lock all fields if ViewOnly.
350           blnFound = False
360           For Each grp In DBEngine.Workspaces(0).Groups
370             If grp.Name = "ViewOnly" Then
380               For Each usr In grp.Users
390                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
400                   blnFound = True
410                   Exit For
420                 End If
430               Next
440             End If
450           Next

460           If blnFound = True Then
470             .frmAccountProfile_Sub.Form.ViewOnly_Set True  ' ** Form Procedure: frmAccountProfile_Sub.
480             With .frmAccountProfile_Sub.Form
490               For Each ctl In .Controls
500                 With ctl
510                   If .Visible = True Then
520                     Select Case .ControlType
                        Case acTextBox, acComboBox
530                       If .Enabled = True And .Locked = False Then
540                         .Locked = True
550                       End If
560                     Case acOptionButton
570                       If .Enabled = True And .Locked = False Then
580                         .Locked = True
590                       End If
600                     Case acCommandButton
610                       If .Enabled = True Then
620                         If .Name = "cmdChoose" Then
630                           .Enabled = False
640                           Me.frmAccountProfile_Sub.Form.cmdChoose_lbl2.ForeColor = WIN_CLR_DISF
650                           Me.frmAccountProfile_Sub.Form.cmdChoose_lbl2_dim_hi.Visible = True
660                         Else
                              ' ** Lock the Freq. check boxes.
670                         End If
680                       End If
690                     End Select
700                   End If
710                 End With
720               Next
730             End With
740           End If

              ' ** Variables are fed empty, then populated ByRef.
750           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
760           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
770           lngClose_Left = .cmdClose.Left
780           lngSub_Width = .frmAccountProfile_Sub.Width
790           lngSub_Height = .frmAccountProfile_Sub.Height
800           lngSubBox_Width = .frmAccountProfile_Sub_box.Width
810           lngSubBox_Height = .frmAccountProfile_Sub_box.Height
820           lngDetail_Height = .frmAccountProfile_Sub.Form.Detail.Height
830           lngSubHLine_Width = .frmAccountProfile_Sub.Form.Footer_hline01.Width
840           lngSizable_Offset = .Sizable_lbl1.Left - .cmdClose.Left
850           lngNavOffset_Top = (.frmAccountProfile_Sub_box.Top + .frmAccountProfile_Sub_box.Height) - .Nav_hline01.Top
860           lngShortMenuOffset_Top = (.ShortcutMenu_lbl.Top - (.frmAccountProfile_Sub_box.Top + .frmAccountProfile_Sub_box.Height))

870         Else
880           Cancel = -1
890         End If
900       Else
910         Cancel = -1
920       End If
930     End With

940     DoCmd.Hourglass False  ' ** This is turned off here, and in Form_Load()! Why is it still on?

950     If Cancel = -1 Then
960       Beep
970       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
980     End If

EXITP:
990     Set ctl = Nothing
1000    Set usr = Nothing
1010    Set grp = Nothing
1020    Exit Sub

ERRH:
1030    DoCmd.Hourglass False
1040    Select Case ERR.Number
        Case Else
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End Select
1070    Resume EXITP

End Sub

Private Sub Form_Load()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim strFilter As String, strSortNow As String

1110    With Me
1120      Select Case .chkIncludeClosed
          Case True
1130        strFilter = vbNullString
1140        .frmAccountProfile_Sub.Form.Filter = strFilter
1150        .frmAccountProfile_Sub.Form.FilterOn = False
1160      Case False
1170        strFilter = "[dateclosed] Is Null"
1180        .frmAccountProfile_Sub.Form.Filter = strFilter
1190        .frmAccountProfile_Sub.Form.FilterOn = True
1200      End Select
1210      strSortNow = "[alphasort]"  ' ** This comes from tmpAccount.
1220      .frmAccountProfile_Sub.Form.OrderBy = strSortNow
1230      .frmAccountProfile_Sub.Form.OrderByOn = True
1240      If gstrAccountNo <> vbNullString Then
1250        .frmAccountProfile_Sub.Form.MoveRec 0, gstrAccountNo  ' ** Form Procedure: frmAccountProfile_Sub.
1260      End If
1270      DoCmd.Hourglass False
1280    End With

EXITP:
1290    Exit Sub

ERRH:
1300    DoCmd.Hourglass False
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt to other forms and functionality.
        ' **   Notes:            E {tabContacts_pgNotes on frmAccountProfile_Sub}
        ' **   Short Name:       N {shortname on frmAccountProfile_Sub}
        ' **   Short Name PQ:    N {shortname_pq on frmAccountProfile_Sub}
        ' **   Contact Name 1:   1 {tabContacts_pgContact1 on frmAccountProfile_Sub}
        ' **   Contact Name 2:   2 {tabContacts_pgContact2 on frmAccountProfile_Sub}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   Contact1 Other:   H {opgLocContact1_optOther on frmAccountProfile_Sub}
        ' **   Next Record:      N {Form_KeyDown}
        ' **   Previous Record:  P {Form_KeyDown}
        ' **   Save:             S {cmdSave}
        ' **   Contact1 USA:     U {opgLocContact1_optUSA on frmAccountProfile_Sub}

        ' ** Shortcut Ctrl-Shift to other forms and functionality.
        ' **   Contact2 Other:   H {opgLocContact2_optOther on frmAccountProfile_Sub}
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Contact2 USA:     U {opgLocContact2_optUSA on frmAccountProfile_Sub}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        With Me
1480          intRetVal = 0
1490          blnEscPressed = True
1500          .frmAccountProfile_Sub.Form.EscPressed_Set blnEscPressed  ' ** Form Procedure: frmAccountProfile_Sub.
1510          cmdClose_Click  ' ** Procedure: Below.
1520        End With
1530      End Select
1540    End If

        ' ** Alt keys.
1550    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1560      Select Case intRetVal
          Case vbKeyE
1570        With Me
1580          intRetVal = 0
1590          If blnSysAcct = False Then
1600            .frmAccountProfile_Sub.SetFocus
1610            .frmAccountProfile_Sub.Form.FocusHolder.SetFocus
1620            .frmAccountProfile_Sub.Form.tabContacts = .frmAccountProfile_Sub.Form.tabContacts_pgNotes.PageIndex
1630            .frmAccountProfile_Sub.Form.Notes.SetFocus
1640            fSetScrollBarPosVT .frmAccountProfile_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
1650          End If
1660        End With
1670      Case vbKeyN
1680        With Me
1690          intRetVal = 0
1700          .frmAccountProfile_Sub.SetFocus
1710          If .frmAccountProfile_Sub.Form.shortname.Visible = True And .frmAccountProfile_Sub.Form.shortname.Enabled = True Then
1720            .frmAccountProfile_Sub.Form.shortname.SetFocus
1730          ElseIf .frmAccountProfile_Sub.Form.shortname_pq.Visible = True And .frmAccountProfile_Sub.Form.shortname_pq.Enabled = True Then
1740            .frmAccountProfile_Sub.Form.shortname_pq.SetFocus
1750          End If
1760        End With
1770      Case vbKeyX
1780        intRetVal = 0
1790        cmdClose_Click  ' ** Procedure: Below.
1800      Case vbKey1
1810        With Me
1820          intRetVal = 0
1830          If blnSysAcct = False Then
1840            .frmAccountProfile_Sub.SetFocus
1850            .frmAccountProfile_Sub.Form.FocusHolder.SetFocus
1860            .frmAccountProfile_Sub.Form.tabContacts = .frmAccountProfile_Sub.Form.tabContacts_pgContact1.PageIndex
1870            .frmAccountProfile_Sub.Form.Contact1_Name.SetFocus
1880            fSetScrollBarPosVT .frmAccountProfile_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
1890          End If
1900        End With
1910      Case vbKey2
1920        With Me
1930          intRetVal = 0
1940          If blnSysAcct = False Then
1950            .frmAccountProfile_Sub.SetFocus
1960            .frmAccountProfile_Sub.Form.FocusHolder.SetFocus
1970            .frmAccountProfile_Sub.Form.tabContacts = .frmAccountProfile_Sub.Form.tabContacts_pgContact2.PageIndex
1980            .frmAccountProfile_Sub.Form.Contact2_Name.SetFocus
1990            fSetScrollBarPosVT .frmAccountProfile_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
2000          End If
2010        End With
2020      End Select
2030    End If

        ' ** Ctrl keys.
2040    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2050      Select Case intRetVal
          Case vbKeyH
2060        With Me
2070          If blnSysAcct = False Then
2080            If .frmAccountProfile_Sub.Form.opgLocContact1 <> .frmAccountProfile_Sub.Form.opgLocContact1_optOther.OptionValue Then
2090              .frmAccountProfile_Sub.SetFocus
2100              .frmAccountProfile_Sub.Form.opgLocContact1 = .frmAccountProfile_Sub.Form.opgLocContact1_optOther.OptionValue
2110              .frmAccountProfile_Sub.Form.opgLocContact1_AfterUpdate  ' ** Form Procedure: frmAccountProfile_Sub.
2120            End If
2130          End If
2140        End With
2150      Case vbKeyN
2160        With Me
2170          intRetVal = 0
2180          .frmAccountProfile_Sub.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmAccountProfile_Sub.
2190        End With
2200      Case vbKeyP
2210        With Me
2220          intRetVal = 0
2230          .frmAccountProfile_Sub.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmAccountProfile_Sub.
2240        End With
2250      Case vbKeyS
2260        With Me
2270          intRetVal = 0
2280          .frmAccountProfile_Sub.SetFocus
2290          .frmAccountProfile_Sub.Form.cmdSave_Click THIS_PROC  ' ** Form Procedure: frmAccountProfile_Sub.
2300        End With
2310      Case vbKeyU
2320        With Me
2330          If blnSysAcct = False Then
2340            If .frmAccountProfile_Sub.Form.opgLocContact1 <> .frmAccountProfile_Sub.Form.opgLocContact1_optUSA.OptionValue Then
2350              .frmAccountProfile_Sub.SetFocus
2360              .frmAccountProfile_Sub.Form.opgLocContact1 = .frmAccountProfile_Sub.Form.opgLocContact1_optUSA.OptionValue
2370              .frmAccountProfile_Sub.Form.opgLocContact1_AfterUpdate  ' ** Form Procedure: frmAccountProfile_Sub.
2380            End If
2390          End If
2400        End With
2410      End Select
2420    End If

        ' ** Ctrl-Shift keys.
2430    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2440      Select Case intRetVal
          Case vbKeyH
2450        With Me
2460          If blnSysAcct = False Then
2470            If .frmAccountProfile_Sub.Form.opgLocContact2 <> .frmAccountProfile_Sub.Form.opgLocContact2_optOther.OptionValue Then
2480              .frmAccountProfile_Sub.SetFocus
2490              .frmAccountProfile_Sub.Form.opgLocContact2 = .frmAccountProfile_Sub.Form.opgLocContact2_optOther.OptionValue
2500              .frmAccountProfile_Sub.Form.opgLocContact2_AfterUpdate  ' ** Form Procedure: frmAccountProfile_Sub.
2510            End If
2520          End If
2530        End With
2540      Case vbKeyT
2550        intRetVal = 0
2560        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2570      Case vbKeyU
2580        With Me
2590          If blnSysAcct = False Then
2600            If .frmAccountProfile_Sub.Form.opgLocContact2 <> .frmAccountProfile_Sub.Form.opgLocContact2_optUSA.OptionValue Then
2610              .frmAccountProfile_Sub.SetFocus
2620              .frmAccountProfile_Sub.Form.opgLocContact2 = .frmAccountProfile_Sub.Form.opgLocContact2_optUSA.OptionValue
2630              .frmAccountProfile_Sub.Form.opgLocContact2_AfterUpdate  ' ** Form Procedure: frmAccountProfile_Sub.
2640            End If
2650          End If
2660        End With
2670      End Select
2680    End If

EXITP:
2690    KeyCode = intRetVal
2700    Exit Sub

ERRH:
2710    intRetVal = 0
2720    Select Case ERR.Number
        Case Else
2730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2740    End Select
2750    Resume EXITP

End Sub

Private Sub Form_Resize()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngSubHeight As Long, lngNavTop As Long, lngShortMenuTop As Long
        Dim lngTmp01 As Long

2810    With Me

2820      GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.

2830      If lngWidth <> lngFrm_Width Or lngHeight <> lngFrm_Height Then

2840        If lngWidth <= lngFrm_Width Then
              ' ** Don't let it shrink narrower than it's opening width.
2850          .cmdClose.Left = (lngClose_Left - (lngFrm_Width - lngWidth))
2860          .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
2870          .Sizable_lbl2.Left = (.cmdClose.Left + lngSizable_Offset)
2880        Else
2890          .cmdClose.Left = (lngClose_Left + (lngWidth - lngFrm_Width))
2900          .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
2910          .Sizable_lbl2.Left = (.cmdClose.Left + lngSizable_Offset)
2920          .frmAccountProfile_Sub.Width = (lngSub_Width + (lngWidth - lngFrm_Width))
2930          .frmAccountProfile_Sub_box.Width = ((lngSub_Width + (lngWidth - lngFrm_Width)) + (2& * lngTpp))
2940          lngTmp01 = ((.cmdClose.Left + .cmdClose.Width) + (9& * lngTpp))
2950          .Header_hline01.Width = lngTmp01
2960          .Header_hline02.Width = lngTmp01
2970          .Header_vline01.Left = lngTmp01
2980          .Header_vline02.Left = lngTmp01
2990          .Footer_hline01.Width = lngTmp01
3000          .Footer_hline02.Width = lngTmp01
3010          .Footer_vline01.Left = lngTmp01
3020          .Footer_vline02.Left = lngTmp01
3030          lngTmp01 = lngSubHLine_Width + (lngWidth - lngFrm_Width)
3040          .frmAccountProfile_Sub.Form.Footer_hline01.Width = lngTmp01
3050          .frmAccountProfile_Sub.Form.Footer_hline02.Width = lngTmp01
3060          .frmAccountProfile_Sub.Form.Footer_vline01.Left = lngTmp01
3070          .frmAccountProfile_Sub.Form.Footer_vline02.Left = lngTmp01
3080        End If

3090        lngSubHeight = .frmAccountProfile_Sub.Height
3100        If lngHeight > lngFrm_Height Then
3110          .Detail.Height = lngDetail_Height + (lngHeight - lngFrm_Height)
3120          .frmAccountProfile_Sub.Height = lngSub_Height + (lngHeight - lngFrm_Height)
3130          .frmAccountProfile_Sub_box.Height = (lngSub_Height + (lngHeight - lngFrm_Height)) + (6& * lngTpp)
3140          .Detail.Height = (.frmAccountProfile_Sub.Top + .frmAccountProfile_Sub.Height) + (32& * lngTpp)
3150          lngNavTop = ((.frmAccountProfile_Sub_box.Top + .frmAccountProfile_Sub_box.Height) - lngNavOffset_Top)
3160          .Nav_hline01.Top = lngNavTop
3170          .Nav_hline02.Top = (lngNavTop + lngTpp)
3180          .Nav_vline01.Top = lngNavTop
3190          .Nav_vline02.Top = lngNavTop
3200          .Nav_vline03.Top = lngNavTop
3210          .Nav_vline04.Top = lngNavTop
3220          .detail_height_line.Top = (.Detail.Height - (24& * lngTpp))
3230          lngShortMenuTop = ((.frmAccountProfile_Sub_box.Top + .frmAccountProfile_Sub_box.Height) + lngShortMenuOffset_Top)
3240          .ShortcutMenu_lbl.Top = lngShortMenuTop
3250        Else
3260          If (lngSub_Height - (lngFrm_Height - lngHeight)) > (.frmAccountProfile_Sub.Form.Detail.Height + (19& * lngTpp)) Then
3270            .frmAccountProfile_Sub.Height = (lngSub_Height - (lngFrm_Height - lngHeight))
3280            .frmAccountProfile_Sub_box.Height = (lngSub_Height - (lngFrm_Height - lngHeight)) + (6& * lngTpp)
3290            lngNavTop = ((.frmAccountProfile_Sub_box.Top + .frmAccountProfile_Sub_box.Height) - lngNavOffset_Top)
3300            .Nav_hline01.Top = lngNavTop
3310            .Nav_hline02.Top = (lngNavTop + lngTpp)
3320            .Nav_vline01.Top = lngNavTop
3330            .Nav_vline02.Top = lngNavTop
3340            .Nav_vline03.Top = lngNavTop
3350            .Nav_vline04.Top = lngNavTop
3360            .detail_height_line.Top = ((lngDetail_Height - (lngFrm_Height - lngHeight)) - (24& * lngTpp))
3370            .Detail.Height = (.frmAccountProfile_Sub.Top + .frmAccountProfile_Sub.Height) + (32& * lngTpp)
3380            lngShortMenuTop = ((.frmAccountProfile_Sub_box.Top + .frmAccountProfile_Sub_box.Height) + lngShortMenuOffset_Top)
3390            .ShortcutMenu_lbl.Top = lngShortMenuTop
3400            .detail_height_line.Top = (.Detail.Height - (24& * lngTpp))
3410          End If
3420        End If

3430      End If

3440    End With

EXITP:
3450    Exit Sub

ERRH:
3460    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
3470    Case Else
3480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3490    End Select
3500    Resume EXITP

End Sub

Private Sub Form_Timer()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

3610    With Me
3620      .TimerInterval = 0&
3630      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
3640        lngGTR_Stat = lngGTR_Stat + 1&
3650        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
3660        lngTmp01 = lngTmp01 + 1&
3670        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
3680        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
3690        If lngTmp01 > 1& Then
3700          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
3710        End If
3720        If lngTmp01 = 24& Then
3730          lngTmp01 = 0&
3740          lngTmp02 = lngTmp02 + 1&
3750          If lngTmp02 < 5& Then
3760            .TimerInterval = 50&
3770          Else
3780            blnGTR_Emblem = False
3790          End If
3800        Else
3810          .TimerInterval = 50&
3820        End If
3830      End If  ' ** blnGTR_Emblem.
3840      If gblnMessage = True Then
            ' ** Coming back from frmMasterBalance.
3850        gblnMessage = False
3860        If .frmAccountProfile_Sub.Form.dateclosed.Enabled = True Then
3870          .frmAccountProfile_Sub.Form.dateclosed = Null
3880        End If
3890        gblnSetFocus = True
3900        .TimerInterval = 100&
3910      ElseIf gblnSetFocus = True Then
3920        gblnSetFocus = False
3930        strFormName = vbNullString
3940  On Error Resume Next
3950        Do Until strFormName = THIS_NAME
3960          DoCmd.SelectObject acForm, THIS_NAME, False
3970          strFormName = Screen.ActiveForm.Name
3980        Loop
3990  On Error GoTo ERRH
4000        .frmAccountProfile_Sub.SetFocus
4010      End If
4020      If gblnGoToReport = True Then
4030        DoCmd.Hourglass True  ' ** Make sure it's still running.
4040        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
4050        Else
4060          DoCmd.Hourglass True  ' ** Make sure it's still running.
4070          If IsNull(garr_varGoToReport(GTR_FRM2)) = False Then
4080            If blnGTR_Emblem = False Then
4090              blnGTR_Emblem = True
4100              lngTmp01 = 0&: lngTmp02 = 0&
4110              .Sizable_lbl1.Visible = False
4120              .Sizable_lbl2.Visible = False
4130            End If
4140          End If
4150          DoEvents
4160          Select Case blnGoingToReport
              Case True
4170            Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                Case True
                  ' ** Just sit there, and we're done.
4180              DoCmd.Hourglass False
4190              blnGoingToReport = False
4200              .GoToReport_arw_pr_img.Visible = False
4210              gblnGoToReport = False
4220              blnGTR_Emblem = False
4230              GTREmblem_Off  ' ** Procedure: Below.
4240              .Sizable_lbl1.Visible = True
4250              .Sizable_lbl2.Visible = True
4260              .cmdClose.SetFocus
4270              DoEvents
4280            Case False
4290              .TimerInterval = 50&
4300              .frmAccountProfile_Sub.Form.cmdChoose_Click  ' ** Form Procedure: frmAccountProfile_Sub.
4310            End Select
4320          Case False
4330            blnGoingToReport = True
4340            Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                Case True
                  ' ** No more forms.
4350              .GoToReport_arw_pr_img.Visible = True
4360              If .cmdPrintReport.Enabled = True Then
4370                .cmdPrintReport.SetFocus
4380              Else
4390                .cmdClose.SetFocus
4400              End If
4410              DoEvents
4420              Beep
4430              DoCmd.Hourglass False
4440              .TimerInterval = (GTR_WAIT * 2&)
4450            Case False
4460              DoCmd.Hourglass True  ' ** Make sure it's still running.
4470              DoEvents
4480              lngBtn_Top = (.frmAccountProfile_Sub.Form.cmdChoose.Top + (0& * lngTpp))
4490              lngBtn_Left = ((.frmAccountProfile_Sub.Form.cmdChoose.Left + .frmAccountProfile_Sub.Form.cmdChoose.Width) - (lngTpp))
4500              .frmAccountProfile_Sub.Form.GoToReport_arw_relate_img.Top = lngBtn_Top
4510              .frmAccountProfile_Sub.Form.GoToReport_arw_relate_img.Left = lngBtn_Left
4520              .frmAccountProfile_Sub.Form.GoToReport_arw_relate_img.Visible = True
4530              .frmAccountProfile_Sub.Form.cmdChoose_lbl.Visible = False
4540              .frmAccountProfile_Sub.Form.cmdChoose_lbl2.Visible = False
4550              DoEvents
4560              .TimerInterval = 50&
4570            End Select
4580          End Select
4590        End If
4600      End If
4610      If gblnTimer = True Then
4620        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
4630          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
4640          gblnTimer = False
4650          DoCmd.Hourglass False
4660        End If
4670      End If
4680    End With

EXITP:
4690    Exit Sub

ERRH:
4700    Select Case ERR.Number
        Case Else
4710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4720    End Select
4730    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

4810    With Me

4820      DoCmd.Hourglass True
4830      DoEvents

4840      FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
4850      DoEvents

4860      Pref_Save THIS_NAME & "_Sub"  ' ** Module Procedure: modPreferenceFuncs.

4870      If blnRelAcctsCalled = True Then
4880        Set dbs = CurrentDb
4890        With dbs
              ' ** Empty tmpAccount_RelAccts.
4900          Set qdf = .QueryDefs("qryAccountProfile_RelAccts_02")
4910          qdf.Execute
4920          Set qdf = Nothing
4930          .Close
4940        End With
4950        Set dbs = Nothing
4960      End If

          ' ** Make sure all reports are closed.
4970      If Reports.Count > 0 Then
4980        Do While Reports.Count > 0
4990          DoCmd.Close acReport, Reports(0).Name
5000        Loop
5010      End If

5020      DoCmd.Hourglass False
5030      gstrFormQuerySpec = vbNullString
5040      gblnSetFocus = True
5050      gstrReturningForm = THIS_NAME
5060      Select Case strCallingForm
          Case "frmMenu_Account"
5070        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
5080      End Select

5090    End With

EXITP:
5100    Set qdf = Nothing
5110    Set dbs = Nothing
5120    Exit Sub

ERRH:
5130    DoCmd.Hourglass False
5140    Select Case ERR.Number
        Case Else
5150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5160    End Select
5170    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

5210    With Me
5220      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
5230        Select Case blnPrintReport_Focus
            Case True
5240          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
5250          .cmdPrintReport_raised_img.Visible = False
5260        Case False
5270          .cmdPrintReport_raised_img.Visible = True
5280          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5290        End Select
5300        .cmdPrintReport_raised_focus_img.Visible = False
5310        .cmdPrintReport_raised_focus_dots_img.Visible = False
5320        .cmdPrintReport_sunken_focus_dots_img.Visible = False
5330        .cmdPrintReport_raised_img_dis.Visible = False
5340      End If
5350    End With

EXITP:
5360    Exit Sub

ERRH:
5370    Select Case ERR.Number
        Case Else
5380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5390    End Select
5400    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

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

Public Sub cmdClose_Click()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean

5610    blnContinue = True

5620    With Me

5630      .TimerInterval = 0&

5640      If blnContinue = True Then

5650        .frmAccountProfile_Sub.Form.accountno.SetFocus

5660        If IsNull(.frmAccountProfile_Sub.Form.accounttype) = True And (IsNull(.frmAccountProfile_Sub.Form.accountno) = False And _
                IIf(Screen.ActiveControl.Name <> "accountno", True, _
                IIf(Trim(.frmAccountProfile_Sub.Form.accountno.text) = vbNullString, False, True))) Then
5670          blnContinue = False
5680          MsgBox "Please enter an account type to continue", vbInformation + vbOKOnly, "Entry Required"
5690          .frmAccountProfile_Sub.Form.accounttype.SetFocus
5700          DoCmd.CancelEvent
5710        Else
5720          If IsNull(.frmAccountProfile_Sub.Form.shortname) = True And (IsNull(.frmAccountProfile_Sub.Form.accountno) = False And _
                  IIf(Screen.ActiveControl.Name <> "accountno", True, _
                  IIf(Trim(.frmAccountProfile_Sub.Form.accountno.text) = vbNullString, False, True))) Then
5730            blnContinue = False
5740            MsgBox "Please enter a Short Name to continue", vbInformation + vbOKOnly, "Entry Required"
5750            .frmAccountProfile_Sub.Form.shortname.SetFocus
5760            DoCmd.CancelEvent
5770          Else
5780            If IsNull(.frmAccountProfile_Sub.Form.accountno) = True And blnAccountHasChanged = True Then
5790              blnContinue = False
5800              MsgBox "Please enter a valid account number.", vbInformation + vbOKOnly, "Entry Required"
5810              .frmAccountProfile_Sub.Form.accountno.SetFocus
5820              DoCmd.CancelEvent
5830            Else
5840              If IsNull(.frmAccountProfile_Sub.Form.accounttype) = True And blnAccountHasChanged = True Then
5850                blnContinue = False
5860                MsgBox "Please enter an account type to continue", vbInformation + vbOKOnly, "Entry Required"
5870                .frmAccountProfile_Sub.Form.accounttype.SetFocus
5880                DoCmd.CancelEvent
5890              Else
5900                If IsNull(.frmAccountProfile_Sub.Form.shortname) And blnAccountHasChanged = True Then
5910                  blnContinue = False
5920                  MsgBox "Please enter a Short Name to continue", vbInformation + vbOKOnly, "Entry Required"
5930                  .frmAccountProfile_Sub.Form.accounttype.SetFocus
5940                  DoCmd.CancelEvent
5950                End If
5960              End If
5970            End If
5980          End If
5990        End If

6000      End If  ' ** blnContinue.

6010      If blnContinue = True Then
6020        If gblnAccountNoWithType = True Then
6030          With Me
                ' ** Verify the user entered a valid account type prefix to the account number.
6040            If InStr(.frmAccountProfile_Sub.Form.accountno, "-") > 0 Then
6050              Set dbs = CurrentDb
6060              Set rst = dbs.OpenRecordset("SELECT accounttype.* FROM accounttype " & _
                    "WHERE accounttype.accounttype = '" & Left(.frmAccountProfile_Sub.Form.accountno, 2) & "'")
6070              If rst.EOF Then
6080                blnContinue = False
6090                MsgBox "The Account Type you entered as an Account Number prefix does not exist." & vbCrLf & _
                      "Please enter a valid Account Type to continue.", vbInformation + vbOKOnly, "Invalid Entry"
6100                .frmAccountProfile_Sub.Form.accountno.Undo
6110                DoCmd.CancelEvent
6120              Else
6130                rst.Close
6140                dbs.Close
6150                .frmAccountProfile_Sub.Form.accounttype = Left(.frmAccountProfile_Sub.Form.accountno, 2)
6160                .frmAccountProfile_Sub.Form.accounttype.Requery
6170              End If
6180            Else
6190              .frmAccountProfile_Sub.Form.accountno = .frmAccountProfile_Sub.Form.accounttype & "-" & _
                    Trim(.frmAccountProfile_Sub.Form.accountno)
6200            End If
6210          End With
6220        End If
6230      End If  ' ** blnContinue.

          ' ** VGC 11/01/2010: For now, I'm using an unused field for the default asset.
6240      If blnContinue = True Then
6250        With Me
6260          If IsNull(.frmAccountProfile_Sub.Form.taxlot) = True Then
6270            .frmAccountProfile_Sub.Form.taxlot = "0"
6280          Else
6290            If IsNumeric(.frmAccountProfile_Sub.Form.taxlot) = False Then
6300              .frmAccountProfile_Sub.Form.taxlot = "0"
6310            End If
6320          End If
6330        End With
6340      End If  ' ** blnContinue.

6350      If blnContinue = True Then

6360        .frmAccountProfile_Sub.Form.cmdSave_Click THIS_PROC  ' ** Form Procedure: frmAccountProfile_Sub.
6370        blnExit = True
6380        .frmAccountProfile_Sub.Form.Exit_Set blnExit  ' ** Form Procedure: frmAccountProfile_Sub.

6390        DoCmd.Close acForm, THIS_NAME

6400      End If  ' ** blnContinue.

6410    End With

EXITP:
6420    Set rst = Nothing
6430    Set qdf = Nothing
6440    Set dbs = Nothing
6450    Exit Sub

ERRH:
6460    Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
6470    Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
6480    Case Else
6490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6500    End Select
6510    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
6690          .frmAccountProfile_Sub.SetFocus
6700          .frmAccountProfile_Sub.Form.accountno.SetFocus
6710          fSetScrollBarPosVT .frmAccountProfile_Sub.Form, 1&  ' ** Module Function: modScrollBarFuncs.
6720        End With
6730      End Select
6740    End If

        ' ** Shift keys.
6750    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6760      Select Case intRetVal
          Case vbKeyTab
6770        With Me
6780          intRetVal = 0
6790          If .cmdPrintReport.Enabled = True Then
6800            .cmdPrintReport.SetFocus
6810          Else
6820            Select Case blnSysAcct
                Case True
6830              .frmAccountProfile_Sub.SetFocus
6840              .frmAccountProfile_Sub.Form.FocusHolder.SetFocus
6850              .frmAccountProfile_Sub.Form.tabContacts = .frmAccountProfile_Sub.Form.tabContacts_pgNotes.PageIndex
6860              .frmAccountProfile_Sub.Form.Notes.SetFocus
6870              fSetScrollBarPosVT .frmAccountProfile_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
6880            Case False
6890              .frmAccountProfile_Sub.SetFocus
6900              .frmAccountProfile_Sub.Form.FocusHolder.SetFocus
6910              .frmAccountProfile_Sub.Form.tabContacts = .frmAccountProfile_Sub.Form.tabContacts_pgContact1.PageIndex
6920              .frmAccountProfile_Sub.Form.Contact1_Email.SetFocus
6930              fSetScrollBarPosVT .frmAccountProfile_Sub.Form, 999&  ' ** Module Function: modScrollBarFuncs.
6940            End Select
6950          End If
6960        End With
6970      End Select
6980    End If

EXITP:
6990    KeyCode = intRetVal
7000    Exit Sub

ERRH:
7010    intRetVal = 0
7020    Select Case ERR.Number
        Case Else
7030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7040    End Select
7050    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

7110    With Me
7120      blnPrintReport_Focus = True
7130      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
7140      .cmdPrintReport_raised_img.Visible = False
7150      .cmdPrintReport_raised_focus_img.Visible = False
7160      .cmdPrintReport_raised_focus_dots_img.Visible = False
7170      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7180      .cmdPrintReport_raised_img_dis.Visible = False
7190    End With

EXITP:
7200    Exit Sub

ERRH:
7210    Select Case ERR.Number
        Case Else
7220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7230    End Select
7240    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

7310    With Me
7320      blnPrintReport_MouseDown = True
7330      .cmdPrintReport_sunken_focus_dots_img.Visible = True
7340      .cmdPrintReport_raised_img.Visible = False
7350      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7360      .cmdPrintReport_raised_focus_img.Visible = False
7370      .cmdPrintReport_raised_focus_dots_img.Visible = False
7380      .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_Click()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String
        Dim intMode As Integer

7510    With Me
7520      gstrFormQuerySpec = THIS_NAME
7530      Select Case gblnSeparateCheckingAccounts
          Case True
7540        strDocName = "rptAccountProfile_SepChkAcct"
7550      Case False
7560        strDocName = "rptAccountProfile_SngChkAcct"
7570      End Select
7580      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
7590        intMode = acViewPreview
7600      Else
7610        intMode = acViewNormal
7620      End If
7630      gstrAccountNo = .frmAccountProfile_Sub.Form.accountno
          '##GTR_Ref: rptAccountProfile_SepChkAcct
          '##GTR_Ref: rptAccountProfile_SngChkAcct
7640      DoCmd.OpenReport strDocName, intMode
7650      If intMode = acViewPreview Then
7660        DoCmd.Maximize
7670        DoCmd.RunCommand acCmdFitToWindow
7680      End If
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

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7820    intShiftDown = (Shift And acShiftMask) > 0
7830    intAltDown = (Shift And acAltMask) > 0
7840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7860      Select Case intRetVal
          Case vbKeyTab
7870        With Me
7880          intRetVal = 0
7890          .cmdClose.SetFocus
7900        End With
7910      End Select
7920    End If

        ' ** Shift keys.
7930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7940      Select Case intRetVal
          Case vbKeyTab
7950        With Me
7960          intRetVal = 0
7970          .frmAccountProfile_Sub.SetFocus
7980        End With
7990      End Select
8000    End If

EXITP:
8010    KeyCode = intRetVal
8020    Exit Sub

ERRH:
8030    intRetVal = 0
8040    Select Case ERR.Number
        Case Else
8050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8060    End Select
8070    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

8110    With Me
8120      If blnPrintReport_MouseDown = False Then
8130        Select Case blnPrintReport_Focus
            Case True
8140          .cmdPrintReport_raised_focus_dots_img.Visible = True
8150          .cmdPrintReport_raised_focus_img.Visible = False
8160        Case False
8170          .cmdPrintReport_raised_focus_img.Visible = True
8180          .cmdPrintReport_raised_focus_dots_img.Visible = False
8190        End Select
8200        .cmdPrintReport_raised_img.Visible = False
8210        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8220        .cmdPrintReport_sunken_focus_dots_img.Visible = False
8230        .cmdPrintReport_raised_img_dis.Visible = False
8240      End If
8250    End With

EXITP:
8260    Exit Sub

ERRH:
8270    Select Case ERR.Number
        Case Else
8280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8290    End Select
8300    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

8410    With Me
8420      .cmdPrintReport_raised_focus_dots_img.Visible = True
8430      .cmdPrintReport_raised_img.Visible = False
8440      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8450      .cmdPrintReport_raised_focus_img.Visible = False
8460      .cmdPrintReport_sunken_focus_dots_img.Visible = False
8470      .cmdPrintReport_raised_img_dis.Visible = False
8480      blnPrintReport_MouseDown = False
8490    End With

EXITP:
8500    Exit Sub

ERRH:
8510    Select Case ERR.Number
        Case Else
8520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8530    End Select
8540    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

8610    With Me
8620      .cmdPrintReport_raised_img.Visible = True
8630      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8640      .cmdPrintReport_raised_focus_img.Visible = False
8650      .cmdPrintReport_raised_focus_dots_img.Visible = False
8660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
8670      .cmdPrintReport_raised_img_dis.Visible = False
8680      blnPrintReport_Focus = False
8690    End With

EXITP:
8700    Exit Sub

ERRH:
8710    Select Case ERR.Number
        Case Else
8720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8730    End Select
8740    Resume EXITP

End Sub

Private Sub frmAccountProfile_Sub_Enter()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "frmAccountProfile_Sub_Enter"

        Dim Button As Integer, Shift As Integer, X As Single, Y As Single

8810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8820    Exit Sub

ERRH:
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

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

Public Sub HasSaved_Set(strThisAcct As String, blnSaved As Boolean)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "HasSaved_Set"

9010    strThisAccountNo = strThisAcct
9020    blnNewHasSaved = blnSaved

EXITP:
9030    Exit Sub

ERRH:
9040    Select Case ERR.Number
        Case Else
9050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9060    End Select
9070    Resume EXITP

End Sub

Public Sub EscPressed_Set(blnEscape As Boolean)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "EscPressed_Set"

9110    blnEscPressed = blnEscape

EXITP:
9120    Exit Sub

ERRH:
9130    Select Case ERR.Number
        Case Else
9140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9150    End Select
9160    Resume EXITP

End Sub

Public Sub AcctChanged_Set(blnAcctChanged As Boolean)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "AcctChanged_Set"

9210    blnAccountHasChanged = blnAcctChanged

EXITP:
9220    Exit Sub

ERRH:
9230    Select Case ERR.Number
        Case Else
9240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9250    End Select
9260    Resume EXITP

End Sub

Public Sub SysAcct_Set(blnIsSysAcct As Boolean)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "SysAcct_Set"

9310    blnSysAcct = blnIsSysAcct

EXITP:
9320    Exit Sub

ERRH:
9330    Select Case ERR.Number
        Case Else
9340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9350    End Select
9360    Resume EXITP

End Sub

Public Sub RelAccts_Set(blnCalledRA As Boolean)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "RelAccts_Set"

9410    blnRelAcctsCalled = blnCalledRA

EXITP:
9420    Exit Sub

ERRH:
9430    Select Case ERR.Number
        Case Else
9440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9450    End Select
9460    Resume EXITP

End Sub

Private Sub ChkAccountLimit()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkAccountLimit"

        Dim lngCurrentAccounts As Long

9510    lngCurrentAccounts = DCount("*", "account")

9520    If gstrTrustDataLocation = vbNullString Then
9530      IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
9540    End If

9550    If Right(gstrTrustDataLocation, 1) <> LNK_SEP Then gstrTrustDataLocation = gstrTrustDataLocation & LNK_SEP

9560    If lngMaxAccounts = 0& Then
9570      lngMaxAccounts = DecodeString(IniFile_Get("License", "Limit", "", gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modStartupFuncs.
9580    End If

EXITP:
9590    Exit Sub

ERRH:
9600    DoCmd.Hourglass False
9610    Select Case ERR.Number
        Case Else
9620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9630    End Select
9640    Resume EXITP

End Sub

Public Sub RestoreForm()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "RestoreForm"

9710    DoCmd.SelectObject acForm, THIS_NAME, False
9720    DoCmd.Restore

EXITP:
9730    Exit Sub

ERRH:
9740    Select Case ERR.Number
        Case Else
9750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9760    End Select
9770    Resume EXITP

End Sub

Public Sub GTREmblem_Set(blnSetEmblem As Boolean)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Set"

9810    With Me
9820      blnGTR_Emblem = blnSetEmblem
9830      If blnGTR_Emblem = False Then
9840        .TimerInterval = 0&
9850        blnGoingToReport = False
9860        GTREmblem_Off  ' ** Procedure: Below.
9870        .Sizable_lbl1.Visible = True
9880        .Sizable_lbl2.Visible = True
9890      End If
9900    End With

EXITP:
9910    Exit Sub

ERRH:
9920    DoCmd.Hourglass False
9930    Select Case ERR.Number
        Case Else
9940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9950    End Select
9960    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

10010   With Me
10020     blnGTR_Emblem = False
10030     For lngX = 1& To 24&
10040       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
10050     Next
10060   End With

EXITP:
10070   Exit Sub

ERRH:
10080   Select Case ERR.Number
        Case Else
10090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10100   End Select
10110   Resume EXITP

End Sub
