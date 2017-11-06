VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Account"

'VGC 10/26/2017: CHANGES!

' ** For GTR, shows arrows, sets focus, then continues.

' ** Shortcut Alt keys responsive from this form:
' **   Add Account:      A {cmdAdd}
' **   Close:            C {cmdClose}
' **   Delete Account:   D {cmdDelete}
' **   Search:           E {cmdSearch}
' **   Hide:             H {cmdHideTransactions}
' **   Inc/Exp Codes:    I {cmdEditIECodes}
' **   Display List:     L {opgSubForm_optList}
' **   Comments:         M {cmdEditComments}
' **   Short Name:       N {shortname on frmMenu_Account_Sub_One, frmMenu_Account_Sub_List}
' **   Display One:      O {opgSubForm_optOne}
' **   Profile:          P {cmdAccountProfile}
' **   Transactions:     R {cmdTransactions}
' **   Assets:           S {cmdAssets}
' **   Tax Codes:        T {cmdEditTaxCodes}
' **   Include Closed:   U {chkIncludeClosed}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next Record:      N {MoveRec on frmMenu_Account_Sub_Nav}
' **   Previous Record:  P {MoveRec on frmMenu_Account_Sub_Nav}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Unlock Var:       C {blnFormOpening: Form_KeyDown}
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

' ***********************************************
' ** Easter Egg (Secret Key):
' ** cmdHideTransactions_KeyDown
' ***********************************************

' ** Array: arr_varCtl().
Private lngCtls As Long, arr_varCtl() As Variant
Const C_ELEMS As Integer = 5  ' ** Array's first-element UBound().
Const C_CNAM As Integer = 0
Const C_CTYP As Integer = 1
Const C_TOP  As Integer = 2
Const C_LFT  As Integer = 3
Const C_WDT  As Integer = 4
Const C_HGT  As Integer = 5

Private lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngDetail_Height As Long, lngSizable_Offset As Long
Private lngGTRArrow_Offset As Long, lngGTREmblem_Offset As Long
Private lngClose_Left As Long, lngAdd_Left As Long, lngDelete_Left As Long

Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnJustClose As Boolean, blnFormOpening As Boolean, blnEmptyHidden As Boolean
Private lngMaxAccounts As Long, lngTpp As Long, strFormClose As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim grp As DAO.Group, usr As DAO.User
        Dim blnFound As Boolean
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = Trim(.OpenArgs)

150           gblnSignal = False

              ' ** Just get these out of the way.
160           MoveToZero  ' ** Procedure: Below
170           DoEvents

              ' ** Variables are fed empty, then populated ByRef.
180           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

190           lngCtls = 0&
200           ReDim arr_varCtl(C_ELEMS, 0)

210           lngForm_Width = .Width
220           lngDetail_Height = .Detail.Height
230           lngClose_Left = .cmdClose.Left
240           lngAdd_Left = .cmdAdd.Left
250           lngDelete_Left = .cmdDelete.Left
260           lngSizable_Offset = (.Width - .Sizable_lbl1.Left)
270           lngGTRArrow_Offset = (.cmdAdd.Left - .GoToReport_arw_add_img.Left)
280           lngGTREmblem_Offset = (.Width - .GoToReport_Emblem_01_img.Left)
290           CtlArray_Load  ' ** Procedure: Below.

300           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

310           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

320           If gblnGoToReport = True Then
330             .GoToReport_Emblem_01_img.Visible = True
340             DoEvents
350           End If

              ' ** This will be turned on for specific menu items.
360           blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.
370           SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
              'SystemColor_Move Me  ' ** Module Procedure: modSystemColorFuncs.
              'lngTpp = GetTPP  ' ** Module Function: modWindowsFunctions.
380           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Check limit to the number of accounts allowed by the license.
390           ChkAccountLimit  ' ** Procedure: Below.

400           gstrAccountNo = vbNullString: gstrAccountName = vbNullString
410           blnJustClose = False: blnFormOpening = False: blnEmptyHidden = False
420           blnGTR_Emblem = False: lngGTR_Stat = 0&

              ' ** Check user status, and disable add/delete if view only.
430           blnFound = False
440           For Each grp In DBEngine.Workspaces(0).Groups
450             If grp.Name = "ViewOnly" Then
460               For Each usr In grp.Users
470                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
480                   blnFound = True
490                   Exit For
500                 End If
510               Next
520             End If
530           Next

540           If blnFound = True Then
550             .cmdAdd.Enabled = False
560             .cmdDelete.Enabled = False
570           End If

              ' ** Loop through all the security groups, checking if our current user is in the admins group.
580           gblnAdmin = False
590           For Each grp In DBEngine.Workspaces(0).Groups
600             If grp.Name = "Admins" Then
610               For Each usr In grp.Users
620                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
630                   gblnAdmin = True
640                   Exit For
650                 End If
660               Next
670             End If
680           Next

690           Select Case .chkLoadTimer_vis
              Case True
700             .chkLoadTimer.Visible = True
710             .chkLoadTimer_lbl1.Visible = True
720             .cmdDelete.Caption = "Delete Account"
730             .cmdDelete.StatusBarText = "Delete the current account - Ctrl+D"
740             .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Ctrl+D"
750           Case False
760             .chkLoadTimer.Visible = False
770             .chkLoadTimer_lbl1.Visible = False
780             .cmdDelete.Caption = "&Delete Account"
790             .cmdDelete.StatusBarText = "Delete the current account - Alt+D"
800             .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Alt+D"
810           End Select
820           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
830           DoEvents

              'If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
              '  .cmdHideTransactions_lbl2.Visible = True
              '  .cmdHideTransactions_lbl3.Visible = True
              '  .HideTrans_Asterisk_img.Visible = True
              'End If

840         Else
850           Cancel = -1
860         End If
870       Else
880         Cancel = -1
890       End If
900     End With

910     If Cancel = -1 Then
920       Beep
930       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
940     End If

EXITP:
950     Set usr = Nothing
960     Set grp = Nothing
970     Exit Sub

ERRH:
980     DoCmd.Hourglass False
990     Select Case ERR.Number
        Case Else
1000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1010    End Select
1020    Resume EXITP

End Sub

Private Sub Form_Load()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim strSortNow As String
        Dim strTmp01 As String

1110    With Me

1120      opgSubform_AfterUpdate  ' ** Procedure: Below.
1130      DoEvents

1140      chkIncludeClosed_AfterUpdate  ' ** Procedure: Below.
1150      DoEvents

1160      strSortNow = "[acct_sort]"
1170      .frmMenu_Account_Sub_Nav.Form.OrderBy = strSortNow
1180      .frmMenu_Account_Sub_Nav.Form.OrderByOn = True
1190  On Error Resume Next
1200      strTmp01 = .frmMenu_Account_Sub_One.Form.Name
1210  On Error GoTo ERRH
1220      If strTmp01 = "frmMenu_Account_Sub_One" Then
1230        .frmMenu_Account_Sub_One.Form.OrderBy = strSortNow
1240        .frmMenu_Account_Sub_One.Form.OrderByOn = True
1250      End If
1260  On Error Resume Next
1270      strTmp01 = .frmMenu_Account_Sub_List.Form.Name
1280  On Error GoTo ERRH
1290      If strTmp01 = "frmMenu_Account_Sub_List" Then
1300        .frmMenu_Account_Sub_List.Form.OrderBy = strSortNow
1310        .frmMenu_Account_Sub_List.Form.OrderByOn = True
1320      End If

1330      .cmdAccountProfile.SetFocus

1340      DoCmd.Hourglass False

1350    End With

EXITP:
1360    Exit Sub

ERRH:
1370    DoCmd.Hourglass False
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strControlName As String
        Dim intRetVal As Integer, Cancel As Integer

1510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1520    intShiftDown = (Shift And acShiftMask) > 0
1530    intAltDown = (Shift And acAltMask) > 0
1540    intCtrlDown = (Shift And acCtrlMask) > 0

1550  On Error Resume Next
1560    strControlName = Screen.ActiveControl.Name
1570  On Error GoTo ERRH

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Short Name:       N {shortname on frmMenu_Account_Sub_One, frmMenu_Account_Sub_List}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next Record:      N {MoveRec on frmMenu_Account_Sub_Nav}
        ' **   Previous Record:  P {MoveRec on frmMenu_Account_Sub_Nav}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Unlock Var:       C {blnFormOpening: Form_KeyDown}
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
1580    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1590      Select Case intRetVal
          Case vbKeyEscape
1600        intRetVal = 0
1610        cmdClose_Click  ' ** Procedure: Below.
1620      End Select
1630    End If

        ' ** Alt keys.
1640    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1650      Select Case intRetVal
          Case vbKeyD
1660        With Me
1670          intRetVal = 0
1680          Select Case .chkLoadTimer.Visible
              Case True
1690            If .chkLoadTimer.Enabled = True Then
1700              .chkLoadTimer.SetFocus
1710            End If
1720            .chkLoadTimer = (Not .chkLoadTimer)
1730            chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
1740          Case False
1750            If .cmdDelete.Enabled = True Then
1760              .cmdDelete.SetFocus
1770              cmdDelete_Click  ' ** Procedure: Below.
1780            Else
1790              Beep
1800            End If
1810          End Select
1820        End With
1830      Case vbKeyN
1840        With Me
1850          intRetVal = 0
1860          Select Case .opgSubForm
              Case .opgSubForm_optOne.OptionValue
1870            .frmMenu_Account_Sub_One.SetFocus
1880            .frmMenu_Account_Sub_One.Form.shortname.SetFocus
1890          Case .opgSubForm_optList.OptionValue
1900            .frmMenu_Account_Sub_List.SetFocus
1910            .frmMenu_Account_Sub_List.Form.shortname.SetFocus
1920          End Select
1930        End With
1940      Case vbKeyX
1950        intRetVal = 0
1960        cmdClose_Click  ' ** Procedure: Below.
1970      End Select
1980    End If

        ' ** Ctrl keys.
1990    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2000      Select Case intRetVal
          Case vbKeyD
2010        With Me
2020          intRetVal = 0
2030          If .chkLoadTimer.Visible = True Then
2040            .cmdDelete.SetFocus
2050            cmdDelete_Click  ' ** Procedure: Below.
2060          Else
2070            Beep
2080          End If
2090        End With
2100      Case vbKeyN
2110        With Me
2120          intRetVal = 0
2130          Select Case .opgSubForm
              Case .opgSubForm_optOne.OptionValue
2140            .frmMenu_Account_Sub_Nav.SetFocus
2150            .frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmMenu_Account_Sub_One.
2160            .frmMenu_Account_Sub_One.SetFocus
2170          Case .opgSubForm_optList.OptionValue
2180            .frmMenu_Account_Sub_Nav.SetFocus
2190            .frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmMenu_Account_Sub_List.
2200            .frmMenu_Account_Sub_List.SetFocus
2210          End Select
2220          If strControlName <> vbNullString Then
2230            .Controls(strControlName).SetFocus
2240          End If
2250        End With
2260      Case vbKeyP
2270        With Me
2280          intRetVal = 0
2290          Select Case .opgSubForm
              Case .opgSubForm_optOne.OptionValue
2300            .frmMenu_Account_Sub_Nav.SetFocus
2310            .frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmMenu_Account_Sub_One.
2320            .frmMenu_Account_Sub_One.SetFocus
2330          Case .opgSubForm_optList.OptionValue
2340            .frmMenu_Account_Sub_Nav.SetFocus
2350            .frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmMenu_Account_Sub_List.
2360            .frmMenu_Account_Sub_List.SetFocus
2370          End Select
2380          If strControlName <> vbNullString Then
2390            .Controls(strControlName).SetFocus
2400          End If
2410        End With
2420      End Select
2430    End If

        ' ** Ctrl-Shift keys.
2440    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2450      Select Case intRetVal
          Case vbKeyC
2460        With Me
              ' ** In case this gets stuck on.
2470          FormCloseVar False  ' ** blnFormOpening = False
2480          DoBeeps 2, 300  ' ** Module Function: modWindowFunctions.
2490        End With
2500      Case vbKeyF
2510        With Me
2520          intRetVal = 0
2530          .FocusHolder.SetFocus
2540        End With
2550      Case vbKeyT
2560        With Me
2570          intRetVal = 0  ' ** User doesn't matter.
2580          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
2590            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2600          ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Module Function: modFileUtilities.
2610            Select Case .chkLoadTimer_vis
                Case True
2620              .FocusHolder.SetFocus
2630              DoEvents
2640              .chkLoadTimer.Visible = False
2650              .chkLoadTimer_lbl1.Visible = False
2660              .cmdDelete.Caption = "&Delete Account"
2670              .cmdDelete.StatusBarText = "Delete the current account - Alt+D"
2680              .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Alt+D"
2690            Case False
2700              .chkLoadTimer.Visible = True
2710              .chkLoadTimer_lbl1.Visible = True
2720              .cmdDelete.Caption = "Delete Account"
2730              .cmdDelete.StatusBarText = "Delete the current account - Ctrl+D"
2740              .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Ctrl+D"
2750            End Select
2760          End If
2770          .chkLoadTimer_vis = .chkLoadTimer.Visible
2780        End With
2790      Case vbKeyX
2800        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
2810          intRetVal = 0
2820          blnJustClose = True
2830          SetOption_Dev  ' ** Module Function: modUtilities.
2840          gblnSignal = True
2850          DoCmd.Close acForm, THIS_NAME
2860          DoCmd.Hourglass False
2870        End If
2880      End Select
2890    End If

EXITP:
2900    KeyCode = intRetVal
2910    Exit Sub

ERRH:
2920    intRetVal = 0
2930    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2940    Case 3709  ' ** The search key was not found in any record.
          ' ** I have no idea!
2950    Case Else
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Private Sub Form_Resize()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

        Dim lngTop As Long, lngLeft As Long, lngWidth As Long, lngHeight As Long
        Dim lngWidth_Diff As Long, lngHeight_Diff As Long, lngNewForm_Width As Long, lngNewDetail_Height As Long
        Dim intModeW As Integer, intModeH As Integer
        Dim strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long
        Dim lngX As Long, lngY As Long

3010    With Me

3020      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowsFunctions.
3030        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
3040      End If

          ' ** Variables are fed empty, then populated ByRef.
3050      GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

3060      lngWidth_Diff = 0&: lngHeight_Diff = 0&: lngNewForm_Width = 0&: lngNewDetail_Height = 0&
3070      intModeW = 0: intModeH = 0

          ' ** Width.
3080      If lngWidth > lngFrm_Width Then
3090        lngWidth_Diff = (lngWidth - lngFrm_Width)
3100        lngNewForm_Width = (lngForm_Width + lngWidth_Diff)
3110        intModeW = 1
3120      ElseIf lngWidth < lngFrm_Width Then
3130        lngWidth_Diff = 0&
3140        lngNewForm_Width = lngForm_Width
3150        intModeW = 2
3160      Else
3170        lngWidth_Diff = 0&
3180        lngNewForm_Width = lngForm_Width
3190        intModeW = 3
3200      End If

          ' ** Height.
3210      If lngHeight > lngFrm_Height Then
3220        lngHeight_Diff = (lngHeight - lngFrm_Height)
3230        lngNewDetail_Height = (lngDetail_Height + lngHeight_Diff)
3240        intModeH = 1
3250      ElseIf lngHeight < lngFrm_Height Then
3260        lngHeight_Diff = 0&
3270        lngNewDetail_Height = lngDetail_Height
3280        intModeH = 2
3290      Else
3300        lngHeight_Diff = 0&
3310        lngNewDetail_Height = lngDetail_Height
3320        intModeH = 3
3330      End If

          ' ** Width.
3340      Select Case intModeW
          Case 1
            ' ** Wider
3350        .Width = lngNewForm_Width
3360        .cmdClose.Left = (lngClose_Left + lngWidth_Diff)
3370        lngTmp02 = Int(lngWidth_Diff / 2)
3380        .cmdAdd.Left = (lngAdd_Left + lngTmp02)
3390        .GoToReport_arw_add_img.Left = (.cmdAdd.Left - lngGTRArrow_Offset)
3400        .cmdDelete.Left = (lngDelete_Left + lngTmp02)
3410        .Sizable_lbl1.Left = (lngNewForm_Width - lngSizable_Offset)
3420        .Sizable_lbl2.Left = (lngNewForm_Width - lngSizable_Offset)
3430        For lngX = 1& To 24&
3440          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Left = (lngNewForm_Width - lngGTREmblem_Offset)
3450        Next
3460        For lngX = 0& To (lngCtls - 1&)
3470          Select Case arr_varCtl(C_CTYP, lngX)
              Case acSubform
3480            .Controls(arr_varCtl(C_CNAM, lngX)).Width = (arr_varCtl(C_WDT, lngX) + lngWidth_Diff)
3490          Case acCommandButton, acBoundObjectFrame
3500            lngTmp02 = Int(lngWidth_Diff / 2)
3510            .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) + lngTmp02)
3520          Case acLine
3530            If Left(arr_varCtl(C_CNAM, lngX), 8) = "Detail_v" Then
3540              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
3550              Select Case lngTmp02
                  Case 1, 2, 3, 4, 13, 14, 15, 16
                    ' ** Detail_vline01,2,3,4,13,14,15,16.
3560                .Controls(arr_varCtl(C_CNAM, lngX)).Left = lngNewForm_Width
3570              Case 11, 12, 17, 18, 19, 20, 21, 22
                    ' ** Detail_vline11,12,17,18,19,20,21,22.
3580                .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) + lngWidth_Diff)
3590              End Select
3600            ElseIf Left(arr_varCtl(C_CNAM, lngX), 8) = "Detail_h" Then
3610              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
3620              Select Case lngTmp02
                  Case 1, 2, 3, 4, 13, 14, 15, 16
                    ' ** Detail_hline01,2,3,4,13,14,15,16.
3630                .Controls(arr_varCtl(C_CNAM, lngX)).Width = lngNewForm_Width
3640              End Select
3650            ElseIf Left(arr_varCtl(C_CNAM, lngX), 5) = "Nav_h" Then
3660              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
3670              Select Case lngTmp02
                  Case 3
                    ' ** Nav_hline03.
3680                .Controls(arr_varCtl(C_CNAM, lngX)).Width = (arr_varCtl(C_WDT, lngX) + lngWidth_Diff)
3690              End Select
3700            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "cmd" Then
3710              lngTmp02 = Int(lngWidth_Diff / 2)
3720              .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) + lngTmp02)
3730            End If
3740          Case acRectangle, acCheckBox, acLabel, acOptionGroup, acOptionButton
3750            If Left(arr_varCtl(C_CNAM, lngX), 3) = "chk" Then
3760              .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) + lngWidth_Diff)
3770            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "opg" Then
3780              lngTmp02 = (arr_varCtl(C_LFT, lngX) + lngWidth_Diff)
3790              If lngTmp02 > .Controls(arr_varCtl(C_CNAM, lngX)).Left Then
3800                lngTmp03 = (lngTmp02 - .Controls(arr_varCtl(C_CNAM, lngX)).Left)
3810                lngTmp04 = Int(lngTmp03 / lngTpp)
3820                For lngY = 1 To lngTmp04
3830                  .Controls(arr_varCtl(C_CNAM, lngX)).Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left + lngTpp)
3840                Next  ' ** lngY.
3850              ElseIf lngTmp02 < .Controls(arr_varCtl(C_CNAM, lngX)).Left Then
3860                lngTmp03 = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTmp02)
3870                lngTmp04 = Int(lngTmp03 / lngTpp)
3880                For lngY = 1 To lngTmp04
3890                  .Controls(arr_varCtl(C_CNAM, lngX)).Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
3900                Next  ' ** lngY.
3910              End If
3920            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "Nav" Then
3930              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
3940              Select Case lngTmp02
                  Case 1
                    ' ** Nav_box01
3950                .Controls(arr_varCtl(C_CNAM, lngX)).Width = (arr_varCtl(C_WDT, lngX) + lngWidth_Diff)
3960              Case 3
3970                .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) + lngWidth_Diff)
3980              End Select
3990            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "cmd" Then
                  ' ** cmdAccount_l_box01, cmdAccount_r_box01
4000              lngTmp02 = Int(lngWidth_Diff / 2)
4010              .Controls(arr_varCtl(C_CNAM, lngX)).Left = (arr_varCtl(C_LFT, lngX) + lngTmp02)
4020            End If
4030          End Select
4040        Next
4050      Case 2
            ' ** Narrower.
            ' ** No narrower than original.
4060      Case 3
            ' ** Original.
4070        .Width = lngForm_Width
4080        .cmdClose.Left = lngClose_Left
4090        .cmdAdd.Left = lngAdd_Left
4100        .GoToReport_arw_add_img.Left = (.cmdAdd.Left - lngGTRArrow_Offset)
4110        .cmdDelete.Left = lngDelete_Left
4120        .Sizable_lbl1.Left = (lngForm_Width - lngSizable_Offset)
4130        .Sizable_lbl2.Left = (lngForm_Width - lngSizable_Offset)
4140        For lngX = 1& To 24&
4150          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Left = (lngForm_Width - lngGTREmblem_Offset)
4160        Next
4170        For lngX = 0& To (lngCtls - 1&)
4180          Select Case arr_varCtl(C_CTYP, lngX)
              Case acSubform
4190            .Controls(arr_varCtl(C_CNAM, lngX)).Width = arr_varCtl(C_WDT, lngX)
4200          Case acCommandButton, acBoundObjectFrame
4210            .Controls(arr_varCtl(C_CNAM, lngX)).Left = arr_varCtl(C_LFT, lngX)
4220          Case acLine
4230            If Left(arr_varCtl(C_CNAM, lngX), 8) = "Detail_v" Then
4240              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
4250              Select Case lngTmp02
                  Case 1, 2, 3, 4, 13, 14, 15, 16
4260                .Controls(arr_varCtl(C_CNAM, lngX)).Left = lngForm_Width
4270              Case 11, 12, 17, 18, 19, 20, 21, 22
4280                .Controls(arr_varCtl(C_CNAM, lngX)).Left = arr_varCtl(C_LFT, lngX)
4290              End Select
4300            ElseIf Left(arr_varCtl(C_CNAM, lngX), 8) = "Detail_h" Then
4310              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
4320              Select Case lngTmp02
                  Case 1, 2, 3, 4, 13, 14, 15, 16
4330                .Controls(arr_varCtl(C_CNAM, lngX)).Width = lngForm_Width
4340              End Select
4350            ElseIf Left(arr_varCtl(C_CNAM, lngX), 5) = "Nav_h" Then
4360              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
4370              Select Case lngTmp02
                  Case 3
4380                .Controls(arr_varCtl(C_CNAM, lngX)).Width = arr_varCtl(C_WDT, lngX)
4390              End Select
4400            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "cmd" Then
4410              .Controls(arr_varCtl(C_CNAM, lngX)).Left = arr_varCtl(C_LFT, lngX)
4420            End If
4430          Case acRectangle, acCheckBox, acLabel, acOptionGroup, acOptionButton
4440            If Left(arr_varCtl(C_CNAM, lngX), 3) = "chk" Then
4450              .Controls(arr_varCtl(C_CNAM, lngX)).Left = arr_varCtl(C_LFT, lngX)
4460            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "opg" Then
4470              lngTmp02 = arr_varCtl(C_LFT, lngX)
4480              If lngTmp02 > .Controls(arr_varCtl(C_CNAM, lngX)).Left Then
4490                lngTmp03 = (lngTmp02 - .Controls(arr_varCtl(C_CNAM, lngX)).Left)
4500                lngTmp04 = Int(lngTmp03 / lngTpp)
4510                For lngY = 1 To lngTmp04
4520                  .Controls(arr_varCtl(C_CNAM, lngX)).Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left + lngTpp)
4530                Next  ' ** lngY.
4540              ElseIf lngTmp02 < .Controls(arr_varCtl(C_CNAM, lngX)).Left Then
4550                lngTmp03 = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTmp02)
4560                lngTmp04 = Int(lngTmp03 / lngTpp)
4570                For lngY = 1 To lngTmp04
4580                  .Controls(arr_varCtl(C_CNAM, lngX)).Left = (.Controls(arr_varCtl(C_CNAM, lngX)).Left - lngTpp)
4590                Next  ' ** lngY.
4600              End If
4610            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "Nav" Then
4620              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
4630              Select Case lngTmp02
                  Case 1
4640                .Controls(arr_varCtl(C_CNAM, lngX)).Width = arr_varCtl(C_WDT, lngX)
4650              Case 3
4660                .Controls(arr_varCtl(C_CNAM, lngX)).Left = arr_varCtl(C_LFT, lngX)
4670              End Select
4680            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "cmd" Then
4690              .Controls(arr_varCtl(C_CNAM, lngX)).Left = arr_varCtl(C_LFT, lngX)
4700            End If
4710          End Select
4720        Next  ' ** lngX.
4730      End Select
4740      .Header_box.Width = lngNewForm_Width
4750      .Header_vline01.Left = lngNewForm_Width
4760      .Header_vline02.Left = lngNewForm_Width
4770      .Header_vline03.Left = lngNewForm_Width
4780      .Header_vline04.Left = lngNewForm_Width
4790      .Header_hline01.Width = lngNewForm_Width
4800      .Header_hline02.Width = lngNewForm_Width
4810      .Header_hline03.Width = lngNewForm_Width
4820      .Header_hline04.Width = lngNewForm_Width
4830      .Footer_vline01.Left = lngNewForm_Width
4840      .Footer_vline02.Left = lngNewForm_Width
4850      .Footer_hline01.Width = lngNewForm_Width
4860      .Footer_hline02.Width = lngNewForm_Width
4870      .Width = lngNewForm_Width

4880      .frmMenu_Account_Sub_One.Form.SubResize_One lngWidth_Diff  ' ** Form Procedure: frmMenu_Account_Sub_One.
4890      DoEvents
4900      .frmMenu_Account_Sub_List.Form.SubResize_List lngWidth_Diff  ' ** Form Procedure: frmMenu_Account_Sub_List.
4910      DoEvents

          ' ** Height.
4920      Select Case intModeH
          Case 1
            ' ** Taller.
4930        .Detail.Height = lngNewDetail_Height
4940        For lngX = 0& To (lngCtls - 1&)
4950          Select Case arr_varCtl(C_CTYP, lngX)
              Case acSubform
4960            strTmp01 = Mid(arr_varCtl(C_CNAM, lngX), (CharPos(arr_varCtl(C_CNAM, lngX), 3, "_") + 1))
4970            Select Case strTmp01
                Case "One", "List"
4980              .Controls(arr_varCtl(C_CNAM, lngX)).Height = (arr_varCtl(C_HGT, lngX) + lngHeight_Diff)
4990            Case "Nav"
5000              .Controls(arr_varCtl(C_CNAM, lngX)).Top = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5010            End Select
5020          Case acCommandButton, acBoundObjectFrame
5030            .Controls(arr_varCtl(C_CNAM, lngX)).Top = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5040          Case acLine
5050            If Left(arr_varCtl(C_CNAM, lngX), 8) = "Detail_h" Then
5060              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
5070              Select Case lngTmp02
                  Case 3, 4, 13, 14, 15, 16
                    ' ** Detail_hline03,4,13,14,15,16.
5080                .Controls(arr_varCtl(C_CNAM, lngX)).Top = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5090              End Select
5100            ElseIf Left(arr_varCtl(C_CNAM, lngX), 8) = "Detail_v" Then
5110              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
5120              Select Case lngTmp02
                  Case 3, 4, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22
                    ' ** Detail_vline03,4,13,14,15,16,17,18,19,20,21,22.
5130                .Controls(arr_varCtl(C_CNAM, lngX)).Top = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5140              Case 9, 10, 11, 12
5150                .Controls(arr_varCtl(C_CNAM, lngX)).Height = (arr_varCtl(C_HGT, lngX) + lngHeight_Diff)
5160              End Select
5170            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "cmd" Then
5180              .Controls(arr_varCtl(C_CNAM, lngX)).Top = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5190            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "Nav" Then
5200              .Controls(arr_varCtl(C_CNAM, lngX)).Top = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5210            End If
5220          Case acRectangle, acCheckBox, acLabel, acOptionGroup, acOptionButton
5230            If Left(arr_varCtl(C_CNAM, lngX), 3) = "opg" Then
5240              lngTmp02 = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5250              If lngTmp02 > .Controls(arr_varCtl(C_CNAM, lngX)).Top Then
5260                lngTmp03 = (lngTmp02 - .Controls(arr_varCtl(C_CNAM, lngX)).Top)
5270                lngTmp04 = Int(lngTmp03 / lngTpp)
5280                For lngY = 1 To lngTmp04
5290  On Error Resume Next
5300                  .Controls(arr_varCtl(C_CNAM, lngX)).Top = (.Controls(arr_varCtl(C_CNAM, lngX)).Top + lngTpp)
5310                  If ERR.Number <> 0 Then
5320  On Error GoTo ERRH
5330                    lngTmp03 = ((.Controls(arr_varCtl(C_CNAM, lngX)).Top + lngTpp) + .Controls(arr_varCtl(C_CNAM, lngX)).Height)
5340                    .Detail.Height = lngTmp03
5350                    .Controls(arr_varCtl(C_CNAM, lngX)).Top = (.Controls(arr_varCtl(C_CNAM, lngX)).Top + lngTpp)
5360                  Else
5370  On Error GoTo ERRH
5380                  End If
5390                  If arr_varCtl(C_CTYP, lngX) = acOptionGroup Then
5400                    .Controls(arr_varCtl(C_CNAM, lngX)).Height = arr_varCtl(C_HGT, lngX)  ' ** Make sure this doesn't keep growing!
5410                  End If
5420                Next  ' ** lngY.
5430              ElseIf lngTmp02 < .Controls(arr_varCtl(C_CNAM, lngX)).Top Then
5440                lngTmp03 = (.Controls(arr_varCtl(C_CNAM, lngX)).Top - lngTmp02)
5450                lngTmp04 = Int(lngTmp03 / lngTpp)
5460                For lngY = 1 To lngTmp04
5470                  .Controls(arr_varCtl(C_CNAM, lngX)).Top = (.Controls(arr_varCtl(C_CNAM, lngX)).Top - lngTpp)
5480                  If arr_varCtl(C_CTYP, lngX) = acOptionGroup Then
5490                    .Controls(arr_varCtl(C_CNAM, lngX)).Height = arr_varCtl(C_HGT, lngX)  ' ** Make sure this doesn't keep growing!
5500                  End If
5510                Next  ' ** lngY.
5520              End If
5530            Else
5540              .Controls(arr_varCtl(C_CNAM, lngX)).Top = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5550            End If
5560          Case acTextBox
                ' ** FocusHolder.
5570            .Controls(arr_varCtl(C_CNAM, lngX)).Top = (arr_varCtl(C_TOP, lngX) + lngHeight_Diff)
5580          End Select
5590        Next  ' ** lngX
5600      Case 2
            ' ** Shorter.
            ' ** No shorter than original.
5610      Case 3
            ' ** Original.
5620        .Detail.Height = lngDetail_Height
5630        For lngX = 0& To (lngCtls - 1&)
5640          Select Case arr_varCtl(C_CTYP, lngX)
              Case acSubform
5650            strTmp01 = Mid(arr_varCtl(C_CNAM, lngX), (CharPos(arr_varCtl(C_CNAM, lngX), 3, "_") + 1))
5660            Select Case strTmp01
                Case "One", "List"
5670              .Controls(arr_varCtl(C_CNAM, lngX)).Height = arr_varCtl(C_HGT, lngX)
5680            Case "Nav"
5690              .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
5700            End Select
5710          Case acCommandButton, acBoundObjectFrame
5720            .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
5730          Case acLine
5740            If Left(arr_varCtl(C_CNAM, lngX), 8) = "Detail_h" Then
5750              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
5760              Select Case lngTmp02
                  Case 3, 4, 13, 14, 15, 16
5770                .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
5780              End Select
5790            ElseIf Left(arr_varCtl(C_CNAM, lngX), 8) = "Detail_v" Then
5800              lngTmp02 = Val(Right(arr_varCtl(C_CNAM, lngX), 2))
5810              Select Case lngTmp02
                  Case 3, 4, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22
5820                .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
5830              Case 9, 10, 11, 12
5840                .Controls(arr_varCtl(C_CNAM, lngX)).Height = arr_varCtl(C_HGT, lngX)
5850              End Select
5860            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "cmd" Then
5870              .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
5880            ElseIf Left(arr_varCtl(C_CNAM, lngX), 3) = "Nav" Then
5890              .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
5900            End If
5910          Case acRectangle, acCheckBox, acLabel, acOptionGroup, acOptionButton
5920            If Left(arr_varCtl(C_CNAM, lngX), 3) = "opg" Then
5930              lngTmp02 = arr_varCtl(C_TOP, lngX)
5940              If lngTmp02 > .Controls(arr_varCtl(C_CNAM, lngX)).Top Then
5950                lngTmp03 = (lngTmp02 - .Controls(arr_varCtl(C_CNAM, lngX)).Top)
5960                lngTmp04 = Int(lngTmp03 / lngTpp)
5970                For lngY = 1 To lngTmp04
5980                  .Controls(arr_varCtl(C_CNAM, lngX)).Top = (.Controls(arr_varCtl(C_CNAM, lngX)).Top + lngTpp)
5990                  If arr_varCtl(C_CTYP, lngX) = acOptionGroup Then
6000                    .Controls(arr_varCtl(C_CNAM, lngX)).Height = arr_varCtl(C_HGT, lngX)  ' ** Make sure this doesn't keep growing!
6010                  End If
6020                Next  ' ** lngY.
6030              ElseIf lngTmp02 < .Controls(arr_varCtl(C_CNAM, lngX)).Top Then
6040                lngTmp03 = (.Controls(arr_varCtl(C_CNAM, lngX)).Top - lngTmp02)
6050                lngTmp04 = Int(lngTmp03 / lngTpp)
6060                For lngY = 1 To lngTmp04
6070                  .Controls(arr_varCtl(C_CNAM, lngX)).Top = (.Controls(arr_varCtl(C_CNAM, lngX)).Top - lngTpp)
6080                  If arr_varCtl(C_CTYP, lngX) = acOptionGroup Then
6090                    .Controls(arr_varCtl(C_CNAM, lngX)).Height = arr_varCtl(C_HGT, lngX)  ' ** Make sure this doesn't keep growing!
6100                  End If
6110                Next  ' ** lngY.
6120              End If
6130            Else
6140              .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
6150            End If
6160          Case acTextBox
6170            .Controls(arr_varCtl(C_CNAM, lngX)).Top = arr_varCtl(C_TOP, lngX)
6180          End Select
6190        Next  ' ** lngX.
6200      End Select
6210      .Detail.Height = lngNewDetail_Height

6220    End With

EXITP:
6230    Exit Sub

ERRH:
6240    Select Case ERR.Number
        Case Else
6250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6260    End Select
6270    Resume EXITP

End Sub

Private Sub Form_Timer()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String
        Dim blnSatFocus As Boolean
        Dim varTmp00 As Variant, lngTmp03 As Long, strTmp04 As String, strTmp05 As String

        Static lngTmp01 As Long, lngTmp02 As Long

6310    With Me
6320      .TimerInterval = 0&
6330      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
6340        lngGTR_Stat = lngGTR_Stat + 1&
6350        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
6360        lngTmp01 = lngTmp01 + 1&
6370        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
6380        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
6390        If lngTmp01 > 1& Then
6400          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
6410        End If
6420        If lngTmp01 = 24& Then
6430          lngTmp01 = 0&
6440          lngTmp02 = lngTmp02 + 1&
6450          If lngTmp02 < 5& Then
6460            .TimerInterval = 50&
6470          Else
6480            blnGTR_Emblem = False
6490          End If
6500        Else
6510          .TimerInterval = 50&
6520        End If
6530      End If  ' ** blnGTR_Emblem.
6540      If gblnSetFocus = True Then
6550        gblnSetFocus = False: blnSatFocus = True
6560        strFormName = vbNullString
6570  On Error Resume Next
6580        Do Until Screen.ActiveForm.Name = strFormName
6590          DoCmd.SelectObject acForm, THIS_NAME, False
6600          strFormName = Screen.ActiveForm.Name
6610        Loop
6620  On Error GoTo ERRH
6630        .frmMenu_Account_Sub_Nav.Form.Requery
6640  On Error Resume Next
6650        strTmp04 = .frmMenu_Account_Sub_One.Form.Name
6660  On Error GoTo ERRH
6670        If strTmp04 = "frmMenu_Account_Sub_One" Then
6680          .frmMenu_Account_Sub_One.Form.Requery
6690        End If
6700  On Error Resume Next
6710        strTmp05 = .frmMenu_Account_Sub_List.Form.Name
6720  On Error GoTo ERRH
6730        If strTmp05 = "frmMenu_Account_Sub_List" Then
6740          .frmMenu_Account_Sub_List.Form.Requery
6750        End If
6760        DoEvents
6770        If gstrActNo <> vbNullString Then
6780          .frmMenu_Account_Sub_Nav.SetFocus
6790          .frmMenu_Account_Sub_Nav.Form.MoveRec 0, gstrActNo  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
6800          Select Case .opgSubForm
              Case .opgSubForm_optOne.OptionValue
6810            .frmMenu_Account_Sub_One.SetFocus
6820          Case .opgSubForm_optList.OptionValue
6830            .frmMenu_Account_Sub_List.SetFocus
6840          End Select
6850        Else
              ' ** If this is demo data, start at '11'
6860          varTmp00 = DLookup("[shortname]", "account", "[accountno] = '11'")
6870          If IsNull(varTmp00) = False Then
6880            If varTmp00 = "William B. Johnson Trust" Then
6890              .frmMenu_Account_Sub_Nav.SetFocus
6900              .frmMenu_Account_Sub_Nav.Form.MoveRec 0, "11"  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
6910              Select Case .opgSubForm
                  Case .opgSubForm_optOne.OptionValue
6920                .frmMenu_Account_Sub_One.SetFocus
6930              Case .opgSubForm_optList.OptionValue
6940                .frmMenu_Account_Sub_List.SetFocus
6950              End Select
6960            End If
6970          End If
6980        End If
6990        DoEvents
7000        DoEvents
7010        gstrActNo = vbNullString
            ' ** gstrAccountNo and gstrAccountName get set here.
7020        ChkAcctFields  ' ** Procedure: Below.
7030      End If
7040      If blnGoingToReport = False Then
7050        If blnSatFocus = True Then
7060          blnSatFocus = False
7070          Select Case gstrReturningForm
              Case "frmAccountProfile"
7080            .cmdAccountProfile.SetFocus
7090          Case "frmAccountTransactions"
7100            .cmdTransactions.SetFocus
7110          Case "frmAccountAssets"
7120            .cmdAssets.SetFocus
7130          Case "frmAccountTaxCodes"
7140            .cmdEditTaxCodes.SetFocus
7150          Case "frmAccountIncExpCodes"
7160            .cmdEditIECodes.SetFocus
7170          Case "frmAccountComments"
7180            .cmdEditComments.SetFocus
7190          Case "frmAccountHideTrans2", "frmAccountHideTrans2_One", "frmAccountHideTrans2_Select"
7200            blnFormOpening = False
7210            .cmdHideTransactions.SetFocus
7220          Case "frmAccountProfile_Add"
7230            If .cmdAdd.Enabled = True Then
7240              .cmdAdd.SetFocus
7250            Else
7260              .cmdClose.SetFocus
7270            End If
7280          Case Else
7290            Select Case .opgSubForm
                Case .opgSubForm_optOne.OptionValue
7300              .frmMenu_Account_Sub_One.SetFocus
7310            Case .opgSubForm_optList.OptionValue
7320              .frmMenu_Account_Sub_List.SetFocus
7330            End Select
7340          End Select
7350          gstrReturningForm = vbNullString
7360          DoCmd.Hourglass False
7370        End If
7380      End If
7390      If gblnGoToReport = True Then
7400        DoCmd.Hourglass True  ' ** Make sure it's still running.
7410        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
7420        Else
7430          If blnGTR_Emblem = False Then
7440            blnGTR_Emblem = True
7450            lngTmp01 = 0&: lngTmp02 = 0&
7460          End If
7470          DoEvents
7480          Select Case blnGoingToReport
              Case True
7490            Select Case garr_varGoToReport(GTR_INV1)
                Case True
7500              Select Case garr_varGoToReport(GTR_CTL1)
                  Case "cmdAccountProfile"
                    ' ** If it's Related Accounts, find one!
7510                .TimerInterval = 50&
7520                cmdAccountProfile_Click  ' ** Procedure: Below.
7530              Case "cmdHideTransactions"
7540                .TimerInterval = 50&
7550                cmdHideTransactions_Click  ' ** Procedure: Below.
7560              Case "cmdAdd"
7570                If .cmdAdd.Enabled = True Then
7580                  .TimerInterval = 50&
7590                  cmdAdd_Click  ' ** Procedure: Below.
7600                Else
                      ' ** Just sit there, and we're done.
7610                  blnGoingToReport = False
7620                  DoCmd.Hourglass False
7630                  gblnGoToReport = False
7640                  blnGTR_Emblem = False
7650                  GTREmblem_Off  ' ** Procedure: Below.
7660                  Beep
7670                End If
7680              Case Else
                    ' ** Shouldn't happen.
7690                blnGoingToReport = False
7700                DoCmd.Hourglass False
7710                gblnGoToReport = False
7720                blnGTR_Emblem = False
7730                GTREmblem_Off  ' ** Procedure: Below.
7740                Beep
7750              End Select
7760            Case False
                  ' ** Just sit there, and we're done.
7770              blnGoingToReport = False
7780              DoCmd.Hourglass False
7790              gblnGoToReport = False
7800              blnGTR_Emblem = False
7810              GTREmblem_Off  ' ** Procedure: Below.
7820              Beep
7830            End Select
7840          Case False
7850            If garr_varGoToReport(GTR_FRM1) = THIS_NAME Then
                  ' ** This is first menu in path.
7860              blnGoingToReport = True
7870              Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                  Case True
                    ' ** No more forms.
7880                If .Controls(garr_varGoToReport(GTR_CTL1)).Enabled = True Then
7890                  .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
7900                  lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
7910                  lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
7920                  If lngBtn_Left = .cmdAccountProfile.Left Then
7930                    .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
7940                    .GoToReport_arw_menu_l_img.Visible = True
7950                  ElseIf lngBtn_Left = .cmdEditTaxCodes.Left Then
7960                    .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
7970                    .GoToReport_arw_menu_r_img.Visible = True
7980                  ElseIf lngBtn_Left = .cmdAdd.Left Then
7990                    .GoToReport_arw_add_img.Visible = True
8000                  End If
8010                  DoEvents
8020                  .TimerInterval = 50&
8030                Else
                      ' ** Just sit there, and we're done.
8040                  blnGoingToReport = False
8050                  DoCmd.Hourglass False
8060                  gblnGoToReport = False
8070                  blnGTR_Emblem = False
8080                  GTREmblem_Off  ' ** Procedure: Below.
8090                  Beep
8100                End If
8110              Case False
                    ' ** Moving on.
8120                If .Controls(garr_varGoToReport(GTR_CTL1)).Enabled = True Then
8130                  .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
8140                  lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
8150                  lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
8160                  If lngBtn_Left = .cmdAccountProfile.Left Then
8170                    .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
8180                    .GoToReport_arw_menu_l_img.Visible = True
8190                  ElseIf lngBtn_Left = .cmdEditTaxCodes.Left Then
8200                    .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
8210                    .GoToReport_arw_menu_r_img.Visible = True
8220                  ElseIf lngBtn_Left = .cmdAdd.Left Then
8230                    .GoToReport_arw_add_img.Visible = True
8240                  End If
8250                  Select Case garr_varGoToReport(GTR_INV1)
                      Case True
8260                    DoEvents
8270                    If garr_varGoToReport(GTR_CTL1) = "cmdHideTransactions" Then
8280                      gstrActNo = GetHiddenTrans  ' ** Module Function: modGoToReportFuncs.
8290                      If gstrActNo = vbNullString Then
                            ' ** Account doesn't matter, but make sure they've got transactions.
8300                        lngTmp03 = GetTransCnt  ' ** Module Function: modGoToReportFuncs.
8310                        If lngTmp03 = 0& Then
8320                          .TimerInterval = 0&
8330                          Beep
8340                          DoCmd.Hourglass False
8350                          MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                                "There are no transactions to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
8360                          .GoToReport_arw_menu_r_img.Visible = False
8370                          gblnGoToReport = False
8380                          blnGoingToReport = False
8390                          GTREmblem_Off  ' ** Procedure: Below.
8400                        Else
                              ' ** Just take the first one.
8410                          .frmMenu_Account_Sub_Nav.SetFocus
8420                          .frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
8430                          Select Case .opgSubForm
                              Case .opgSubForm_optOne.OptionValue
8440                            .frmMenu_Account_Sub_One.SetFocus
8450                          Case .opgSubForm_optList.OptionValue
8460                            .frmMenu_Account_Sub_List.SetFocus
8470                          End Select
8480                          DoEvents
8490                          ChkAcctFields  ' ** Procedure: Below.
8500                          .cmdHideTransactions.SetFocus
8510                          .TimerInterval = 50&
8520                        End If
8530                      Else
                            ' ** Use the account with the most hidden transactions.
8540                        .frmMenu_Account_Sub_Nav.SetFocus
8550                        .frmMenu_Account_Sub_Nav.Form.MoveRec 0, gstrActNo  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
8560                        Select Case .opgSubForm
                            Case .opgSubForm_optOne.OptionValue
8570                          .frmMenu_Account_Sub_One.SetFocus
8580                        Case .opgSubForm_optList.OptionValue
8590                          .frmMenu_Account_Sub_List.SetFocus
8600                        End Select
8610                        DoEvents
8620                        ChkAcctFields  ' ** Procedure: Below.
8630                        .cmdHideTransactions.SetFocus
8640                        .TimerInterval = 50&
8650                      End If
8660                    ElseIf garr_varGoToReport(GTR_RNAM) = "rptListOfRelatedAccounts" Then
8670                      gstrActNo = GetRelAcct  ' ** Module Function: modGoToReportFuncs.
8680                      If gstrActNo = vbNullString Then
                            ' ** No matter, but make sure they've got accounts.
8690                        lngTmp03 = GetAcctCnt  ' ** Module Function: modGoToReportFuncs.
8700                        If lngTmp03 = 0& Then  ' ** The 2 system accounts have already been subtracted.
8710                          .TimerInterval = 0&
8720                          Beep
8730                          DoCmd.Hourglass False
8740                          MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                                "There are no accounts to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
8750                          .GoToReport_arw_menu_r_img.Visible = False
8760                          gblnGoToReport = False
8770                          blnGoingToReport = False
8780                          GTREmblem_Off  ' ** Procedure: Below.
8790                        Else
                              ' ** Just take the first one.
8800                          .frmMenu_Account_Sub_Nav.SetFocus
8810                          .frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
8820                          Select Case .opgSubForm
                              Case .opgSubForm_optOne.OptionValue
8830                            .frmMenu_Account_Sub_One.SetFocus
8840                          Case .opgSubForm_optList.OptionValue
8850                            .frmMenu_Account_Sub_List.SetFocus
8860                          End Select
8870                          DoEvents
8880                          ChkAcctFields  ' ** Procedure: Below.
8890                          .cmdAccountProfile.SetFocus
8900                          .TimerInterval = 50&
8910                        End If
8920                      Else
                            ' ** Use an account with the most related accounts.
8930                        .frmMenu_Account_Sub_Nav.SetFocus
8940                        .frmMenu_Account_Sub_Nav.Form.MoveRec 0, gstrActNo  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
8950                        Select Case .opgSubForm
                            Case .opgSubForm_optOne.OptionValue
8960                          .frmMenu_Account_Sub_One.SetFocus
8970                        Case .opgSubForm_optList.OptionValue
8980                          .frmMenu_Account_Sub_List.SetFocus
8990                        End Select
9000                        DoEvents
9010                        ChkAcctFields  ' ** Procedure: Below.
9020                        .cmdAccountProfile.SetFocus
9030                        .TimerInterval = 50&
9040                      End If
9050                    Else
                          ' ** Account Profile, both edit and add.
9060                      .TimerInterval = 50&
9070                    End If
9080                  Case False
                        ' ** Just sit there, and we're done.
9090                    blnGoingToReport = False
9100                    DoCmd.Hourglass False
9110                    gblnGoToReport = False
9120                    blnGTR_Emblem = False
9130                    GTREmblem_Off  ' ** Procedure: Below.
9140                    Beep
9150                  End Select
9160                Else
                      ' ** Just sit there, and we're done.
9170                  blnGoingToReport = False
9180                  DoCmd.Hourglass False
9190                  gblnGoToReport = False
9200                  blnGTR_Emblem = False
9210                  GTREmblem_Off  ' ** Procedure: Below.
9220                  Beep
9230                End If
9240              End Select
9250            Else
                  ' ** Shouldn't be here, then!
9260              blnGoingToReport = False
9270              DoCmd.Hourglass False
9280              gblnGoToReport = False
9290              blnGTR_Emblem = False
9300              GTREmblem_Off  ' ** Procedure: Below.
9310              Beep
9320            End If
9330          End Select
9340        End If
9350      End If
9360      If gblnTimer = True Then
9370        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
9380          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
9390          gblnTimer = False
9400          DoCmd.Hourglass False
9410        End If
9420      End If
9430    End With

EXITP:
9440    Exit Sub

ERRH:
9450    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Ignore.
9460    Case 3420  ' ** Object invalid or no longer set.
          ' ** Carrying on after form is closed.
9470    Case Else
9480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9490    End Select
9500    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngCnt As Long
        Dim lngX As Long

9610    With Me

9620      If gblnSignal = False Then  ' ** Borrowing this variable.
9630        DoCmd.Hourglass True
9640        DoEvents
9650      End If

          ' ** Make sure all reports are closed.
9660      Do While Reports.Count > 0
9670        DoCmd.Close acReport, Reports(0).Name
9680      Loop

9690      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
9700      DoEvents

9710      blnGoingToReport = False

9720      If blnJustClose = False Then

9730        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
9740        DoEvents

9750        AcctNameChk  ' ** Module Function: modStringFuncs.

9760        Set dbs = CurrentDb
            ' ** Empty tmpAccount.
9770        Set qdf = dbs.QueryDefs("qryAccountProfile_01_01")
9780        qdf.Execute
9790        Set qdf = Nothing
9800        dbs.Close
9810        Set dbs = Nothing
9820        DoEvents

            ' ** Make sure all windows are closed.
9830        lngCnt = Forms.Count
9840        If lngCnt > 1 Then
9850          For lngX = (lngCnt - 1&) To 0& Step -1&
9860            If Forms(lngX).Name <> THIS_NAME Then
9870              DoCmd.Close acForm, Forms(lngX).Name
9880            End If
9890          Next
9900        End If
9910        DoEvents

9920        gblnAdmin = False
9930        gstrFormQuerySpec = vbNullString
9940        gblnSetFocus = True
9950        gstrReturningForm = THIS_NAME
9960        Select Case strCallingForm
            Case "frmMenu_Main"
9970          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Title"
9980        End Select

9990      End If  ' ** blnJustClose.

10000   End With

EXITP:
10010   Set qdf = Nothing
10020   Set dbs = Nothing
10030   Exit Sub

ERRH:
10040   DoCmd.Hourglass False
10050   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Ignore.
10060   Case Else
10070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10080   End Select
10090   Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

10110   With Me
10120     If .cmdClose.Caption <> "&Close" Then
10130       .cmdClose.Caption = "&Close"
10140     End If
10150   End With

EXITP:
10160   Exit Sub

ERRH:
10170   Select Case ERR.Number
        Case Else
10180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10190   End Select
10200   Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

10310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10320   Exit Sub

ERRH:
10330   Select Case ERR.Number
        Case Else
10340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10350   End Select
10360   Resume EXITP

End Sub

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

10410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10420   Exit Sub

ERRH:
10430   Select Case ERR.Number
        Case Else
10440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10450   End Select
10460   Resume EXITP

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

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

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

Public Sub cmdClose_Click()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

10710   With Me
10720     If blnFormOpening = False Then

10730       DoCmd.Hourglass True
10740       DoEvents

10750       .TimerInterval = 0&

10760       blnJustClose = False
10770       DoCmd.Close acForm, THIS_NAME

10780     End If
10790   End With

EXITP:
10800   Exit Sub

ERRH:
10810   blnFormOpening = False
10820   DoCmd.Hourglass False
10830   Select Case ERR.Number
        Case Else
10840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10850   End Select
10860   Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10920   intShiftDown = (Shift And acShiftMask) > 0
10930   intAltDown = (Shift And acAltMask) > 0
10940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10960     Select Case intRetVal
          Case vbKeyTab
10970       With Me
10980         intRetVal = 0
10990         Select Case .opgSubForm
              Case .opgSubForm_optOne.OptionValue
11000           .cmdAccountProfile.SetFocus
11010         Case .opgSubForm_optList.OptionValue
11020           .frmMenu_Account_Sub_List.SetFocus
11030           .frmMenu_Account_Sub_List.Form.accountno.SetFocus
11040         End Select
11050       End With
11060     End Select
11070   End If

        ' ** Shift keys.
11080   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11090     Select Case intRetVal
          Case vbKeyTab
11100       With Me
11110         intRetVal = 0
11120         If .cmdDelete.Enabled = True Then
11130           .cmdDelete.SetFocus
11140         ElseIf .cmdAdd.Enabled = True Then
11150           .cmdAdd.SetFocus
11160         ElseIf .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
11170           .chkLoadTimer.SetFocus
11180         Else
11190           .cmdSearch.SetFocus
11200         End If
11210       End With
11220     End Select
11230   End If

EXITP:
11240   KeyCode = intRetVal
11250   Exit Sub

ERRH:
11260   intRetVal = 0
11270   Select Case ERR.Number
        Case Else
11280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11290   End Select
11300   Resume EXITP

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

11410   With Me
11420     .cmdClose.Caption = "< Ba&ck"
11430   End With

EXITP:
11440   Exit Sub

ERRH:
11450   Select Case ERR.Number
        Case Else
11460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11470   End Select
11480   Resume EXITP

End Sub

Public Sub cmdAccountProfile_Click()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountProfile_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strOpenArgs As String
        Dim strDocName As String
        Dim blnRetVal As Boolean

11510   With Me
11520     If blnFormOpening = False Then

11530       .TimerInterval = 0&

11540       If .accountno = vbNullString Then
11550         MsgBox "There must be an existing account to continue.", vbInformation + vbOKOnly, "Entry Required"
11560       Else

11570         strDocName = "frmAccountProfile"

11580         FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
11590         DoEvents

11600         If .chkLoadTimer.Visible = True Then
11610           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
11620           If gblnTimer = True Then
                  'gblnSetFocus = True
11630             DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
11640           End If
11650         End If
11660         DoEvents

11670         DoCmd.Hourglass True
11680         DoEvents

11690         blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.
11700         ChkAcctFields  ' ** Procedure: Below.
11710         DoEvents

11720         Set dbs = CurrentDb
              ' ** Empty tmpAccount.
11730         Set qdf = dbs.QueryDefs("qryAccountProfile_01_01")
11740         qdf.Execute
11750         Set qdf = Nothing
11760         DoEvents

              ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
              ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
              ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
              ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
              ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
              ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
              ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
              ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
11770         Set qdf = dbs.QueryDefs("qryAccountProfile_01_02")
11780         qdf.Execute
11790         Set qdf = Nothing
11800         DoEvents

11810         dbs.Close
11820         Set dbs = Nothing

11830         Select Case .chkIncludeClosed
              Case True
11840           strOpenArgs = "True"
11850         Case False
11860           strOpenArgs = "False"
11870         End Select

11880         gstrAccountNo = .accountno
11890         gstrActNo = .accountno
11900         gblnSetFocus = True
11910         blnJustClose = True
11920         DoCmd.Close acForm, THIS_NAME
11930         DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & strOpenArgs
11940         DoEvents
11950         If gblnGoToReport = True Then
11960           Forms(strDocName).TimerInterval = 100&
11970           blnGoingToReport = False
11980           blnGTR_Emblem = False
11990         End If

12000       End If  ' ** vbNullString.
12010     End If  ' ** blnFormOpening.
12020     DoCmd.Hourglass False
12030   End With

EXITP:
12040   Set qdf = Nothing
12050   Set dbs = Nothing
12060   Exit Sub

ERRH:
12070   blnFormOpening = False
12080   DoCmd.Hourglass False
12090   Select Case ERR.Number
        Case Else
12100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12110   End Select
12120   Resume EXITP

End Sub

Private Sub cmdAccountProfile_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountProfile_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12220   intShiftDown = (Shift And acShiftMask) > 0
12230   intAltDown = (Shift And acAltMask) > 0
12240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12260     Select Case intRetVal
          Case vbKeyTab
12270       With Me
12280         intRetVal = 0
12290         .cmdTransactions.SetFocus
12300       End With
12310     Case vbKeyLeft
12320       With Me
12330         intRetVal = 0
12340         .cmdEditTaxCodes.SetFocus
12350       End With
12360     Case vbKeyRight
12370       With Me
12380         intRetVal = 0
12390         .cmdEditTaxCodes.SetFocus
12400       End With
12410     Case vbKeyUp
12420       With Me
12430         intRetVal = 0
12440         .cmdAssets.SetFocus
12450       End With
12460     Case vbKeyDown
12470       With Me
12480         intRetVal = 0
12490         .cmdTransactions.SetFocus
12500       End With
12510     End Select
12520   End If

        ' ** Shift keys.
12530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12540     Select Case intRetVal
          Case vbKeyTab
12550       With Me
12560         intRetVal = 0
12570         Select Case .opgSubForm
              Case .opgSubForm_optOne.OptionValue
12580           .cmdClose.SetFocus
12590         Case .opgSubForm_optList.OptionValue
12600           .frmMenu_Account_Sub_List.SetFocus  ' ** Leave it on whatever record it was on.
12610           .frmMenu_Account_Sub_List.Form.accountno.SetFocus
12620         End Select
12630       End With
12640     End Select
12650   End If

EXITP:
12660   KeyCode = intRetVal
12670   Exit Sub

ERRH:
12680   intRetVal = 0
12690   Select Case ERR.Number
        Case Else
12700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12710   End Select
12720   Resume EXITP

End Sub

Private Sub cmdTransactions_Click()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactions_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String

12810   With Me
12820     If blnFormOpening = False Then

12830       DoCmd.Hourglass True
12840       DoEvents

12850       .TimerInterval = 0&

12860       ChkAcctFields  ' ** Procedure: Below.

12870       Set dbs = CurrentDb

            ' ** Ledger, by specified [actno].
12880       Set qdf = dbs.QueryDefs("qryAccountMenu_04")
12890       With qdf.Parameters
12900         ![actno] = Me.accountno
12910       End With
12920       Set rst = qdf.OpenRecordset
12930       If rst.BOF = True And rst.EOF = True Then
12940         rst.Close
12950         dbs.Close
12960         DoCmd.Hourglass False
12970         MsgBox "This account has no transactions.", vbInformation + vbOKOnly, "No Transactions"
12980       Else

12990         rst.Close
13000         dbs.Close

13010         gstrAccountNo = .accountno
13020         gstrAccountName = .shortname
13030         gstrActNo = .accountno
13040         gblnSetFocus = True
13050         blnJustClose = True: blnFormOpening = True
13060         strDocName = "frmAccountTransactions"
13070         DoCmd.Close acForm, THIS_NAME
13080         DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & gstrAccountNo & "~" & gstrAccountName

13090         DoCmd.Hourglass False

13100       End If
13110     End If
13120   End With

EXITP:
13130   Set rst = Nothing
13140   Set qdf = Nothing
13150   Set dbs = Nothing
13160   Exit Sub

ERRH:
13170   blnFormOpening = False
13180   DoCmd.Hourglass False
13190   Select Case ERR.Number
        Case Else
13200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13210   End Select
13220   Resume EXITP

End Sub

Private Sub cmdTransactions_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransactions_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13320   intShiftDown = (Shift And acShiftMask) > 0
13330   intAltDown = (Shift And acAltMask) > 0
13340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13360     Select Case intRetVal
          Case vbKeyLeft
13370       With Me
13380         intRetVal = 0
13390         .cmdEditComments.SetFocus
13400       End With
13410     Case vbKeyRight
13420       With Me
13430         intRetVal = 0
13440         .cmdEditComments.SetFocus
13450       End With
13460     Case vbKeyUp
13470       With Me
13480         intRetVal = 0
13490         .cmdAccountProfile.SetFocus
13500       End With
13510     Case vbKeyDown
13520       With Me
13530         intRetVal = 0
13540         .cmdAssets.SetFocus
13550       End With
13560     End Select
13570   End If

EXITP:
13580   KeyCode = intRetVal
13590   Exit Sub

ERRH:
13600   intRetVal = 0
13610   Select Case ERR.Number
        Case Else
13620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13630   End Select
13640   Resume EXITP

End Sub

Private Sub cmdAssets_Click()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssets_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String

13710   With Me
13720     If blnFormOpening = False Then

13730       DoCmd.Hourglass True
13740       DoEvents

13750       .TimerInterval = 0&

13760       ChkAcctFields  ' ** Procedure: Below.

13770       Set dbs = CurrentDb

            ' ** ActiveAssets, by specified [actno].
13780       Set qdf = dbs.QueryDefs("qryAccountMenu_05")
13790       With qdf.Parameters
13800         ![actno] = Me.accountno
13810       End With
13820       Set rst = qdf.OpenRecordset
13830       If rst.BOF = True And rst.EOF = True Then
13840         rst.Close
13850         dbs.Close
13860         DoCmd.Hourglass False
13870         MsgBox "This account has no assets.", vbInformation + vbOKOnly, "No Assets Found"
13880       Else

13890         rst.Close
13900         dbs.Close

13910         gstrAccountNo = .accountno
13920         gstrAccountName = .shortname
13930         gstrActNo = .accountno
13940         gblnSetFocus = True
13950         blnJustClose = True: blnFormOpening = True
13960         strDocName = "frmAccountAssets"
13970         DoCmd.Close acForm, THIS_NAME
13980         DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & gstrAccountNo & "~" & gstrAccountName

13990         DoCmd.Hourglass False

14000       End If
14010     End If
14020   End With

EXITP:
14030   Set rst = Nothing
14040   Set qdf = Nothing
14050   Set dbs = Nothing
14060   Exit Sub

ERRH:
14070   blnFormOpening = False
14080   DoCmd.Hourglass False
14090   Select Case ERR.Number
        Case Else
14100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14110   End Select
14120   Resume EXITP

End Sub

Private Sub cmdAssets_KeyDown(KeyCode As Integer, Shift As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14220   intShiftDown = (Shift And acShiftMask) > 0
14230   intAltDown = (Shift And acAltMask) > 0
14240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14260     Select Case intRetVal
          Case vbKeyLeft
14270       With Me
14280         intRetVal = 0
14290         .cmdHideTransactions.SetFocus
14300       End With
14310     Case vbKeyRight
14320       With Me
14330         intRetVal = 0
14340         .cmdHideTransactions.SetFocus
14350       End With
14360     Case vbKeyUp
14370       With Me
14380         intRetVal = 0
14390         .cmdTransactions.SetFocus
14400       End With
14410     Case vbKeyDown
14420       With Me
14430         intRetVal = 0
14440         .cmdAccountProfile.SetFocus
14450       End With
14460     End Select
14470   End If

EXITP:
14480   KeyCode = intRetVal
14490   Exit Sub

ERRH:
14500   intRetVal = 0
14510   Select Case ERR.Number
        Case Else
14520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14530   End Select
14540   Resume EXITP

End Sub

Private Sub cmdEditTaxCodes_Click()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEditTaxCodes_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String
        Dim lngTmp01 As Long

14610   With Me
14620     If blnFormOpening = False Then

14630       DoCmd.Hourglass True
14640       DoEvents

14650       .TimerInterval = 0&

14660       gstrAccountNo = .accountno
14670       gstrActNo = .accountno

14680       ChkAcctFields  ' ** Procedure: Below.

14690       Set dbs = CurrentDb

            ' ** Ledger, by specified [actno].
14700       Set qdf = dbs.QueryDefs("qryAccountMenu_04")
14710       With qdf.Parameters
14720         ![actno] = Me.accountno
14730       End With
14740       Set rst = qdf.OpenRecordset
14750       If rst.BOF = True And rst.EOF = True Then
14760         rst.Close
14770         Set rst = Nothing
14780         Set qdf = Nothing
14790         dbs.Close
14800         Set dbs = Nothing
14810         DoCmd.Hourglass False
14820         MsgBox "This account has no transactions.", vbInformation + vbOKOnly, "No Transactions"
14830       Else
14840         rst.Close
14850         Set rst = Nothing
14860         Set qdf = Nothing
14870         DoEvents

14880         blnFormOpening = True

              ' ** Empty tmpEdit05.
14890         Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_05")
14900         qdf.Execute
14910         Set qdf = Nothing
14920         DoEvents

              ' ** Empty tmpEdit06.
14930         Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_06")
14940         qdf.Execute
14950         Set qdf = Nothing
14960         DoEvents

              ' ** Append Ledger to tmpEdit05, by specified [actno].  #curr_id
14970         Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_07")
14980         With qdf.Parameters
14990           ![actno] = gstrAccountNo
15000         End With
15010         qdf.Execute
15020         Set qdf = Nothing
15030         DoEvents

              ' ** Append LedgerArchive to tmpEdit05, by specified [actno].  #curr_id
15040         Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_08")
15050         With qdf.Parameters
15060           ![actno] = gstrAccountNo
15070         End With
15080         qdf.Execute
15090         Set qdf = Nothing
15100         DoEvents

              ' ** Append tmpEdit05, with qryAccountProfile_TaxCodes_09 (tblForm_Graphics,
              ' ** just frmAccountTaxCodes_Sub), to tmpEdit06; Cartesian.
15110         Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_10")
15120         qdf.Execute
15130         Set qdf = Nothing
15140         DoEvents

15150         lngTmp01 = 0&

15160         Select Case gblnLinkRevTaxCodes
              Case True
                ' ** qryAccountProfile_TaxCodes_11 (tmpEdit06, for OutOfSync = 1, with taxcode_new,
                ' ** revcode_ID_new, revcode_TYPE_new, revcode_DESC_new), grouped and summed, by OutOfSync.
15170           Set qdf = dbs.QueryDefs("qryAccountProfile_TaxCodes_12")
15180           Set rst = qdf.OpenRecordset
15190           With rst
15200             If .BOF = True And .EOF = True Then
                    ' ** No out-of-sync transactions.
15210             Else
15220               .MoveFirst
15230               lngTmp01 = Nz(![OutOfSync], 0&)
15240             End If
15250             .Close
15260           End With
15270           Set rst = Nothing
15280           Set qdf = Nothing
15290           DoEvents
15300         Case False
15310           lngTmp01 = 0&
15320         End Select

15330         dbs.Close
15340         Set dbs = Nothing

15350         gblnSetFocus = True
15360         blnJustClose = True
15370         strDocName = "frmAccountTaxCodes"
15380         DoCmd.Close acForm, THIS_NAME
15390         DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "^" & CStr(lngTmp01)
15400       End If
15410     End If
15420     DoCmd.Hourglass False
15430   End With

EXITP:
15440   Set rst = Nothing
15450   Set qdf = Nothing
15460   Set dbs = Nothing
15470   Exit Sub

ERRH:
15480   blnFormOpening = False
15490   DoCmd.Hourglass False
15500   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
15510   Case Else
15520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15530   End Select
15540   Resume EXITP

End Sub

Private Sub cmdEditTaxCodes_KeyDown(KeyCode As Integer, Shift As Integer)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEditTaxCodes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15620   intShiftDown = (Shift And acShiftMask) > 0
15630   intAltDown = (Shift And acAltMask) > 0
15640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15660     Select Case intRetVal
          Case vbKeyLeft
15670       With Me
15680         intRetVal = 0
15690         .cmdAccountProfile.SetFocus
15700       End With
15710     Case vbKeyRight
15720       With Me
15730         intRetVal = 0
15740         .cmdAccountProfile.SetFocus
15750       End With
15760     Case vbKeyUp
15770       With Me
15780         intRetVal = 0
15790         .cmdHideTransactions.SetFocus
15800       End With
15810     Case vbKeyDown
15820       With Me
15830         intRetVal = 0
15840         .cmdEditIECodes.SetFocus
15850       End With
15860     End Select
15870   End If

EXITP:
15880   KeyCode = intRetVal
15890   Exit Sub

ERRH:
15900   intRetVal = 0
15910   Select Case ERR.Number
        Case Else
15920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15930   End Select
15940   Resume EXITP

End Sub

Private Sub cmdEditIECodes_Click()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEditIECodes_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String
        Dim lngTmp01 As Long

16010   With Me
16020     If blnFormOpening = False Then

16030       DoCmd.Hourglass True
16040       DoEvents

16050       .TimerInterval = 0&

16060       gstrAccountNo = .accountno
16070       gstrActNo = .accountno

16080       ChkAcctFields  ' ** Procedure: Below.

16090       Set dbs = CurrentDb

            ' ** Ledger, by specified [actno].
16100       Set qdf = dbs.QueryDefs("qryAccountMenu_04")
16110       With qdf.Parameters
16120         ![actno] = gstrAccountNo
16130       End With
16140       Set rst = qdf.OpenRecordset
16150       If rst.BOF = True And rst.EOF = True Then
16160         rst.Close
16170         dbs.Close
16180         Set rst = Nothing
16190         Set qdf = Nothing
16200         Set dbs = Nothing
16210         DoCmd.Hourglass False
16220         MsgBox "This account has no transactions.", vbInformation + vbOKOnly, "No Transactions"
16230       Else
16240         rst.Close
16250         Set rst = Nothing
16260         Set qdf = Nothing
16270         DoEvents

16280         blnFormOpening = True

              ' ** Empty tmpEdit05.
16290         Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_05")
16300         qdf.Execute
16310         Set qdf = Nothing
16320         DoEvents

              ' ** Empty tmpEdit06.
16330         Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_06")
16340         qdf.Execute
16350         Set qdf = Nothing
16360         DoEvents

              ' ** Append Ledger to tmpEdit05, by specified [actno].  #curr_id
16370         Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_07")
16380         With qdf.Parameters
16390           ![actno] = gstrAccountNo
16400         End With
16410         qdf.Execute
16420         Set qdf = Nothing
16430         DoEvents

              ' ** Append LedgerArchive to tmpEdit05, by specified [actno].  #curr_id
16440         Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_08")
16450         With qdf.Parameters
16460           ![actno] = gstrAccountNo
16470         End With
16480         qdf.Execute
16490         Set qdf = Nothing
16500         DoEvents

              ' ** Append tmpEdit05, with qryAccountProfile_IncExpCodes_09 (tblForm_Graphics,
              ' ** just 'frmAccountIncExpCodes_Sub'), to tmpEdit06; Cartesian.
16510         Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_10")
16520         qdf.Execute
16530         Set qdf = Nothing
16540         DoEvents

16550         lngTmp01 = 0&

16560         Select Case gblnLinkRevTaxCodes
              Case True
                ' ** qryAccountProfile_IncExpCodes_11 (tmpEdit06, for OutOfSync = 1, with taxcode_new,
                ' ** revcode_ID_new, revcode_TYPE_new, revcode_DESC_new), grouped and summed, by OutOfSync.
16570           Set qdf = dbs.QueryDefs("qryAccountProfile_IncExpCodes_12")
16580           Set rst = qdf.OpenRecordset
16590           With rst
16600             If .BOF = True And .EOF = True Then
                    ' ** No out-of-sync transactions.
16610             Else
16620               .MoveFirst
16630               lngTmp01 = Nz(![OutOfSync], 0&)
16640             End If
16650             .Close
16660           End With
16670           Set rst = Nothing
16680           Set qdf = Nothing
16690           DoEvents
16700         Case False
16710           lngTmp01 = 0&
16720         End Select

16730         dbs.Close
16740         Set dbs = Nothing

16750         gblnSetFocus = True
16760         blnJustClose = True
16770         strDocName = "frmAccountIncExpCodes"
16780         DoCmd.Close acForm, THIS_NAME
16790         DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "^" & CStr(lngTmp01)

16800       End If
16810     End If
16820   End With

EXITP:
16830   Set rst = Nothing
16840   Set qdf = Nothing
16850   Set dbs = Nothing
16860   Exit Sub

ERRH:
16870   blnFormOpening = False
16880   DoCmd.Hourglass False
16890   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
16900   Case Else
16910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16920   End Select
16930   Resume EXITP

End Sub

Private Sub cmdEditIECodes_KeyDown(KeyCode As Integer, Shift As Integer)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEditIECodes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17020   intShiftDown = (Shift And acShiftMask) > 0
17030   intAltDown = (Shift And acAltMask) > 0
17040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17060     Select Case intRetVal
          Case vbKeyLeft
17070       With Me
17080         intRetVal = 0
17090         .cmdAccountProfile.SetFocus
17100       End With
17110     Case vbKeyRight
17120       With Me
17130         intRetVal = 0
17140         .cmdAccountProfile.SetFocus
17150       End With
17160     Case vbKeyUp
17170       With Me
17180         intRetVal = 0
17190         .cmdEditTaxCodes.SetFocus
17200       End With
17210     Case vbKeyDown
17220       With Me
17230         intRetVal = 0
17240         .cmdEditComments.SetFocus
17250       End With
17260     End Select
17270   End If

EXITP:
17280   KeyCode = intRetVal
17290   Exit Sub

ERRH:
17300   intRetVal = 0
17310   Select Case ERR.Number
        Case Else
17320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17330   End Select
17340   Resume EXITP

End Sub

Private Sub cmdEditComments_Click()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEditComments_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String

17410   With Me
17420     If blnFormOpening = False Then

17430       DoCmd.Hourglass True
17440       DoEvents

17450       .TimerInterval = 0&

17460       gstrAccountNo = .accountno
17470       gstrActNo = .accountno

17480       ChkAcctFields  ' ** Procedure: Below.

17490       Set dbs = CurrentDb

            ' ** Ledger, by specified [actno].
17500       Set qdf = dbs.QueryDefs("qryAccountMenu_04")
17510       With qdf.Parameters
17520         ![actno] = gstrAccountNo
17530       End With
17540       Set rst = qdf.OpenRecordset
17550       If rst.BOF = True And rst.EOF = True Then
17560         rst.Close
17570         dbs.Close
17580         DoCmd.Hourglass False
17590         MsgBox "This account has no transactions.", vbInformation + vbOKOnly, "No transactions"
17600         Set rst = Nothing
17610         Set qdf = Nothing
17620         Set dbs = Nothing
17630       Else
17640         rst.Close
17650         Set rst = Nothing
17660         Set qdf = Nothing

17670         blnFormOpening = True

              ' ** Empty tmpEdit02.
17680         Set qdf = dbs.QueryDefs("qryAccountProfile_Comments_02")
17690         qdf.Execute
17700         Set qdf = Nothing
17710         DoEvents

              ' ** Empty tmpEdit10.
17720         Set qdf = dbs.QueryDefs("qryAccountProfile_Comments_03")
17730         qdf.Execute
17740         Set qdf = Nothing
17750         DoEvents

              ' ** Append Ledger, linked to MasterAsset, to tmpEdit02, by specified GlobalVarGet('gstrAccountNo').  #curr_id
17760         Set qdf = dbs.QueryDefs("qryAccountProfile_Comments_04")
17770         qdf.Execute
17780         Set qdf = Nothing
17790         DoEvents

              ' ** Append LedgerArchive, linked to MasterAsset, to tmpEdit02, by specified GlobalVarGet('gstrAccountNo').  #curr_id
17800         Set qdf = dbs.QueryDefs("qryAccountProfile_Comments_05")
17810         qdf.Execute
17820         Set qdf = Nothing
17830         DoEvents

              ' ** Append tmpEdit02 to tmpEdit10.  #curr_id
17840         Set qdf = dbs.QueryDefs("qryAccountProfile_Comments_06")
17850         qdf.Execute
17860         Set qdf = Nothing
17870         DoEvents

17880         dbs.Close
17890         Set dbs = Nothing

17900         gblnSetFocus = True
17910         blnJustClose = True
17920         strDocName = "frmAccountComments"
17930         DoCmd.Close acForm, THIS_NAME
17940         DoCmd.OpenForm strDocName, , , , , , THIS_NAME

17950       End If
17960     End If
17970   End With

EXITP:
17980   Set rst = Nothing
17990   Set qdf = Nothing
18000   Set dbs = Nothing
18010   Exit Sub

ERRH:
18020   blnFormOpening = False
18030   DoCmd.Hourglass False
18040   Select Case ERR.Number
        Case Else
18050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18060   End Select
18070   Resume EXITP

End Sub

Private Sub cmdEditComments_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEditComments_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18120   intShiftDown = (Shift And acShiftMask) > 0
18130   intAltDown = (Shift And acAltMask) > 0
18140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18160     Select Case intRetVal
          Case vbKeyLeft
18170       With Me
18180         intRetVal = 0
18190         .cmdTransactions.SetFocus
18200       End With
18210     Case vbKeyRight
18220       With Me
18230         intRetVal = 0
18240         .cmdTransactions.SetFocus
18250       End With
18260     Case vbKeyUp
18270       With Me
18280         intRetVal = 0
18290         .cmdEditIECodes.SetFocus
18300       End With
18310     Case vbKeyDown
18320       With Me
18330         intRetVal = 0
18340         .cmdHideTransactions.SetFocus
18350       End With
18360     End Select
18370   End If

EXITP:
18380   KeyCode = intRetVal
18390   Exit Sub

ERRH:
18400   intRetVal = 0
18410   Select Case ERR.Number
        Case Else
18420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18430   End Select
18440   Resume EXITP

End Sub

Private Sub cmdHideTransactions_Click()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHideTransactions_Click"

        Dim strDocName As String
        Dim blnNoTrans As Boolean, blnHasArch As Boolean
        Dim varTmp00 As Variant, varTmp01 As Variant

18510   With Me
18520     If blnEmptyHidden = False Then
18530       If blnFormOpening = False Then

18540         DoCmd.Hourglass True
18550         DoEvents

18560         .TimerInterval = 0&

18570         strDocName = "frmAccountHideTrans2_Select"

18580         gstrAccountNo = .accountno
18590         gstrActNo = .accountno

18600         ChkAcctFields  ' ** Procedure: Below.

18610         If blnEmptyHidden = False Then

18620           blnNoTrans = False: blnHasArch = False
18630           varTmp00 = DCount("*", "ledger", "[accountno] = '" & gstrAccountNo & "'")
18640           varTmp01 = DCount("*", "LedgerArchive", "[accountno] = '" & gstrAccountNo & "'")
18650           Select Case IsNull(varTmp00)
                Case True
18660             Select Case IsNull(varTmp01)
                  Case True
18670               blnNoTrans = True
18680             Case False
18690               If varTmp01 = 0 Then
18700                 blnNoTrans = True
18710               Else
18720                 blnHasArch = True
18730               End If
18740             End Select
18750           Case False
18760             If varTmp00 = 0 Then
18770               Select Case IsNull(varTmp01)
                    Case True
18780                 blnNoTrans = True
18790               Case False
18800                 If varTmp01 = 0 Then
18810                   blnNoTrans = True
18820                 Else
18830                   blnHasArch = True
18840                 End If
18850               End Select
18860             End If
18870           End Select
18880           varTmp00 = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs.
18890           varTmp01 = ZeroIfNull(varTmp01)  ' ** Module Function: modStringFuncs.

18900           Select Case blnNoTrans
                Case True
18910             blnFormOpening = False
18920             If gblnGoToReport = True Then
18930               .TimerInterval = 0&
18940               blnGoingToReport = False
18950               blnGTR_Emblem = False
18960               gblnGoToReport = False
18970               .GoToReport_arw_menu_r_img.Visible = False
18980               GTREmblem_Off  ' ** Procedure: Below.
18990               Beep
19000               DoCmd.Hourglass False
19010               MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There are no transactions to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
19020             Else
19030               Beep
19040               DoCmd.Hourglass False
19050               MsgBox "This account has no transactions.", vbInformation + vbOKOnly, "No Transactions"
19060             End If
19070           Case False
19080             gblnSetFocus = True
19090             strDocName = "frmAccountHideTrans2_Select"
19100             DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & CStr(varTmp00) & "~" & CStr(varTmp01)
19110             If gblnGoToReport = True Then
19120               Forms(strDocName).TimerInterval = 100&
19130               blnGoingToReport = False
                    ' ** blnGTR_Emblem and blnJustClose handled by frmAccountHideTrans2_Select.
19140             End If
19150           End Select

19160         Else
19170           DoCmd.Hourglass False
19180         End If  ' ** blnEmptyHidden.

19190       End If  ' ** blnFormOpening.
19200     Else
19210       blnEmptyHidden = False
19220     End If  ' ** blnEmptyHidden.
19230   End With

EXITP:
19240   Exit Sub

ERRH:
19250   blnFormOpening = False
19260   DoCmd.Hourglass False
19270   Select Case ERR.Number
        Case Else
19280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19290   End Select
19300   Resume EXITP

End Sub

Private Sub cmdHideTransactions_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Easter Egg (Secret Key) to empty the LedgerHidden and tblLedgerHidden tables, so that they can be regenerated.

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHideTransactions_KeyDown"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19420   intShiftDown = (Shift And acShiftMask) > 0
19430   intAltDown = (Shift And acAltMask) > 0
19440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19460     Select Case intRetVal
          Case vbKeyTab
19470       With Me
19480         intRetVal = 0
19490         .cmdSearch.SetFocus
19500       End With
19510     Case vbKeyLeft
19520       With Me
19530         intRetVal = 0
19540         .cmdAssets.SetFocus
19550       End With
19560     Case vbKeyRight
19570       With Me
19580         intRetVal = 0
19590         .cmdAssets.SetFocus
19600       End With
19610     Case vbKeyUp
19620       With Me
19630         intRetVal = 0
19640         .cmdEditComments.SetFocus
19650       End With
19660     Case vbKeyDown
19670       With Me
19680         intRetVal = 0
19690         .cmdEditTaxCodes.SetFocus
19700       End With
19710     End Select
19720   End If

        ' ** Shift keys.
19730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19740     Select Case intRetVal
          Case vbKeyTab
19750       With Me
19760         intRetVal = 0
19770         .cmdEditComments.SetFocus
19780       End With
19790     End Select
19800   End If

        ' ** Ctrl keys.
19810   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
19820     Select Case intRetVal
          Case vbKeyReturn
19830       intRetVal = 0
19840       If blnEmptyHidden = False Then
19850         blnEmptyHidden = True
19860         Beep
19870         If MsgBox("Rebuild the list of of hidden ledger entries?" & vbCrLf & vbCrLf & _
                  "If you've received a 'Group Size Not Accommodated' message," & vbCrLf & _
                  "or other error, when entering the hidden transaction window," & vbCrLf & _
                  "rebuilding this list may solve the problem." & vbCrLf & vbCrLf & _
                  "This will not change the status of any entries that are currently" & vbCrLf & _
                  "hidden, only rebuild a secondary list of specific matches.", _
                  vbQuestion + vbOKCancel, ("Rebuild Hidden Ledger List" & Space(50))) = vbOK Then
19880           DoCmd.Hourglass True
19890           DoEvents
19900           Set dbs = CurrentDb
19910           With dbs
                  ' ** Empty LedgerHidden (in prep for regen).
19920             Set qdf = .QueryDefs("qryAccountHide_11j")
19930             qdf.Execute
19940             Set qdf = Nothing
                  ' ** Empty tblLedgerHidden (in prep for regen).
19950             Set qdf = .QueryDefs("qryAccountHide_11k")
19960             qdf.Execute
19970             Set qdf = Nothing
19980             .Close
19990           End With
20000           Set dbs = Nothing
20010           DoCmd.Hourglass False
20020           MsgBox "Finished." & vbCrLf & "This secondary list will now regenerate the next" & vbCrLf & _
                  "time you click the 'Hide Transactions' button.", vbInformation + vbOKOnly, ("Finished" & Space(40))
                'blnEmptyHidden = False
20030         Else
                'blnEmptyHidden = False
20040         End If
20050       Else
20060         blnEmptyHidden = False
20070       End If  ' ** blnEmptyHidden.
20080     End Select
20090   End If

EXITP:
20100   Set qdf = Nothing
20110   Set dbs = Nothing
20120   KeyCode = intRetVal
20130   Exit Sub

ERRH:
20140   DoCmd.Hourglass False
20150   intRetVal = 0
20160   Select Case ERR.Number
        Case Else
20170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20180   End Select
20190   Resume EXITP

End Sub

Private Sub cmdHideTransactions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHideTransactions_MouseDown"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

20210   If Button = acLeftButton And Shift = acCtrlMask Then
20220     blnEmptyHidden = True
20230     Beep
20240     If MsgBox("Rebuild the list of of hidden ledger entries?" & vbCrLf & vbCrLf & _
              "If you've received a 'Group Size Not Accommodated' message," & vbCrLf & _
              "or other error, when entering the hidden transaction window," & vbCrLf & _
              "rebuilding this list may solve the problem." & vbCrLf & vbCrLf & _
              "This will not change the status of any entries that are currently" & vbCrLf & _
              "hidden, only rebuild a secondary list of specific matches.", _
              vbQuestion + vbOKCancel, ("Rebuild Hidden Ledger List" & Space(50))) = vbOK Then
20250       DoCmd.Hourglass True
20260       DoEvents
20270       Set dbs = CurrentDb
20280       With dbs
              ' ** Empty LedgerHidden (in prep for regen).
20290         Set qdf = .QueryDefs("qryAccountHide_11j")
20300         qdf.Execute
20310         Set qdf = Nothing
              ' ** Empty tblLedgerHidden (in prep for regen).
20320         Set qdf = .QueryDefs("qryAccountHide_11k")
20330         qdf.Execute
20340         Set qdf = Nothing
20350         .Close
20360       End With
20370       Set dbs = Nothing
20380       DoCmd.Hourglass False
20390       MsgBox "Finished." & vbCrLf & "This secondary list will now regenerate the next" & vbCrLf & _
              "time you click the 'Hide Transactions' button.", vbInformation + vbOKOnly, ("Finished" & Space(40))
            'blnEmptyHidden = False
20400     Else
            'blnEmptyHidden = False
20410     End If
20420   End If

EXITP:
20430   Set qdf = Nothing
20440   Set dbs = Nothing
20450   Exit Sub

ERRH:
20460   Select Case ERR.Number
        Case Else
20470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20480   End Select
20490   Resume EXITP

End Sub

Private Sub opgSubform_AfterUpdate()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgSubForm_AfterUpdate"

        Dim strAccountNo As String
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

20510   With Me
20520     strAccountNo = .accountno
20530     Select Case .opgSubForm
          Case .opgSubForm_optOne.OptionValue
20540       .opgSubForm_optOne_lbl.FontBold = True
20550       .opgSubForm_optList_lbl.FontBold = False
20560       .frmMenu_Account_Sub_One.Visible = True
20570       .frmMenu_Account_Sub_List.Visible = False
20580       .Detail_vline09.Visible = False
20590       .Detail_vline10.Visible = False
20600       .Detail_vline11.Visible = False
20610       .Detail_vline12.Visible = False
20620       blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.
20630       .frmMenu_Account_Sub_Nav.SetFocus
20640       .frmMenu_Account_Sub_Nav.Form.MoveRec 0, strAccountNo  ' ** Form Procedure: frmMenu_Account_Sub_One.
20650       .frmMenu_Account_Sub_One.SetFocus
20660     Case .opgSubForm_optList.OptionValue
20670       .opgSubForm_optOne_lbl.FontBold = False
20680       .opgSubForm_optList_lbl.FontBold = True
20690       .frmMenu_Account_Sub_One.Visible = False
20700       .frmMenu_Account_Sub_List.Visible = True
20710       .Detail_vline09.Visible = True
20720       .Detail_vline10.Visible = True
20730       .Detail_vline11.Visible = True
20740       .Detail_vline12.Visible = True
20750       blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
20760       .frmMenu_Account_Sub_Nav.SetFocus
20770       .frmMenu_Account_Sub_Nav.Form.MoveRec 0, strAccountNo  ' ** Form Procedure: frmMenu_Account_Sub_List.
20780       .frmMenu_Account_Sub_List.SetFocus
20790 On Error Resume Next
20800       strTmp01 = .frmMenu_Account_Sub_List.Form.Name
20810 On Error GoTo ERRH
20820       If strTmp01 = "frmMenu_Account_Sub_List" Then
20830         .frmMenu_Account_Sub_List.Form.accountno.SetFocus
20840       End If
20850     End Select
20860   End With

EXITP:
20870   Exit Sub

ERRH:
20880   Select Case ERR.Number
        Case Else
20890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20900   End Select
20910   Resume EXITP

End Sub

Private Sub opgSubForm_optOne_KeyDown(KeyCode As Integer, Shift As Integer)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgSubForm_optOne_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21020   intShiftDown = (Shift And acShiftMask) > 0
21030   intAltDown = (Shift And acAltMask) > 0
21040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21070       With Me
21080         intRetVal = 0
21090         .chkIncludeClosed.SetFocus
21100       End With
21110     End Select
21120   End If

        ' ** Shift keys.
21130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21150       With Me
21160         intRetVal = 0
21170         .cmdHideTransactions.SetFocus
21180       End With
21190     End Select
21200   End If

EXITP:
21210   KeyCode = intRetVal
21220   Exit Sub

ERRH:
21230   intRetVal = 0
21240   Select Case ERR.Number
        Case Else
21250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21260   End Select
21270   Resume EXITP

End Sub

Private Sub opgSubForm_optList_KeyDown(KeyCode As Integer, Shift As Integer)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgSubForm_optList_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21320   intShiftDown = (Shift And acShiftMask) > 0
21330   intAltDown = (Shift And acAltMask) > 0
21340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21370       With Me
21380         intRetVal = 0
21390         .chkIncludeClosed.SetFocus
21400       End With
21410     End Select
21420   End If

        ' ** Shift keys.
21430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21450       With Me
21460         intRetVal = 0
21470         .cmdHideTransactions.SetFocus
21480       End With
21490     End Select
21500   End If

EXITP:
21510   KeyCode = intRetVal
21520   Exit Sub

ERRH:
21530   intRetVal = 0
21540   Select Case ERR.Number
        Case Else
21550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21560   End Select
21570   Resume EXITP

End Sub

Private Sub chkIncludeClosed_AfterUpdate()

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_AfterUpdate"

        Dim strFilter As String, strRecSource As String
        Dim strTmp01 As String

21610   With Me
21620     Select Case .chkIncludeClosed
          Case True
21630       .chkIncludeClosed_lbl.FontBold = True
21640       .frmMenu_Account_Sub_Nav.SetFocus
21650       .frmMenu_Account_Sub_Nav.Form.Filter = vbNullString
21660       .frmMenu_Account_Sub_Nav.Form.FilterOn = False
21670       If .frmMenu_Account_Sub_One.Visible = True Then
21680         .frmMenu_Account_Sub_One.SetFocus
21690       End If
21700       .frmMenu_Account_Sub_One.Form.Filter = vbNullString
21710       .frmMenu_Account_Sub_One.Form.FilterOn = False
21720       .frmMenu_Account_Sub_One.Form.dateclosed.Visible = True
21730       strRecSource = .frmMenu_Account_Sub_One.Form.RecordSource
21740       .frmMenu_Account_Sub_One.Form.RecordSource = strRecSource
21750       If .frmMenu_Account_Sub_List.Visible = True Then
21760         .frmMenu_Account_Sub_List.SetFocus
21770       End If
21780       .frmMenu_Account_Sub_List.Form.Filter = vbNullString
21790       .frmMenu_Account_Sub_List.Form.FilterOn = False
21800       .frmMenu_Account_Sub_List.Form.dateclosed.Visible = True
21810       .frmMenu_Account_Sub_List.Form.dateclosed_lbl.Visible = True
21820       .frmMenu_Account_Sub_List.Form.dateclosed_lbl_line.Visible = True
21830       strRecSource = .frmMenu_Account_Sub_List.Form.RecordSource
21840       .frmMenu_Account_Sub_List.Form.RecordSource = strRecSource
21850     Case False
21860       .chkIncludeClosed_lbl.FontBold = False
21870       strFilter = "[dateclosed] Is Null"
21880       .frmMenu_Account_Sub_Nav.SetFocus
21890       .frmMenu_Account_Sub_Nav.Form.Filter = strFilter
21900       .frmMenu_Account_Sub_Nav.Form.FilterOn = True
21910       If .frmMenu_Account_Sub_One.Visible = True Then
21920         .frmMenu_Account_Sub_One.SetFocus
21930       End If
21940       .frmMenu_Account_Sub_One.Form.Filter = strFilter
21950       .frmMenu_Account_Sub_One.Form.FilterOn = True
21960       .frmMenu_Account_Sub_One.Form.dateclosed.Visible = False
21970       If .frmMenu_Account_Sub_List.Visible = True Then
21980         .frmMenu_Account_Sub_List.SetFocus
21990       End If
22000 On Error Resume Next
22010       strTmp01 = .frmMenu_Account_Sub_List.Form.Name
22020 On Error GoTo ERRH
22030       If strTmp01 = "frmMenu_Account_Sub_List" Then
22040         .frmMenu_Account_Sub_List.Form.Filter = strFilter
22050         .frmMenu_Account_Sub_List.Form.FilterOn = True
22060         .frmMenu_Account_Sub_List.Form.dateclosed.Visible = False
22070         .frmMenu_Account_Sub_List.Form.dateclosed_lbl.Visible = False
22080         .frmMenu_Account_Sub_List.Form.dateclosed_lbl_line.Visible = False
22090       End If
22100     End Select
22110   End With

EXITP:
22120   Exit Sub

ERRH:
22130   Select Case ERR.Number
        Case Else
22140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22150   End Select
22160   Resume EXITP

End Sub

Private Sub chkIncludeClosed_KeyDown(KeyCode As Integer, Shift As Integer)

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeClosed_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22220   intShiftDown = (Shift And acShiftMask) > 0
22230   intAltDown = (Shift And acAltMask) > 0
22240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22270       With Me
22280         intRetVal = 0
22290         .cmdSearch.SetFocus
22300       End With
22310     End Select
22320   End If

        ' ** Shift keys.
22330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22350       With Me
22360         intRetVal = 0
22370         .opgSubForm.SetFocus
22380       End With
22390     End Select
22400   End If

EXITP:
22410   KeyCode = intRetVal
22420   Exit Sub

ERRH:
22430   intRetVal = 0
22440   Select Case ERR.Number
        Case Else
22450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22460   End Select
22470   Resume EXITP

End Sub

Private Sub cmdSearch_Click()

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSearch_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

22510   With Me
22520     If blnFormOpening = False Then

22530       .TimerInterval = 0&

22540       blnContinue = True

22550       blnFormOpening = True
22560       DoCmd.OpenForm "frmAccountSearch", , , , , acDialog, THIS_NAME
22570       blnFormOpening = False

22580       If gstrAccountNo <> vbNullString Then

22590         If Len(gstrAccountNo) > 15 Then
22600           blnContinue = False
22610           MsgBox "The Account Number must be <= 15 characters.", vbExclamation + vbOKOnly, "Invalid Account Number"
22620         Else

22630           Set dbs = CurrentDb
                ' ** Account, by specified [actno].
22640           Set qdf = dbs.QueryDefs("qryAccountMenu_02")
22650           With qdf.Parameters
22660             ![actno] = gstrAccountNo
22670           End With
22680           Set rst = qdf.OpenRecordset
22690           If rst.EOF Then
                  ' ** Account doesn't exist.
22700             blnContinue = False
22710             MsgBox "Not a valid account number.", vbExclamation + vbOKOnly, "Account Not Found"
22720           Else
22730             With Me
22740               If IsNull(rst![dateclosed]) = False And .chkIncludeClosed = False Then
22750                 Beep
22760                 msgResponse = MsgBox("The account specified is closed, and you" & vbCrLf & _
                        "have chosen not to view closed accounts." & vbCrLf & vbCrLf & _
                        "Do you wish to view this account anyway?", vbQuestion + vbOKCancel, "Account Is Closed")
22770                 If msgResponse <> vbOK Then
22780                   blnContinue = False
22790                 Else
22800                   .chkIncludeClosed = True
22810                   chkIncludeClosed_AfterUpdate  ' ** Procedure: Below.
22820                   DoEvents
22830                 End If
22840               End If
22850               If blnContinue = True Then
22860                 Select Case .opgSubForm
                      Case .opgSubForm_optOne.OptionValue
22870                   .frmMenu_Account_Sub_One.SetFocus
22880                   .frmMenu_Account_Sub_One.Form.accountno.SetFocus
22890                   If gstrAccountNo <> vbNullString Then
22900                     .frmMenu_Account_Sub_Nav.Form.MoveRec 0, gstrAccountNo  ' ** Form Procedure: frmMenu_Account_Sub_One.
22910                   End If
22920                 Case .opgSubForm_optList.OptionValue
22930                   .frmMenu_Account_Sub_List.SetFocus
22940                   .frmMenu_Account_Sub_List.Form.accountno.SetFocus
22950                   If gstrAccountNo <> vbNullString Then
22960                     .frmMenu_Account_Sub_List.Form.MoveRec 0, gstrAccountNo  ' ** Form Procedure: frmMenu_Account_Sub_List.
22970                   End If
22980                 End Select
22990               End If
23000             End With
23010           End If
23020           rst.Close
23030           dbs.Close

23040         End If
23050       End If

23060     End If
23070   End With

EXITP:
23080   Set rst = Nothing
23090   Set qdf = Nothing
23100   Set dbs = Nothing
23110   Exit Sub

ERRH:
23120   blnFormOpening = False
23130   DoCmd.Hourglass False
23140   Select Case ERR.Number
        Case Else
23150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23160   End Select
23170   Resume EXITP

End Sub

Private Sub cmdSearch_KeyDown(KeyCode As Integer, Shift As Integer)

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSearch_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23220   intShiftDown = (Shift And acShiftMask) > 0
23230   intAltDown = (Shift And acAltMask) > 0
23240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23260     Select Case intRetVal
          Case vbKeyTab
23270       With Me
23280         intRetVal = 0
23290         If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
23300           .chkLoadTimer.SetFocus
23310         ElseIf .cmdAdd.Enabled = True Then
23320           .cmdAdd.SetFocus
23330         ElseIf .cmdDelete.Enabled = True Then
23340           .cmdDelete.SetFocus
23350         Else
23360           .cmdClose.SetFocus
23370         End If
23380       End With
23390     End Select
23400   End If

        ' ** Shift keys.
23410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23420     Select Case intRetVal
          Case vbKeyTab
23430       With Me
23440         intRetVal = 0
23450         .cmdHideTransactions.SetFocus
23460       End With
23470     End Select
23480   End If

EXITP:
23490   KeyCode = intRetVal
23500   Exit Sub

ERRH:
23510   intRetVal = 0
23520   Select Case ERR.Number
        Case Else
23530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23540   End Select
23550   Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

23610   With Me
23620     If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
23630       gblnTimer = False
23640     End If
23650     Select Case .chkLoadTimer
          Case True
23660       .chkLoadTimer_lbl.FontBold = True
23670       If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
23680         gblnTimer = True
23690       End If
23700     Case False
23710       .chkLoadTimer_lbl.FontBold = False
23720     End Select
23730   End With

EXITP:
23740   Exit Sub

ERRH:
23750   Select Case ERR.Number
        Case Else
23760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23770   End Select
23780   Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23820   intShiftDown = (Shift And acShiftMask) > 0
23830   intAltDown = (Shift And acAltMask) > 0
23840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23860     Select Case intRetVal
          Case vbKeyTab
23870       With Me
23880         intRetVal = 0
23890         If .cmdAdd.Enabled = True Then
23900           .cmdAdd.SetFocus
23910         ElseIf .cmdDelete.Enabled = True Then
23920           .cmdDelete.SetFocus
23930         Else
23940           .cmdClose.SetFocus
23950         End If
23960       End With
23970     End Select
23980   End If

        ' ** Shift keys.
23990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24000     Select Case intRetVal
          Case vbKeyTab
24010       With Me
24020         intRetVal = 0
24030         .cmdSearch.SetFocus
24040       End With
24050     End Select
24060   End If

EXITP:
24070   KeyCode = intRetVal
24080   Exit Sub

ERRH:
24090   intRetVal = 0
24100   Select Case ERR.Number
        Case Else
24110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24120   End Select
24130   Resume EXITP

End Sub

Private Sub cmdAdd_Click()

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

24210   With Me
24220     If blnFormOpening = False Then

24230       DoCmd.Hourglass True
24240       DoEvents

24250       .TimerInterval = 0&

24260       Set dbs = CurrentDb
            ' ** Empty tmpAccount.
24270       Set qdf = dbs.QueryDefs("qryAccountProfile_15")
24280       qdf.Execute
24290       Set qdf = Nothing
24300       DoEvents
24310       dbs.Close
24320       Set dbs = Nothing
24330       DoEvents

24340       gstrAccountNo = vbNullString
24350       gblnSetFocus = True
24360       blnJustClose = True: blnFormOpening = True
24370       strDocName = "frmAccountProfile_Add"
24380       DoCmd.Close acForm, THIS_NAME
24390       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
24400       If gblnGoToReport = True Then
24410         Forms(strDocName).TimerInterval = 50&
24420         blnGoingToReport = False
24430         blnGTR_Emblem = False
24440       End If

24450     End If
24460   End With

EXITP:
24470   Set qdf = Nothing
24480   Set dbs = Nothing
24490   Exit Sub

ERRH:
24500   blnFormOpening = False
24510   DoCmd.Hourglass False
24520   Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
24530   Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
24540   Case Else
24550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24560   End Select
24570   Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24620   intShiftDown = (Shift And acShiftMask) > 0
24630   intAltDown = (Shift And acAltMask) > 0
24640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24660     Select Case intRetVal
          Case vbKeyTab
24670       With Me
24680         intRetVal = 0
24690         If .cmdDelete.Enabled = True Then
24700           .cmdDelete.SetFocus
24710         Else
24720           .cmdClose.SetFocus
24730         End If
24740       End With
24750     End Select
24760   End If

        ' ** Shift keys.
24770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24780     Select Case intRetVal
          Case vbKeyTab
24790       With Me
24800         intRetVal = 0
24810         If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
24820           .chkLoadTimer.SetFocus
24830         Else
24840           .cmdSearch.SetFocus
24850         End If
24860       End With
24870     End Select
24880   End If

EXITP:
24890   KeyCode = intRetVal
24900   Exit Sub

ERRH:
24910   intRetVal = 0
24920   Select Case ERR.Number
        Case Else
24930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24940   End Select
24950   Resume EXITP

End Sub

Private Sub cmdDelete_Click()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst As DAO.Recordset
        Dim frm As Access.Form, frm2 As Access.Form, frm3 As Access.Form
        Dim strAccountNo As String
        Dim dblICash As Double, dblPCash As Double, lngActiveAssets As Long
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long
        Dim strTmp01 As String, lngTmp02 As Long
        Dim lngX As Long
        Dim blnContinue As Boolean

25010   With Me
25020     If blnFormOpening = False Then

25030       DoCmd.Hourglass True
25040       DoEvents

25050       .TimerInterval = 0&

25060       blnContinue = True

25070       Select Case .opgSubForm
            Case .opgSubForm_optOne.OptionValue
25080         Set frm = .frmMenu_Account_Sub_One.Form
25090         Set frm2 = .frmMenu_Account_Sub_List.Form
25100         Set frm3 = .frmMenu_Account_Sub_Nav.Form
25110       Case .opgSubForm_optList.OptionValue
25120         Set frm = .frmMenu_Account_Sub_List.Form
25130         Set frm2 = .frmMenu_Account_Sub_One.Form
25140         Set frm3 = .frmMenu_Account_Sub_Nav.Form
25150       End Select

25160       strAccountNo = frm.accountno

25170       If strAccountNo = "INCOME O/U" Or strAccountNo = "99-INCOME O/U" Then
25180         blnContinue = False
25190         DoCmd.Hourglass False
25200         Beep
25210         MsgBox "You cannot delete this account.", vbInformation + vbOKOnly, "Invalid Action"
25220       Else
25230         If strAccountNo = "SUSPENSE" Or strAccountNo = "99-SUSPENSE" Then
25240           blnContinue = False
25250           DoCmd.Hourglass False
25260           Beep
25270           MsgBox "You cannot delete this account.", vbInformation + vbOKOnly, "Invalid Action"
25280         Else
25290           If IsNull(frm.ICash) = True Then dblICash = 0# Else dblICash = frm.ICash
25300           If IsNull(frm.PCash) = True Then dblPCash = 0# Else dblPCash = frm.PCash
25310           If IsNull(frm.ActiveAssets) = True Then lngActiveAssets = 0& Else lngActiveAssets = frm.ActiveAssets
25320           If (dblICash + dblPCash) <> 0 Or lngActiveAssets > 0 Then
25330             DoCmd.Hourglass False
25340             Beep
25350             If MsgBox("Cash and/or Assets are held by this account." & vbCrLf & _
                      "If you choose to delete it, all the data related to it will be deleted." & vbCrLf & vbCrLf & _
                      "Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Account") <> vbYes Then
25360               blnContinue = False
25370             End If
25380           End If
25390         End If
25400       End If
25410       DoEvents

25420       If blnContinue = True Then

25430         If IsNull(frm.shortname) = True Then
25440           strTmp01 = "{not named}"
25450         Else
25460           If frm.shortname = vbNullString Then
25470             strTmp01 = "{not named}"
25480           Else
25490             strTmp01 = frm.shortname
25500           End If
25510         End If
25520         DoCmd.Hourglass False
25530         Beep
25540         msgResponse = MsgBox("Are you sure you want to delete '" & strAccountNo & " - " & strTmp01 & "'?", _
                vbQuestion + vbYesNo, "Delete Account")

25550         If msgResponse = vbYes Then

25560           DoCmd.Hourglass True
25570           DoEvents

25580           blnFormOpening = True

                ' ** Check for Related Accounts.
25590           If IsNull(frm.related_accountno) = False Then
                  ' ** Simple comma-separated list, e.g., '11,20,45'.
                  ' ** There should always be at least 2, the related account, and the one you were on when you chose it.
                  ' ** If there are only 2, empty the other one.
                  ' ** If 3 or more, each must have this one removed from its list.
25600             strTmp01 = Trim(frm.related_accountno)
25610             If InStr(strTmp01, ",") = 0 Then
                    ' ** Who knows, but we don't care.
25620               strTmp01 = vbNullString
25630             End If
25640             If strTmp01 <> vbNullString Then
25650               Set dbs = CurrentDb
25660               With dbs
                      ' ** Account, all non-Null related_accountno.
25670                 Set qdf1 = .QueryDefs("qryAccountProfile_09")
25680                 Set rst = qdf1.OpenRecordset
25690                 With rst
25700                   If .BOF = True And .EOF = True Then
                          ' ** Well, THIS account has one!
25710                   Else
25720                     .MoveLast
25730                     lngTmp02 = .RecordCount
25740                     .MoveFirst
25750                     For lngX = 1& To (lngTmp02 - 1&)
25760                       If ![accountno] <> strAccountNo Then
25770                         If InStr(![related_accountno], strAccountNo) > 0 Then
25780                           If Left(![related_accountno], (Len(strAccountNo) + 1)) = (strAccountNo & ",") Then
                                  ' ** This one's the first in the list.
25790                             strTmp01 = Mid(![related_accountno], (InStr(![related_accountno], ",") + 1))
25800                             strTmp01 = Trim(strTmp01)
25810                             If InStr(strTmp01, ",") = 0 Then
                                    ' ** Only these two, so Null it out.
25820                               .Edit
25830                               ![related_accountno] = Null
25840                               .Update
25850                             Else
                                    ' ** Leave the remaining ones.
25860                               .Edit
25870                               ![related_accountno] = strTmp01
25880                               .Update
25890                             End If
25900                           ElseIf Right(![related_accountno], (Len(strAccountNo) + 1)) = ("," & strAccountNo) Then
                                  ' ** This one's the last in the list.
25910                             strTmp01 = Left(![related_accountno], (Len(![related_accountno]) - Len("," & strAccountNo)))
25920                             strTmp01 = Trim(strTmp01)
25930                             If InStr(strTmp01, ",") = 0 Then
                                    ' ** Only these two, so Null it out.
25940                               .Edit
25950                               ![related_accountno] = Null
25960                               .Update
25970                             Else
                                    ' ** Leave the remaining ones.
25980                               .Edit
25990                               ![related_accountno] = strTmp01
26000                               .Update
26010                             End If
26020                           Else
26030                             If InStr(![related_accountno], "," & strAccountNo & ",") > 0 Then
                                    ' ** 1st half includes the comma, 2nd half doesn't.
26040                               strTmp01 = Left(![related_accountno], InStr(![related_accountno], "," & strAccountNo & ",")) & _
                                      Mid(![related_accountno], InStr(![related_accountno], "," & _
                                      strAccountNo & ",") + Len("," & strAccountNo & ","))
26050                               strTmp01 = Trim(strTmp01)
                                    ' ** Leave the remaining ones.
26060                               .Edit
26070                               ![related_accountno] = strTmp01
26080                               .Update
26090                             Else
                                    ' ** It's only a partial match, so move on.
26100                             End If
26110                           End If
26120                         End If
26130                       End If
26140                       If lngX < lngTmp02 Then .MoveNext
26150                     Next
26160                   End If
26170                   .Close
26180                 End With
26190                 .Close
26200               End With
26210             End If
26220           End If
26230           DoEvents

26240           Set dbs = CurrentDb

                ' ** Subtract all currently held assets from MasterAsset.
                'strSQL = "SELECT ActiveAssets.accountno, ActiveAssets.assetno, Sum(ActiveAssets.shareface) AS SumOfshareface "
                'strSQL = strSQL & "FROM ActiveAssets "
                'strSQL = strSQL & "WHERE ActiveAssets.accountno ='" & strAccountNo & "' "
                'strSQL = strSQL & "GROUP BY ActiveAssets.accountno, ActiveAssets.assetno;"

                ' ** ActiveAssets, grouped and summed, by accountno, assetno, by specified [actno].
26250           Set qdf1 = dbs.QueryDefs("qryAccountProfile_Delete_02")
26260           With qdf1.Parameters
26270             ![actno] = strAccountNo
26280           End With
26290           Set rst = qdf1.OpenRecordset
26300           With rst
26310             If .BOF = True And .EOF = True Then
                    ' ** No Active Assets.
26320             Else
26330               .MoveLast
26340               lngRecs = .RecordCount
26350               .MoveFirst
26360               For lngX = 1& To lngRecs
                      'strSQL = "UPDATE masterasset SET masterasset.shareface = masterasset.shareface - " & rst![SumOfshareface]
                      'strSQL = strSQL & " WHERE ((masterasset.assetno)= " & rst![assetno] & ");"
                      ' ** Update qryAccountProfile_Delete_03 (MasterAsset, with shareface_new, by specified [astno], [sharfac]).
26370                 Set qdf2 = dbs.QueryDefs("qryAccountProfile_Delete_04")
26380                 With qdf2.Parameters
26390                   ![astno] = rst![assetno]
26400                   ![sharfac] = rst![shareface]
26410                 End With
26420                 qdf2.Execute dbFailOnError
26430                 Set qdf2 = Nothing
26440                 If lngX < lngRecs Then .MoveNext
26450               Next
26460             End If
26470             .Close
26480           End With
26490           Set rst = Nothing
26500           Set qdf1 = Nothing
26510           DoEvents

                ' *****************************************************************************************************
                ' ** NOTE: THIS WILL CASCADE DELETE LEDGER, JOURNAL, BALANCE, AND ALL OTHER TABLES LINKED TO ACCOUNT!
                ' *****************************************************************************************************

                ' ** Cascade deletes these tables:
                ' **   ActiveAssets
                ' **   Asset
                ' **   Balance
                ' **   Journal
                ' **   Ledger
                ' **   LedgerHidden
                ' **   PortfolioModel
                ' **   tblCheckBank
                ' **   tblCurrency_Account
                ' **   tblLedgerHidden

                ' ** These are NOT deleted:
                ' **   LedgerArchive
                ' **   tblCheckPOSPay_Detail
                ' **   tblCheckReconcile_Account
                ' **   tblCheckVoid
                ' **   tblPortfolioModeling

                ' ** Delete tblCheckVoid, by specified [actno].
26520           Set qdf1 = dbs.QueryDefs("qryAccountProfile_Delete_07")
26530           With qdf1.Parameters
26540             ![actno] = strAccountNo
26550           End With
26560           qdf1.Execute
26570           Set qdf1 = Nothing
26580           DoEvents

                ' ** Delete Account, by specified [actno].
26590           Set qdf1 = dbs.QueryDefs("qryAccountProfile_Delete_01")
26600           With qdf1.Parameters
26610             ![actno] = strAccountNo
26620           End With
26630           qdf1.Execute dbFailOnError
26640           Set qdf1 = Nothing
26650           DoEvents

                ' ** Delete LedgerArchive, by specified [actno].
26660           Set qdf1 = dbs.QueryDefs("qryAccountProfile_Delete_05")
26670           With qdf1.Parameters
26680             ![actno] = strAccountNo
26690           End With
26700           qdf1.Execute
26710           Set qdf1 = Nothing
26720           DoEvents

                ' ** Delete tmpForm_Graphics_Account, by specified [actno].
26730           Set qdf1 = dbs.QueryDefs("qryAccountProfile_01e")
26740           With qdf1.Parameters
26750             ![actno] = strAccountNo
26760           End With
26770           qdf1.Execute
26780           Set qdf1 = Nothing
26790           DoEvents

                ' ** Delete tblCurrency_Account, by specified [actno].  'SHOULD BE HANDLED ABOVE!
26800           Set qdf1 = dbs.QueryDefs("qryAccountProfile_Delete_06")
26810           With qdf1.Parameters
26820             ![actno] = strAccountNo
26830           End With
26840           qdf1.Execute
26850           Set qdf1 = Nothing
26860           DoEvents

                ' ** Delete tblCheckPOSPay_Detail, by specified [actno].
26870           Set qdf1 = dbs.QueryDefs("qryAccountProfile_Delete_08")
26880           With qdf1.Parameters
26890             ![actno] = strAccountNo
26900           End With
26910           qdf1.Execute
26920           Set qdf1 = Nothing
26930           DoEvents

                ' ** Delete tblCheckReconcile_Account, by specified [actno].
26940           Set qdf1 = dbs.QueryDefs("qryAccountProfile_Delete_09")
26950           With qdf1.Parameters
26960             ![actno] = strAccountNo
26970           End With
26980           qdf1.Execute
26990           Set qdf1 = Nothing
27000           DoEvents

                ' ** Delete tblPortfolioModeling, by specified [actno].
27010           Set qdf1 = dbs.QueryDefs("qryAccountProfile_Delete_10")
27020           With qdf1.Parameters
27030             ![actno] = strAccountNo
27040           End With
27050           qdf1.Execute
27060           Set qdf1 = Nothing
27070           DoEvents

27080           dbs.Close
27090           Set dbs = Nothing

27100           DoEvents
27110           frm.Requery
27120           frm2.Requery
27130           frm3.Requery

27140           ChkAccountLimit  ' ** Procedure: Below.
27150           DoEvents

27160         End If  ' ** msgResponse.

27170         Select Case .opgSubForm
              Case .opgSubForm_optOne.OptionValue
27180           .cmdAccountProfile.SetFocus
27190         Case .opgSubForm_optList.OptionValue
27200           .frmMenu_Account_Sub_List.SetFocus
27210           .frmMenu_Account_Sub_List.Form.accountno.SetFocus
27220         End Select

27230       End If  ' ** blnContinue.

27240       blnFormOpening = False

27250       DoCmd.Hourglass False

27260     End If
27270   End With  ' ** Me.

EXITP:
27280   Set frm = Nothing
27290   Set frm2 = Nothing
27300   Set frm3 = Nothing
27310   Set rst = Nothing
27320   Set qdf1 = Nothing
27330   Set qdf2 = Nothing
27340   Set dbs = Nothing
27350   Exit Sub

ERRH:
27360   blnFormOpening = False
27370   DoCmd.Hourglass False
27380   Select Case ERR.Number
        Case Else
27390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27400   End Select
27410   Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27520   intShiftDown = (Shift And acShiftMask) > 0
27530   intAltDown = (Shift And acAltMask) > 0
27540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27560     Select Case intRetVal
          Case vbKeyTab
27570       With Me
27580         intRetVal = 0
27590         .cmdClose.SetFocus
27600       End With
27610     End Select
27620   End If

        ' ** Shift keys.
27630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27640     Select Case intRetVal
          Case vbKeyTab
27650       With Me
27660         intRetVal = 0
27670         If .cmdAdd.Enabled = True Then
27680           .cmdAdd.SetFocus
27690         ElseIf .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
27700           .chkLoadTimer.SetFocus
27710         Else
27720           .cmdSearch.SetFocus
27730         End If
27740       End With
27750     End Select
27760   End If

EXITP:
27770   KeyCode = intRetVal
27780   Exit Sub

ERRH:
27790   intRetVal = 0
27800   Select Case ERR.Number
        Case Else
27810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27820   End Select
27830   Resume EXITP

End Sub

Public Sub ChkAcctFields()

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkAcctFields"

        Dim strTmp01 As String, strTmp02 As String

27910   With Me
27920 On Error Resume Next
27930     strTmp01 = .frmMenu_Account_Sub_One.Form.Name
27940 On Error GoTo ERRH
27950 On Error Resume Next
27960     strTmp02 = .frmMenu_Account_Sub_List.Form.Name
27970 On Error GoTo ERRH
27980     Select Case .opgSubForm
          Case .opgSubForm_optOne.OptionValue
27990       If strTmp01 = "frmMenu_Account_Sub_One" Then
28000         If IsNull(.frmMenu_Account_Sub_One.Form.accountno) = False Then
28010           If .frmMenu_Account_Sub_One.Form.accountno <> vbNullString Then
28020             If IsNull(.accountno) = True Then
28030               .accountno = .frmMenu_Account_Sub_One.Form.accountno
28040               .dateclosed = .frmMenu_Account_Sub_One.Form.dateclosed
28050               .shortname = .frmMenu_Account_Sub_One.Form.shortname
28060               .legalname = .frmMenu_Account_Sub_One.Form.legalname
28070               .hidden_trans = .frmMenu_Account_Sub_One.Form.hidden_trans
28080             End If
28090           End If
28100         End If
28110       End If
28120     Case .opgSubForm_optList.OptionValue
28130       If strTmp02 = "frmMenu_Account_Sub_List" Then
28140         If IsNull(.frmMenu_Account_Sub_List.Form.accountno) = False Then
28150           If .frmMenu_Account_Sub_List.Form.accountno <> vbNullString Then
28160             If IsNull(.accountno) = True Then
28170               .accountno = .frmMenu_Account_Sub_List.Form.accountno
28180               .dateclosed = .frmMenu_Account_Sub_List.Form.dateclosed
28190               .shortname = .frmMenu_Account_Sub_List.Form.shortname
28200               .legalname = .frmMenu_Account_Sub_List.Form.legalname
28210               .hidden_trans = .frmMenu_Account_Sub_List.Form.hidden_trans
28220             End If
28230           End If
28240         End If
28250       End If
28260     End Select
28270     gstrAccountNo = .accountno
28280     gstrAccountName = .shortname
28290   End With

EXITP:
28300   Exit Sub

ERRH:
28310   DoCmd.Hourglass False
28320   Select Case ERR.Number
        Case Else
28330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28340   End Select
28350   Resume EXITP

End Sub

Private Sub ChkAccountLimit()

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkAccountLimit"

        Dim lngCurrentAccounts As Long

28410   With Me

28420     lngCurrentAccounts = DCount("*", "account")

28430     If gstrTrustDataLocation = vbNullString Then
28440       IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
28450     End If

28460     If Right(gstrTrustDataLocation, 1) <> LNK_SEP Then gstrTrustDataLocation = gstrTrustDataLocation & LNK_SEP

28470     If lngMaxAccounts = 0& Then
28480       lngMaxAccounts = DecodeString(IniFile_Get("License", "Limit", "", gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modStartupFuncs.
28490     End If

28500     If lngCurrentAccounts < lngMaxAccounts Then
28510       .cmdAdd.Enabled = True
28520     Else
28530       .cmdAdd.Enabled = False
28540     End If

28550   End With

EXITP:
28560   Exit Sub

ERRH:
28570   Select Case ERR.Number
        Case Else
28580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28590   End Select
28600   Resume EXITP

End Sub

Public Sub JustClose_Set(blnCloseIt As Boolean)

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "JustClose_Set"

28710   blnJustClose = blnCloseIt

EXITP:
28720   Exit Sub

ERRH:
28730   Select Case ERR.Number
        Case Else
28740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28750   End Select
28760   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

28810   With Me
28820     blnGTR_Emblem = False
28830     For lngX = 1& To 24&
28840 On Error Resume Next
28850       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
28860       If ERR.Number <> 0 Then
28870 On Error GoTo ERRH
28880         Exit For
28890       Else
28900 On Error GoTo ERRH
28910       End If
28920     Next
28930   End With

EXITP:
28940   Exit Sub

ERRH:
28950   Select Case ERR.Number
        Case Else
28960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28970   End Select
28980   Resume EXITP

End Sub

Public Sub FormCloseVar(blnLock As Boolean)

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "FormCloseVar"

29010   blnFormOpening = blnLock

EXITP:
29020   Exit Sub

ERRH:
29030   Select Case ERR.Number
        Case Else
29040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29050   End Select
29060   Resume EXITP

End Sub

Private Sub MoveToZero()

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveToZero"

29110   With Me
29120     .form_width_line.Left = 0&
29130     .detail_height_line.Top = 0&
29140     .GoToReport_arw_menu_l_img_lbl.Top = 0&
29150     .GoToReport_arw_menu_l_img_lbl.Left = 0&
29160     .GoToReport_arw_menu_r_img_lbl.Top = 0&
29170     .GoToReport_arw_menu_r_img_lbl.Left = 0&
29180     .GoToReport_arw_add_img_lbl.Left = 0&
29190     .GoToReport_Emblem_lbl.Left = 0&
29200     .HideTrans_Asterisk_img.Top = 0&
29210     .HideTrans_Asterisk_img.Left = 0&
29220     .HideTrans_Asterisk_img_lbl.Top = 0&
29230     .HideTrans_Asterisk_img_lbl.Left = 0&
29240     .cmdHideTransactions_lbl2.Top = 0&
29250     .cmdHideTransactions_lbl2.Left = 0&
29260     .cmdHideTransactions_lbl3.Top = 0&
29270     .cmdHideTransactions_lbl3.Left = 0&
29280   End With

EXITP:
29290   Exit Sub

ERRH:
29300   Select Case ERR.Number
        Case Else
29310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29320   End Select
29330   Resume EXITP

End Sub

Private Sub CtlArray_Load()

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "CtlArray_Load"

        Dim ctl As Access.Control
        Dim blnFound As Boolean
        Dim lngE As Long

29410   With Me
29420     If lngCtls = 0& Or IsEmpty(arr_varCtl) = True Then
29430       For Each ctl In .Detail.Controls
29440         blnFound = False
29450         With ctl
29460           Select Case .ControlType
                Case acCommandButton
29470             If Left(.Name, 3) = "cmd" Then
29480               blnFound = True
29490             End If
29500           Case acSubform, acRectangle, acLine, acCheckBox, acOptionGroup, acOptionButton
29510             blnFound = True
29520           Case acBoundObjectFrame
29530             If Left(.Name, 10) = "GoToReport" Then
29540               blnFound = True
29550             End If
29560           Case acLabel
29570             If Left(.Name, 3) = "opg" Or Left(.Name, 3) = "chk" Then
29580               blnFound = True
29590             ElseIf .Name = "cmdHideTransactions_lbl" Or .Name = "FormClose_lbl1" Or .Name = "ShortcutMenu_lbl" Then
29600               blnFound = True
29610             End If
29620           Case acTextBox
29630             If .Name = "FocusHolder" Then
29640               blnFound = True
29650             End If
29660           End Select
29670           If blnFound = True Then
29680             lngCtls = lngCtls + 1&
29690             lngE = lngCtls - 1&
29700             ReDim Preserve arr_varCtl(C_ELEMS, lngE)
29710             arr_varCtl(C_CNAM, lngE) = .Name
29720             arr_varCtl(C_CTYP, lngE) = .ControlType
29730             arr_varCtl(C_TOP, lngE) = .Top
29740             arr_varCtl(C_LFT, lngE) = .Left
29750             arr_varCtl(C_WDT, lngE) = .Width
29760             arr_varCtl(C_HGT, lngE) = .Height
29770           End If
29780         End With
29790       Next
29800     End If
29810   End With

EXITP:
29820   Set ctl = Nothing
29830   Exit Sub

ERRH:
29840   Select Case ERR.Number
        Case Else
29850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29860   End Select
29870   Resume EXITP

End Sub
