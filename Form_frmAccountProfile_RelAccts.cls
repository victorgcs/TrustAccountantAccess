VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile_RelAccts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile_RelAccts"

'VGC 09/02/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Accounts:         A {frmAccountProfile_RelAccts_Sub}
' **   Sort By Number:   B {opgOrderBy_optNumber}
' **   Cancel:           C {cmdCancel}
' **   Clear Closed:     E {cmdClearClosed}
' **   Sort By Name:     M {opgOrderBy_optName}
' **   Select None:      N {cmdSelectNone}
' **   OK:               O {cmdOK}
' **   Print Report:     P {cmdPrintReport}
' **   Reset:            R {cmdSelectReset}
' **   Sort By Selected: S {opgOrderBy_optSelected}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAccountProfile_RelAccts_Sub}

' ********************************************************************************
' ** NOTE: I believe my checking Related Accounts as frmAccountProfile opens
' ** has rendered the ClosedAccts section on the form moot. It'll never show up!
' ********************************************************************************

'Forms(0).lbxAccounts.Top = 450
'Forms(0).lbxAccounts.Left = 255
'Forms(0).lbxAccounts.Width = 3180
'Forms(0).lbxAccounts.Height = 3390

Private lngRAs As Long, arr_varRA() As Variant    ' ** Complete, full list.
Private lngCRAs As Long, arr_varCRA() As Variant  ' ** Closed accounts still present.
Private lngORAs As Long, arr_varORA() As Variant  ' ** Original list on entry.
Private Const RA_ELEMS As Integer = 3  ' ** Array's first-element UBound()
Private Const RA_ACCT     As Integer = 0
Private Const RA_FOUND    As Integer = 1
Private Const RA_CLOSED   As Integer = 2
Private Const RA_LST_ELEM As Integer = 3

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean
Private blnSelectReset_Focus As Boolean, blnSelectReset_MouseDown As Boolean
Private blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean
Private blnClearClosed_Focus As Boolean, blnClearClosed_MouseDown As Boolean

Private blnRptVis As Boolean, blnGoingToReport As Boolean, blnIsOpen As Boolean
Private blnOrderBy1 As Boolean, blnOrderBy2 As Boolean, strQryLastAppend As String
Private strCallingForm As String, strCallingFormSub As String, strRelAccts As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRels As Long, arr_varRel As Variant
        Dim blnNotSaved As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngY As Long, lngE As Long

110     blnNotSaved = False

120     lngRAs = 0&
130     ReDim arr_varRA(RA_ELEMS, 0)
140     lngCRAs = 0&
150     ReDim arr_varCRA(RA_ELEMS, 0)
160     lngORAs = 0&
170     ReDim arr_varORA(RA_ELEMS, 0)

180     With Me
190       If IsNull(.OpenArgs) = False Then
200         If Trim(.OpenArgs) <> vbNullString Then
210           strCallingForm = Trim(.OpenArgs)
220           intPos01 = InStr(strCallingForm, "~")
230           If intPos01 > 0 Then

240             If gblnGoToReport = True And .TimerInterval = 0& Then
250               .TimerInterval = 100&
260             End If

270             If intPos01 = Len(strCallingForm) Then
280               strRelAccts = vbNullString
290               strCallingForm = Left(strCallingForm, (intPos01 - 1))
300             Else
310               strRelAccts = Mid(strCallingForm, (intPos01 + 1))
320               strCallingForm = Left(strCallingForm, (intPos01 - 1))
330             End If

340             Select Case strCallingForm
                Case "frmAccountProfile"
350               strCallingFormSub = "frmAccountProfile_Sub"
360             Case "frmAccountProfile_Add"
370               strCallingFormSub = "frmAccountProfile_Add_Sub"
380             End Select

390             strQryLastAppend = vbNullString
400             blnOrderBy1 = False: blnOrderBy2 = False: blnRptVis = False

410             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

420             blnSelectReset_Focus = False: blnSelectReset_MouseDown = False
430             blnSelectNone_Focus = False: blnSelectNone_MouseDown = False
440             blnClearClosed_Focus = False: blnClearClosed_MouseDown = False
450             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

460             .cmdClearClosed.Enabled = False
470             .cmdClearClosed.Visible = False
480             .cmdClearClosed_raised_img.Visible = False
490             .cmdClearClosed_raised_semifocus_dots_img.Visible = False
500             .cmdClearClosed_raised_focus_img.Visible = False
510             .cmdClearClosed_raised_focus_dots_img.Visible = False
520             .cmdClearClosed_sunken_focus_dots_img.Visible = False
530             .cmdClearClosed_raised_img_dis.Visible = False

540             If strRelAccts <> vbNullString Then

550               intPos01 = InStr(strRelAccts, ",")
560               If intPos01 > 0 Then
570                 strTmp01 = strRelAccts
580                 strTmp02 = vbNullString
590                 Do While intPos01 > 0
600                   strTmp02 = Left(strTmp01, (intPos01 - 1))  ' ** This 1 account number.
610                   strTmp01 = Trim(Mid(strTmp01, (intPos01 + 1)))  ' ** Remaining related accounts.
620                   lngRAs = lngRAs + 1&
630                   lngE = lngRAs - 1&
640                   ReDim Preserve arr_varRA(RA_ELEMS, lngE)
650                   arr_varRA(RA_ACCT, lngE) = strTmp02
660                   arr_varRA(RA_FOUND, lngE) = CBool(False)
670                   arr_varRA(RA_CLOSED, lngE) = CBool(False)
680                   arr_varRA(RA_LST_ELEM, lngE) = -1&
690                   intPos01 = InStr(strTmp01, ",")
700                 Loop
710                 lngRAs = lngRAs + 1&
720                 lngE = lngRAs - 1&
730                 ReDim Preserve arr_varRA(RA_ELEMS, lngE)
740                 arr_varRA(RA_ACCT, lngE) = strTmp01
750                 arr_varRA(RA_FOUND, lngE) = CBool(False)
760                 arr_varRA(RA_CLOSED, lngE) = CBool(False)
770                 arr_varRA(RA_LST_ELEM, lngE) = -1&
780               Else

                    ' ** Check whether it's got spaces, which was how earlier
                    ' ** versions of Trust Accountant were to separate related accounts.
790                 intPos01 = InStr(strRelAccts, " ")
800                 If intPos01 > 0 Then
810                   strTmp01 = strRelAccts
820                   strTmp02 = vbNullString
830                   Do While intPos01 > 0
840                     strTmp02 = Left(strTmp01, (intPos01 - 1))  ' ** This 1 account number.
850                     strTmp01 = Trim(Mid(strTmp01, (intPos01 + 1)))  ' ** Remaining related accounts.
860                     lngRAs = lngRAs + 1&
870                     lngE = lngRAs - 1&
880                     ReDim Preserve arr_varRA(RA_ELEMS, lngE)
890                     arr_varRA(RA_ACCT, lngE) = strTmp02
900                     arr_varRA(RA_FOUND, lngE) = CBool(False)
910                     arr_varRA(RA_CLOSED, lngE) = CBool(False)
920                     arr_varRA(RA_LST_ELEM, lngE) = -1&
930                     intPos01 = InStr(strTmp01, " ")
940                   Loop
950                   lngRAs = lngRAs + 1&
960                   lngE = lngRAs - 1&
970                   ReDim Preserve arr_varRA(RA_ELEMS, lngE)
980                   arr_varRA(RA_ACCT, lngE) = strTmp01
990                   arr_varRA(RA_FOUND, lngE) = CBool(False)
1000                  arr_varRA(RA_CLOSED, lngE) = CBool(False)
1010                  arr_varRA(RA_LST_ELEM, lngE) = -1&
1020                Else
                      ' ** Related accounts should always include this one,
                      ' ** and at least one other.
1030                  If strRelAccts <> Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno Then
                        ' ** Add this account to the list.
1040                    lngRAs = 2&
1050                    lngE = lngRAs - 1&
1060                    ReDim arr_varRA(RA_ELEMS, lngE)
1070                    arr_varRA(RA_ACCT, 0) = Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno
1080                    arr_varRA(RA_FOUND, 0) = CBool(False)
1090                    arr_varRA(RA_CLOSED, 0) = CBool(False)
1100                    arr_varRA(RA_LST_ELEM, 0) = -1&
1110                    arr_varRA(RA_ACCT, lngE) = strRelAccts
1120                    arr_varRA(RA_FOUND, lngE) = CBool(False)
1130                    arr_varRA(RA_CLOSED, lngE) = CBool(False)
1140                    arr_varRA(RA_LST_ELEM, lngE) = -1&
1150                  Else
                        ' ** If it's only got this account, it needs at least 1 more.
1160                    lngRAs = 1&
1170                    lngE = lngRAs - 1&
1180                    ReDim arr_varRA(RA_ELEMS, lngE)
1190                    arr_varRA(RA_ACCT, lngE) = strRelAccts
1200                    arr_varRA(RA_FOUND, lngE) = CBool(False)
1210                    arr_varRA(RA_CLOSED, lngE) = CBool(False)
1220                    arr_varRA(RA_LST_ELEM, lngE) = -1&
1230                  End If
1240                End If

1250              End If

1260            Else
                  ' ** None are listed, so begin with this one alone.
1270              lngRAs = 1&
1280              lngE = lngRAs - 1&
1290              ReDim arr_varRA(RA_ELEMS, lngE)
1300              arr_varRA(RA_ACCT, lngE) = Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno
1310              arr_varRA(RA_FOUND, lngE) = CBool(False)
1320              arr_varRA(RA_CLOSED, lngE) = CBool(False)
1330              arr_varRA(RA_LST_ELEM, lngE) = -1&
1340            End If
1350            DoEvents

                ' **************************************************
                ' ** Array: arr_varRA()
                ' **
                ' **   Field  Element  Name          Constant
                ' **   =====  =======  ============  =============
                ' **     1       0     accountno     RA_ACCT
                ' **     2       1     Found         RA_FOUND
                ' **     3       2     Closed        RA_CLOSED
                ' **     4       3     relacct_id    RA_LST_ELEM
                ' **
                ' **************************************************

                'THIS CAN NEVER BE GREATER THAN ZERO!
                'WHAT'S IT DOING HERE?
1360            If lngCRAs > 0& Then
1370              .frmAccountProfile_RelAccts_Sub.Form.CRelArray_Set arr_varCRA, lngCRAs  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
1380            End If
1390            DoEvents

1400            blnIsOpen = True

                ' ** Highlight the related accounts.
1410            .frmAccountProfile_RelAccts_Sub.Form.HighlightRelated arr_varRA, lngRAs, False, blnIsOpen  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
1420            DoEvents

1430            arr_varRel = .frmAccountProfile_RelAccts_Sub.Form.RelArray_Get  ' ** Form Function: frmAccountProfile_RelAccts_Sub.
1440            lngRels = (UBound(arr_varRel, 2) + 1&)
1450            For lngX = 0& To (lngRels - 1&)
1460              For lngY = 0& To (lngRAs - 1&)
1470                If arr_varRA(RA_ACCT, lngY) = arr_varRel(RA_ACCT, lngX) Then
1480                  arr_varRA(RA_FOUND, lngY) = arr_varRel(RA_FOUND, lngX)
1490                  arr_varRA(RA_LST_ELEM, lngY) = arr_varRel(RA_LST_ELEM, lngX)
1500                  Exit For
1510                End If
1520              Next
1530            Next

1540            For lngX = 0& To (lngRAs - 1&)
1550              DoEvents

1560              If arr_varRA(RA_FOUND, lngX) = False Then
                    ' ** Account not found. Closed?
1570                If arr_varRA(RA_ACCT, lngX) = Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno Then
1580                  blnNotSaved = True
1590                  Cancel = -1
1600                  MsgBox "Please save the new account before choosing Related Accounts.", vbInformation + vbOKOnly, "Account Must Be Saved"
1610                  Exit For
1620                Else
1630                  Set dbs = CurrentDb
1640                  With dbs
1650                    Set rst = .OpenRecordset("account", dbOpenDynaset, dbReadOnly)
1660                    With rst
1670                      .FindFirst "[accountno] = '" & arr_varRA(RA_ACCT, lngX) & "'"
1680                      If .BOF = True And .EOF = True Then
                            ' ** Account not found. Notify user, and remove from list.
1690                        Beep
1700                        MsgBox "Account Number " & arr_varRA(RA_ACCT, lngX) & " was not found in the current " & vbCrLf & _
                              "account list and will be removed." & vbCrLf & vbCrLf & _
                              "If you believe this to be in error, choose it again.", vbInformation + vbOKOnly, "Account Not Found"
1710                      Else
1720                        If IsNull(![dateclosed]) = False Then
1730                          arr_varRA(RA_FOUND, lngX) = True
1740                          arr_varRA(RA_CLOSED, lngX) = True
1750                          lngCRAs = lngCRAs + 1&
1760                          lngE = lngCRAs - 1&
1770                          ReDim Preserve arr_varCRA(RA_ELEMS, lngE)
1780                          arr_varCRA(RA_ACCT, lngE) = arr_varRA(RA_ACCT, lngX)
1790                          arr_varCRA(RA_FOUND, lngE) = arr_varRA(RA_FOUND, lngX)
1800                          arr_varCRA(RA_CLOSED, lngE) = arr_varRA(RA_CLOSED, lngX)
1810                          arr_varCRA(RA_LST_ELEM, lngE) = arr_varRA(RA_LST_ELEM, lngX)
1820                        Else
                              ' ** Then it should be in the list box!
                              ' ** Remove it and let them choose it again!
1830                          Beep
1840                          MsgBox "Account Number " & arr_varRA(RA_ACCT, lngX) & " was not found in the current " & vbCrLf & _
                                "account list and will be removed." & vbCrLf & vbCrLf & _
                                "If you believe this to be in error, choose it again.", vbInformation + vbOKOnly, "Account Not Found"
1850                        End If
1860                      End If
1870                      .Close
1880                    End With  ' ** rst.
1890                    Set rst = Nothing
1900                    .Close
1910                  End With  ' ** dbs.
1920                  Set dbs = Nothing
1930                End If
1940              End If

1950            Next  ' ** lngRAs: lngX.
1960            DoEvents

1970            If Cancel = 0 Then

1980              If lngCRAs > 0 Then
1990                .frmAccountProfile_RelAccts_Sub.Form.CRelArray_Set arr_varCRA, lngCRAs  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
2000              End If
2010              DoEvents

                  ' ** Put the current list into the related_accounts field here.
2020              If lngRAs > 0& Then
2030                strTmp01 = vbNullString: strTmp02 = vbNullString
2040                For lngX = 0& To (lngRAs - 1&)
2050                  If arr_varRA(RA_FOUND, lngX) = False And arr_varRA(RA_CLOSED, lngX) = False Then
                        ' ** Account not found; skip it.
2060                  Else
2070                    lngORAs = lngORAs + 1&
2080                    lngE = lngORAs - 1&
2090                    ReDim Preserve arr_varORA(RA_ELEMS, lngE)
2100                    arr_varORA(RA_ACCT, lngE) = arr_varRA(RA_ACCT, lngE)
2110                    arr_varORA(RA_FOUND, lngE) = arr_varRA(RA_FOUND, lngE)
2120                    arr_varORA(RA_CLOSED, lngE) = arr_varRA(RA_CLOSED, lngE)
2130                    arr_varORA(RA_LST_ELEM, lngE) = arr_varRA(RA_LST_ELEM, lngE)
2140                    If arr_varRA(RA_CLOSED, lngX) = True Then
2150                      strTmp02 = strTmp02 & arr_varRA(RA_ACCT, lngX) & ", "
2160                    End If
2170                    strTmp01 = strTmp01 & arr_varRA(RA_ACCT, lngX) & ","
2180                  End If
2190                Next
2200                strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))  ' ** Remove last comma.
2210                If strTmp02 <> vbNullString Then
2220                  strTmp02 = Trim(strTmp02)
2230                  strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
2240                  .ClosedAccts.Visible = True
2250                  .ClosedAccts_lbl2.Visible = True
2260                  .ClosedAccts = strTmp02
2270                  .cmdClearClosed.Visible = True
2280                  .cmdClearClosed.Enabled = True
2290                  .cmdClearClosed_raised_img.Visible = True
2300                End If
2310                .related_accountno = strTmp01
2320              End If

2330              .frmAccountProfile_RelAccts_Sub.Form.RelArray_Set arr_varRA, lngRAs  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
2340              DoEvents
2350              .frmAccountProfile_RelAccts_Sub.Form.ORelArray_Set arr_varORA, lngORAs  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
2360              DoEvents

                  '.frmAccountProfile_RelAccts_Sub.SetFocus
                  'DoEvents

2370              blnOrderBy1 = True
2380              .TimerInterval = 100&

2390            End If  ' ** Cancel.

                'If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
                '  .ClosedAccts.Visible = True
                '  .ClosedAccts_lbl2.Visible = True
                '  .ClosedAccts = "12,13,14,15"
                '  .cmdClearClosed.Visible = True
                '  .cmdClearClosed.Enabled = True
                '  .cmdClearClosed_raised_img.Visible = True
                'End If

2400          Else
2410            Cancel = -1
2420          End If
2430        Else
2440          Cancel = -1
2450        End If
2460      Else
2470        Cancel = -1
2480      End If
2490    End With

2500    If Cancel = -1 And blnNotSaved = False Then  ' ** It's already shown a message.
2510      Beep
2520      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
2530    End If

EXITP:
2540    Set rst = Nothing
2550    Set dbs = Nothing
2560    Exit Sub

ERRH:
2570    Select Case ERR.Number
        Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAccountProfile_RelAccts_Sub}

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyEscape
2770        intRetVal = 0
2780        cmdCancel_Click  ' ** Procedure: Below.
2790      End Select
2800    End If

        ' ** Ctrl keys.
2810    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2820      Select Case intRetVal
          Case vbKeyS
2830        With Me
2840          intRetVal = 0
2850          .frmAccountProfile_RelAccts_Sub.SetFocus
2860          .frmAccountProfile_RelAccts_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
2870        End With
2880      End Select
2890    End If

        ' ** Ctrl-Shift keys.
2900    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2910      Select Case intRetVal
          Case vbKeyF
2920        With Me
2930          intRetVal = 0
2940          .FocusHolder.SetFocus
2950        End With
2960      End Select
2970    End If

EXITP:
2980    KeyCode = intRetVal
2990    Exit Sub

ERRH:
3000    intRetVal = 0
3010    Select Case ERR.Number
        Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub

Private Sub Form_Timer()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

3110    With Me
3120      .TimerInterval = 0&
3130      If blnOrderBy1 = True Then
3140        blnOrderBy1 = False
3150        opgOrderBy_AfterUpdate  ' ** Procedure: Below.
3160        DoEvents
3170      ElseIf blnOrderBy2 = True Then
3180        blnOrderBy2 = False
3190        .frmAccountProfile_RelAccts_Sub.Form.HighlightRelated arr_varRA, lngRAs, False, blnIsOpen  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
3200        DoEvents
3210        blnIsOpen = False
3220        .frmAccountProfile_RelAccts_Sub.Form.Desc.ForeColor = CLR_BLK
3230        .frmAccountProfile_RelAccts_Sub.Form.Desc.BackStyle = acBackStyleNormal
3240        .frmAccountProfile_RelAccts_Sub.Form.IsRelated_box1.ForeColor = CLR_IEE_RED
3250        .frmAccountProfile_RelAccts_Sub.Form.IsRelated_box2.ForeColor = CLR_IEE_RED
3260        .frmAccountProfile_RelAccts_Sub.SetFocus
3270      ElseIf blnRptVis = True Then
3280        blnRptVis = False
3290        If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3300          If Forms(strCallingForm).Visible = True Then
3310            Forms(strCallingForm).Visible = False
3320            DoEvents
3330          End If
3340          If Forms(strCallingForm).Visible = True Then
3350            blnRptVis = True
3360            .TimerInterval = 1000&
3370          End If
3380        End If
3390      End If
3400      If gblnGoToReport = True Then
3410        Select Case blnGoingToReport
            Case True
3420          DoCmd.Hourglass False  ' ** It should already be off.
3430          blnGoingToReport = False
3440          gblnGoToReport = False
3450          .GoToReport_arw_pr_img.Visible = False
3460          DoEvents
3470        Case False
3480          DoCmd.Hourglass True  ' ** Make sure it's still running.
3490          DoEvents
3500          blnGoingToReport = True
3510          .GoToReport_arw_pr_img.Visible = True
3520          If .cmdPrintReport.Enabled = True Then
3530            .cmdPrintReport.SetFocus
3540          Else
3550            .cmdCancel.SetFocus
3560          End If
3570          DoEvents
3580          Forms(strCallingForm).frmAccountProfile_Sub.Form.GoToReport_arw_relate_img.Visible = False
3590          Forms(strCallingForm).frmAccountProfile_Sub.Form.cmdChoose_lbl.Visible = True
3600          Forms(strCallingForm).frmAccountProfile_Sub.Form.cmdChoose_lbl2.Visible = True
3610          DoEvents
3620          Forms(strCallingForm).GTREmblem_Set False  ' ** Form Procedure: frmAccountProfile.
3630          Beep
3640          DoCmd.Hourglass False
3650          DoEvents
3660          .TimerInterval = (GTR_WAIT * 2&)
3670        End Select
3680      Else
3690        DoCmd.Hourglass False
3700      End If
3710    End With

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
3740    Case Else
3750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3760    End Select
3770    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3810    If DataErr <> 0 Then
3820      Select Case DataErr
          Case 7787  ' ** The record your are modifying has been changed by another user.
            ' ** Ignore.
3830        Response = acDataErrContinue
3840      Case Else
3850        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3860        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3870      End Select
3880    End If

EXITP:
3890    Exit Sub

ERRH:
3900    Select Case ERR.Number
        Case Else
3910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3920    End Select
3930    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String

4010    strDocName = "rptListOfRelatedAccounts"
4020    If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
4030      DoCmd.Close acReport, strDocName
4040    End If

4050    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
4060    Exit Sub

ERRH:
4070    THAT_PROC = THIS_PROC
4080    That_Erl = Erl
4090    Select Case ERR.Number
        Case Else
4100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select
4120    Resume EXITP

End Sub

Private Sub Form_Close()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Close"

4210    With Me
4220      .TimerInterval = 0&
4230      If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
4240        DoCmd.SelectObject acForm, strCallingForm, False
4250      Else
4260        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Account"
4270      End If
4280    End With

EXITP:
4290    Exit Sub

ERRH:
4300    THAT_PROC = THIS_PROC
4310    That_Erl = Erl
4320    Select Case ERR.Number
        Case Else
4330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4340    End Select
4350    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4410    With Me
4420      If .cmdSelectReset_raised_focus_dots_img.Visible = True Or .cmdSelectReset_raised_focus_img.Visible = True Then
4430        Select Case blnSelectReset_Focus
            Case True
4440          .cmdSelectReset_raised_semifocus_dots_img.Visible = True
4450          .cmdSelectReset_raised_img.Visible = False
4460        Case False
4470          .cmdSelectReset_raised_img.Visible = True
4480          .cmdSelectReset_raised_semifocus_dots_img.Visible = False
4490        End Select
4500        .cmdSelectReset_raised_focus_img.Visible = False
4510        .cmdSelectReset_raised_focus_dots_img.Visible = False
4520        .cmdSelectReset_sunken_focus_dots_img.Visible = False
4530        .cmdSelectReset_raised_img_dis.Visible = False
4540      End If
4550      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
4560        Select Case blnSelectNone_Focus
            Case True
4570          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
4580          .cmdSelectNone_raised_img.Visible = False
4590        Case False
4600          .cmdSelectNone_raised_img.Visible = True
4610          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
4620        End Select
4630        .cmdSelectNone_raised_focus_img.Visible = False
4640        .cmdSelectNone_raised_focus_dots_img.Visible = False
4650        .cmdSelectNone_sunken_focus_dots_img.Visible = False
4660        .cmdSelectNone_raised_img_dis.Visible = False
4670      End If
4680      If .cmdClearClosed_raised_focus_dots_img.Visible = True Or .cmdClearClosed_raised_focus_img.Visible = True Then
4690        Select Case blnClearClosed_Focus
            Case True
4700          .cmdClearClosed_raised_semifocus_dots_img.Visible = True
4710          .cmdClearClosed_raised_img.Visible = False
4720        Case False
4730          .cmdClearClosed_raised_img.Visible = True
4740          .cmdClearClosed_raised_semifocus_dots_img.Visible = False
4750        End Select
4760        .cmdClearClosed_raised_focus_img.Visible = False
4770        .cmdClearClosed_raised_focus_dots_img.Visible = False
4780        .cmdClearClosed_sunken_focus_dots_img.Visible = False
4790        .cmdClearClosed_raised_img_dis.Visible = False
4800      End If
4810      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
4820        Select Case blnPrintReport_Focus
            Case True
4830          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
4840          .cmdPrintReport_raised_img.Visible = False
4850        Case False
4860          .cmdPrintReport_raised_img.Visible = True
4870          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4880        End Select
4890        .cmdPrintReport_raised_focus_img.Visible = False
4900        .cmdPrintReport_raised_focus_dots_img.Visible = False
4910        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4920        .cmdPrintReport_raised_img_dis.Visible = False
4930      End If
4940    End With

EXITP:
4950    Exit Sub

ERRH:
4960    Select Case ERR.Number
        Case Else
4970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4980    End Select
4990    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

5010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub cmdOK_Click()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRels As Long, arr_varRel As Variant
        Dim strAcct As String
        Dim blnFound As Boolean
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long, lngE As Long

5110    With Me

5120      arr_varRel = .frmAccountProfile_RelAccts_Sub.Form.RelArray_Get  ' ** Form Function: frmAccountProfile_RelAccts_Sub.
5130      lngRels = (UBound(arr_varRel, 2) + 1&)
5140      For lngX = 0& To (lngRels - 1&)
5150        For lngY = 0& To (lngRAs - 1&)
5160          If arr_varRA(RA_ACCT, lngY) = arr_varRel(RA_ACCT, lngX) Then
5170            arr_varRA(RA_FOUND, lngY) = arr_varRel(RA_FOUND, lngX)
5180            arr_varRA(RA_LST_ELEM, lngY) = arr_varRel(RA_LST_ELEM, lngX)
5190            Exit For
5200          End If
5210        Next
5220      Next

5230      If lngRAs > 0& Then

            ' ** Make sure the list still contains this accountno.
5240        blnFound = False
5250        For lngX = 0& To (lngRAs - 1)
5260          If arr_varRA(RA_ACCT, lngX) = Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno Then
5270            blnFound = True
5280            Exit For
5290          End If
5300        Next
5310        If blnFound = False Then
5320          lngRAs = lngRAs + 1&
5330          lngE = lngRAs - 1&
5340          ReDim Preserve arr_varRA(RA_ELEMS, lngE)
5350          arr_varRA(RA_ACCT, lngE) = Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno
5360          arr_varRA(RA_FOUND, lngE) = CBool(True)
5370          If IsNull(Forms(strCallingForm).Controls(strCallingFormSub).Form.dateclosed) = True Then
5380            arr_varRA(RA_CLOSED, lngE) = CBool(False)
5390          Else
5400            arr_varRA(RA_CLOSED, lngE) = CBool(True)
5410          End If
5420          arr_varRA(RA_LST_ELEM, lngE) = -1&
5430          .related_accountno = .related_accountno & "," & Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno
5440          DoEvents
5450          Beep
5460          MsgBox "Related Accounts must include this Account Number.", vbInformation + vbOKOnly, "Current Account Required"
5470        End If

5480        strTmp01 = .related_accountno

5490        Set dbs = CurrentDb
            ' ** Put the same list in all the related accounts.
5500        strAcct = Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno
5510        For lngX = 0& To (lngRAs - 1&)
              ' ** Update Account for related accounts, by specified [actno], [relacct].
5520          Set rst = Forms(strCallingForm).Controls(strCallingFormSub).Form.RecordsetClone
5530          With rst
5540            .FindFirst "[accountno] = '" & arr_varRA(RA_ACCT, lngX) & "'"
5550            If .NoMatch = False Then
5560              .Edit
5570              ![related_accountno] = strTmp01
5580              .Update
5590            End If
5600            .Close
5610          End With
5620        Next
5630        With Forms(strCallingForm).Controls(strCallingFormSub).Form
5640          .Requery
5650          Set rst = .RecordsetClone
5660          With rst
5670            .FindFirst "[accountno] = '" & strAcct & "'"
5680            If .NoMatch = False Then
5690              Forms(strCallingForm).Controls(strCallingFormSub).Form.Bookmark = .Bookmark
5700            End If
5710            .Close
5720          End With
5730        End With

5740        dbs.Close
5750      End If

5760      If strRelAccts <> vbNullString Then
5770        strAcct = Forms(strCallingForm).Controls(strCallingFormSub).Form.accountno
            ' ** There were existing related accounts.
5780        If IsNull(.related_accountno) = False Then
5790          If strRelAccts <> .related_accountno Then
                ' ** The list is different.
5800            For lngX = 0& To (lngORAs - 1&)
5810              arr_varORA(RA_FOUND, lngX) = False
5820              For lngY = 0& To (lngRAs - 1&)
5830                If arr_varRA(RA_ACCT, lngY) = arr_varORA(RA_ACCT, lngX) Then
5840                  arr_varORA(RA_FOUND, lngX) = True
5850                  Exit For
5860                End If
5870              Next
5880              If arr_varORA(RA_FOUND, lngX) = False Then
                    ' ** This account no longer specified 'related'.
                    ' ** Null out related_accountno in Account, by specified [actno].
5890                With Forms(strCallingForm).Controls(strCallingFormSub).Form
5900                  Set rst = .RecordsetClone
5910                  With rst
5920                    .FindFirst "[accountno] = '" & arr_varORA(RA_ACCT, lngX) & "'"
5930                    If .NoMatch = False Then
5940                      .Edit
5950                      ![related_accountno] = Null
5960                      .Update
5970                    End If
5980                    .Close
5990                  End With
6000                End With
6010              End If
6020            Next
6030          End If
6040        Else
              ' ** All related accounts were cleared.
              ' ** Update Account, related_accountno = Null, by specified [relacts].
6050          With Forms(strCallingForm).Controls(strCallingFormSub).Form
6060            Set rst = .RecordsetClone
6070            With rst
6080              Do While True
6090                .FindFirst "[related_accountno] = '" & strRelAccts & "'"
6100                If .NoMatch = False Then
6110                  .Edit
6120                  ![related_accountno] = Null
6130                  .Update
6140                Else
6150                  Exit Do
6160                End If
6170              Loop
6180              .Close
6190            End With
6200          End With
6210        End If
6220        With Forms(strCallingForm).Controls(strCallingFormSub).Form
6230          .Requery
6240          Set rst = .RecordsetClone
6250          With rst
6260            .FindFirst "[accountno] = '" & strAcct & "'"
6270            If .NoMatch = False Then
6280              Forms(strCallingForm).Controls(strCallingFormSub).Form.Bookmark = .Bookmark
6290            End If
6300            .Close
6310          End With
6320        End With
6330      End If

          ' ** Update frmAccountProfile.
6340      Forms(strCallingForm).Controls(strCallingFormSub).Form.Refresh

6350    End With

6360    DoCmd.Close acForm, THIS_NAME

EXITP:
6370    Set rst = Nothing
6380    Set qdf = Nothing
6390    Set dbs = Nothing
6400    Exit Sub

ERRH:
6410    THAT_PROC = THIS_PROC
6420    That_Erl = Erl
6430    Select Case ERR.Number
          'Case 3218  ' ** Couldn't update; currently locked.
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6520    intShiftDown = (Shift And acShiftMask) > 0
6530    intAltDown = (Shift And acAltMask) > 0
6540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6560      Select Case intRetVal
          Case vbKeyTab
6570        With Me
6580          intRetVal = 0
6590          .cmdCancel.SetFocus
6600        End With
6610      End Select
6620    End If

        ' ** Shift keys.
6630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6640      Select Case intRetVal
          Case vbKeyTab
6650        With Me
6660          intRetVal = 0
6670          If .cmdPrintReport.Enabled = True Then
6680            .cmdPrintReport.SetFocus
6690          Else
6700            .related_accountno.SetFocus
6710          End If
6720        End With
6730      End Select
6740    End If

EXITP:
6750    KeyCode = intRetVal
6760    Exit Sub

ERRH:
6770    intRetVal = 0
6780    Select Case ERR.Number
        Case Else
6790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6800    End Select
6810    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

6910    DoCmd.Close acForm, THIS_NAME

EXITP:
6920    Exit Sub

ERRH:
6930    Select Case ERR.Number
        Case Else
6940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6950    End Select
6960    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
7090          .frmAccountProfile_RelAccts_Sub.SetFocus
7100        End With
7110      End Select
7120    End If

        ' ** Shift keys.
7130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7140      Select Case intRetVal
          Case vbKeyTab
7150        With Me
7160          intRetVal = 0
7170          .cmdOK.SetFocus
7180        End With
7190      End Select
7200    End If

EXITP:
7210    KeyCode = intRetVal
7220    Exit Sub

ERRH:
7230    intRetVal = 0
7240    Select Case ERR.Number
        Case Else
7250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7260    End Select
7270    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

7310    With Me
7320      blnPrintReport_Focus = True
7330      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
7340      .cmdPrintReport_raised_img.Visible = False
7350      .cmdPrintReport_raised_focus_img.Visible = False
7360      .cmdPrintReport_raised_focus_dots_img.Visible = False
7370      .cmdPrintReport_sunken_focus_dots_img.Visible = False
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

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

7510    With Me
7520      blnPrintReport_MouseDown = True
7530      .cmdPrintReport_sunken_focus_dots_img.Visible = True
7540      .cmdPrintReport_raised_img.Visible = False
7550      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7560      .cmdPrintReport_raised_focus_img.Visible = False
7570      .cmdPrintReport_raised_focus_dots_img.Visible = False
7580      .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_Click()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

7710    With Me
7720      strDocName = "rptListOfRelatedAccounts"
7730      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
7740        Forms(strCallingForm).Visible = False
7750        Forms(THIS_NAME).Visible = False
7760        DoCmd.OpenReport strDocName, acViewPreview
7770        DoCmd.Maximize
7780        DoCmd.RunCommand acCmdFitToWindow
7790        blnRptVis = True
7800        .TimerInterval = 1000&
7810      Else
            '##GTR_Ref: rptListOfRelatedAccounts
7820        DoCmd.OpenReport strDocName, acViewNormal
7830      End If
7840    End With

EXITP:
7850    Exit Sub

ERRH:
7860    Forms(THIS_NAME).Visible = True
7870    Forms(strCallingForm).Visible = True
7880    Select Case ERR.Number
        Case Else
7890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7900    End Select
7910    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8020    intShiftDown = (Shift And acShiftMask) > 0
8030    intAltDown = (Shift And acAltMask) > 0
8040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8060      Select Case intRetVal
          Case vbKeyTab
8070        With Me
8080          intRetVal = 0
8090          .cmdOK.SetFocus
8100        End With
8110      End Select
8120    End If

        ' ** Shift keys.
8130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8140      Select Case intRetVal
          Case vbKeyTab
8150        With Me
8160          intRetVal = 0
8170          .related_accountno.SetFocus
8180        End With
8190      End Select
8200    End If

EXITP:
8210    KeyCode = intRetVal
8220    Exit Sub

ERRH:
8230    intRetVal = 0
8240    Select Case ERR.Number
        Case Else
8250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8260    End Select
8270    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

8310    With Me
8320      If blnPrintReport_MouseDown = False Then
8330        Select Case blnPrintReport_Focus
            Case True
8340          .cmdPrintReport_raised_focus_dots_img.Visible = True
8350          .cmdPrintReport_raised_focus_img.Visible = False
8360        Case False
8370          .cmdPrintReport_raised_focus_img.Visible = True
8380          .cmdPrintReport_raised_focus_dots_img.Visible = False
8390        End Select
8400        .cmdPrintReport_raised_img.Visible = False
8410        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8420        .cmdPrintReport_sunken_focus_dots_img.Visible = False
8430        .cmdPrintReport_raised_img_dis.Visible = False
8440      End If
8450    End With

EXITP:
8460    Exit Sub

ERRH:
8470    Select Case ERR.Number
        Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

8610    With Me
8620      .cmdPrintReport_raised_focus_dots_img.Visible = True
8630      .cmdPrintReport_raised_img.Visible = False
8640      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8650      .cmdPrintReport_raised_focus_img.Visible = False
8660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
8670      .cmdPrintReport_raised_img_dis.Visible = False
8680      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

8810    With Me
8820      .cmdPrintReport_raised_img.Visible = True
8830      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8840      .cmdPrintReport_raised_focus_img.Visible = False
8850      .cmdPrintReport_raised_focus_dots_img.Visible = False
8860      .cmdPrintReport_sunken_focus_dots_img.Visible = False
8870      .cmdPrintReport_raised_img_dis.Visible = False
8880      blnPrintReport_Focus = False
8890    End With

EXITP:
8900    Exit Sub

ERRH:
8910    Select Case ERR.Number
        Case Else
8920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8930    End Select
8940    Resume EXITP

End Sub

Private Sub frmAccountProfile_RelAccts_Sub_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "frmAccountProfile_RelAccts_Sub_box_MouseMove"

9010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9020    Exit Sub

ERRH:
9030    Select Case ERR.Number
        Case Else
9040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9050    End Select
9060    Resume EXITP

End Sub

Private Sub opgOrderBy_AfterUpdate()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgOrderBy_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQry_New As String

9110    With Me
9120      .frmAccountProfile_RelAccts_Sub.Form.Desc.ForeColor = MY_CLR_BGE
9130      .frmAccountProfile_RelAccts_Sub.Form.Desc.BackStyle = acBackStyleTransparent
9140      strQry_New = vbNullString
9150      Select Case .opgOrderBy
          Case .opgOrderBy_optNumber.OptionValue
9160        .opgOrderBy_optNumber_lbl.FontBold = True
9170        .opgOrderBy_optName_lbl.FontBold = False
9180        .opgOrderBy_optSelected_lbl.FontBold = False
            ' ** Append qryAccountProfile_RelAccts_03 (Account, linked to
            ' ** qryAccountMenu_01_10 (xx), qryAccountNoDropDown_06 (xx), accountno,
            ' ** Desc, with cash criteria; sort: accountno) to tmpAccount_RelAccts.
9190        strQry_New = "qryAccountProfile_RelAccts_03a"
9200        .frmAccountProfile_RelAccts_Sub.Form.RecordSource = "qryAccountProfile_RelAccts_06_01"
9210      Case .opgOrderBy_optName.OptionValue
9220        .opgOrderBy_optNumber_lbl.FontBold = False
9230        .opgOrderBy_optName_lbl.FontBold = True
9240        .opgOrderBy_optSelected_lbl.FontBold = False
            ' ** Append qryAccountProfile_RelAccts_04 (Account, linked to
            ' ** qryAccountMenu_01_10 (xx), qryAccountNoDropDown_06 (xx), accountno,
            ' ** shortname, with cash criteria; sort: shortname) to tmpAccount_RelAccts.
9250        strQry_New = "qryAccountProfile_RelAccts_04a"
9260        .frmAccountProfile_RelAccts_Sub.Form.RecordSource = "qryAccountProfile_RelAccts_06_02"
9270      Case .opgOrderBy_optSelected.OptionValue
9280        .opgOrderBy_optNumber_lbl.FontBold = False
9290        .opgOrderBy_optName_lbl.FontBold = False
9300        .opgOrderBy_optSelected_lbl.FontBold = True
            ' ** Append qryAccountProfile_RelAccts_05 (Account, linked to
            ' ** qryAccountMenu_01_10 (xx), qryAccountNoDropDown_06 (xx), accountno,
            ' ** shortname, with cash criteria; sort: IsRelated) to tmpAccount_RelAccts.
9310        strQry_New = "qryAccountProfile_RelAccts_05a"
9320        .frmAccountProfile_RelAccts_Sub.Form.RecordSource = "qryAccountProfile_RelAccts_06_03"
9330      End Select
9340      If strQry_New <> vbNullString And strQry_New <> strQryLastAppend Then
9350        Set dbs = CurrentDb
9360        With dbs
              ' ** Empty tmpAccount_RelAccts.
9370          Set qdf = .QueryDefs("qryAccountProfile_RelAccts_02")
9380          qdf.Execute
9390          Set qdf = Nothing
9400          Set qdf = .QueryDefs(strQry_New)
9410          qdf.Execute
9420          Set qdf = Nothing
9430          strQryLastAppend = strQry_New
9440          .Close
9450        End With
9460        Set dbs = Nothing
            'If blnIsOpen = False Then
            '  .frmAccountProfile_RelAccts_Sub.Form.Desc.ForeColor = CLR_BLK
            '  .frmAccountProfile_RelAccts_Sub.Form.Desc.BackStyle = acBackStyleNormal
            'End If
9470        blnOrderBy2 = True
9480        .TimerInterval = 100&
9490      End If
9500    End With

EXITP:
9510    Set qdf = Nothing
9520    Set dbs = Nothing
9530    Exit Sub

ERRH:
9540    Select Case ERR.Number
        Case Else
9550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9560    End Select
9570    Resume EXITP

End Sub

Private Sub cmdClearClosed_GotFocus()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearClosed_GotFocus"

9610    With Me
9620      blnClearClosed_Focus = True
9630      .cmdClearClosed_raised_semifocus_dots_img.Visible = True
9640      .cmdClearClosed_raised_img.Visible = False
9650      .cmdClearClosed_raised_focus_img.Visible = False
9660      .cmdClearClosed_raised_focus_dots_img.Visible = False
9670      .cmdClearClosed_sunken_focus_dots_img.Visible = False
9680      .cmdClearClosed_raised_img_dis.Visible = False
9690    End With

EXITP:
9700    Exit Sub

ERRH:
9710    Select Case ERR.Number
        Case Else
9720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9730    End Select
9740    Resume EXITP

End Sub

Private Sub cmdClearClosed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearClosed_MouseDown"

9810    With Me
9820      blnClearClosed_MouseDown = True
9830      .cmdClearClosed_sunken_focus_dots_img.Visible = True
9840      .cmdClearClosed_raised_img.Visible = False
9850      .cmdClearClosed_raised_semifocus_dots_img.Visible = False
9860      .cmdClearClosed_raised_focus_img.Visible = False
9870      .cmdClearClosed_raised_focus_dots_img.Visible = False
9880      .cmdClearClosed_raised_img_dis.Visible = False
9890    End With

EXITP:
9900    Exit Sub

ERRH:
9910    Select Case ERR.Number
        Case Else
9920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9930    End Select
9940    Resume EXITP

End Sub

Private Sub cmdClearClosed_Click()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearClosed_Click"

        Dim lngRels As Long, arr_varRel() As Variant
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long, lngE As Long

10010   With Me
10020     .ClosedAccts = Null
10030     .ClosedAccts.Visible = False
10040     .cmdPrintReport.SetFocus
10050     .cmdClearClosed.Visible = False
10060     If lngCRAs = lngRAs Then
10070       lngRAs = 0&
10080       ReDim arr_varRA(RA_ELEMS, 0)
10090       lngCRAs = 0&
10100       ReDim arr_varCRA(RA_ELEMS, 0)
10110       .related_accountno = Null
10120     Else
10130       lngRels = 0&
10140       ReDim arr_varRel(RA_ELEMS, 0)
10150       For lngX = 0& To (lngCRAs - 1&)
10160         For lngY = 0& To (lngRAs - 1&)
10170           If arr_varRA(RA_ACCT, lngY) = arr_varCRA(RA_ACCT, lngX) Then
10180             arr_varRA(RA_ACCT, lngY) = vbNullString
10190             Exit For
10200           End If
10210         Next
10220       Next
10230       For lngX = 0& To (lngRAs - 1&)
10240         If arr_varRA(RA_ACCT, lngX) <> vbNullString Then
10250           lngRels = lngRels + 1
10260           lngE = lngRels - 1&
10270           ReDim Preserve arr_varRel(RA_ELEMS, lngE)
10280           arr_varRel(RA_ACCT, lngE) = arr_varRA(RA_ACCT, lngX)
10290           arr_varRel(RA_FOUND, lngE) = arr_varRA(RA_FOUND, lngX)
10300           arr_varRel(RA_CLOSED, lngE) = arr_varRA(RA_CLOSED, lngX)
10310           arr_varRel(RA_LST_ELEM, lngE) = arr_varRA(RA_LST_ELEM, lngX)
10320         End If
10330       Next
10340       lngRAs = 0&
10350       ReDim arr_varRA(RA_ELEMS, 0)
10360       strTmp01 = vbNullString
10370       For lngX = 0& To (lngRels - 1&)
10380         lngRAs = lngRAs + 1&
10390         lngE = lngRAs - 1&
10400         ReDim Preserve arr_varRA(RA_ELEMS, lngE)
10410         arr_varRA(RA_ACCT, lngE) = arr_varRel(RA_ACCT, lngX)
10420         arr_varRA(RA_FOUND, lngE) = arr_varRel(RA_FOUND, lngX)
10430         arr_varRA(RA_CLOSED, lngE) = arr_varRel(RA_CLOSED, lngX)
10440         arr_varRA(RA_LST_ELEM, lngE) = arr_varRel(RA_LST_ELEM, lngX)
10450         strTmp01 = strTmp01 & arr_varRel(RA_ACCT, lngX) & ","
10460       Next
10470       strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
10480       .related_accountno = strTmp01
10490       .frmAccountProfile_RelAccts_Sub.Form.RelArray_Set arr_varRA, lngRAs  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
10500       .frmAccountProfile_RelAccts_Sub.Form.CRelArray_Set arr_varCRA, lngCRAs  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
10510     End If
10520   End With

EXITP:
10530   Exit Sub

ERRH:
10540   Select Case ERR.Number
        Case Else
10550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10560   End Select
10570   Resume EXITP

End Sub

Private Sub cmdClearClosed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearClosed_MouseMove"

10610   With Me
10620     If blnClearClosed_MouseDown = False Then
10630       Select Case blnClearClosed_Focus
            Case True
10640         .cmdClearClosed_raised_focus_dots_img.Visible = True
10650         .cmdClearClosed_raised_focus_img.Visible = False
10660       Case False
10670         .cmdClearClosed_raised_focus_img.Visible = True
10680         .cmdClearClosed_raised_focus_dots_img.Visible = False
10690       End Select
10700       .cmdClearClosed_raised_img.Visible = False
10710       .cmdClearClosed_raised_semifocus_dots_img.Visible = False
10720       .cmdClearClosed_sunken_focus_dots_img.Visible = False
10730       .cmdClearClosed_raised_img_dis.Visible = False
10740     End If
10750   End With

EXITP:
10760   Exit Sub

ERRH:
10770   Select Case ERR.Number
        Case Else
10780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10790   End Select
10800   Resume EXITP

End Sub

Private Sub cmdClearClosed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearClosed_MouseUp"

10910   With Me
10920     .cmdClearClosed_raised_focus_dots_img.Visible = True
10930     .cmdClearClosed_raised_img.Visible = False
10940     .cmdClearClosed_raised_semifocus_dots_img.Visible = False
10950     .cmdClearClosed_raised_focus_img.Visible = False
10960     .cmdClearClosed_sunken_focus_dots_img.Visible = False
10970     .cmdClearClosed_raised_img_dis.Visible = False
10980     blnClearClosed_MouseDown = False
10990   End With

EXITP:
11000   Exit Sub

ERRH:
11010   Select Case ERR.Number
        Case Else
11020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11030   End Select
11040   Resume EXITP

End Sub

Private Sub cmdClearClosed_LostFocus()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearClosed_LostFocus"

11110   With Me
11120     .cmdClearClosed_raised_img.Visible = True
11130     .cmdClearClosed_raised_semifocus_dots_img.Visible = False
11140     .cmdClearClosed_raised_focus_img.Visible = False
11150     .cmdClearClosed_raised_focus_dots_img.Visible = False
11160     .cmdClearClosed_sunken_focus_dots_img.Visible = False
11170     .cmdClearClosed_raised_img_dis.Visible = False
11180     blnClearClosed_Focus = False
11190   End With

EXITP:
11200   Exit Sub

ERRH:
11210   Select Case ERR.Number
        Case Else
11220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11230   End Select
11240   Resume EXITP

End Sub

Private Sub cmdSelectReset_GotFocus()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectReset_GotFocus"

11310   With Me
11320     blnSelectReset_Focus = True
11330     .cmdSelectReset_raised_semifocus_dots_img.Visible = True
11340     .cmdSelectReset_raised_img.Visible = False
11350     .cmdSelectReset_raised_focus_img.Visible = False
11360     .cmdSelectReset_raised_focus_dots_img.Visible = False
11370     .cmdSelectReset_sunken_focus_dots_img.Visible = False
11380     .cmdSelectReset_raised_img_dis.Visible = False
11390   End With

EXITP:
11400   Exit Sub

ERRH:
11410   Select Case ERR.Number
        Case Else
11420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11430   End Select
11440   Resume EXITP

End Sub

Private Sub cmdSelectReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectReset_MouseDown"

11510   With Me
11520     blnSelectReset_MouseDown = True
11530     .cmdSelectReset_sunken_focus_dots_img.Visible = True
11540     .cmdSelectReset_raised_img.Visible = False
11550     .cmdSelectReset_raised_semifocus_dots_img.Visible = False
11560     .cmdSelectReset_raised_focus_img.Visible = False
11570     .cmdSelectReset_raised_focus_dots_img.Visible = False
11580     .cmdSelectReset_raised_img_dis.Visible = False
11590   End With

EXITP:
11600   Exit Sub

ERRH:
11610   Select Case ERR.Number
        Case Else
11620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11630   End Select
11640   Resume EXITP

End Sub

Private Sub cmdSelectReset_Click()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectReset_Click"

        Dim strTmp01 As String
        Dim lngX As Long

11710   With Me
11720     .related_accountno = Null
11730     strTmp01 = vbNullString
11740     If lngORAs > 0& Then
11750       For lngX = 0& To (lngORAs - 1&)
11760         strTmp01 = strTmp01 & arr_varORA(RA_ACCT, lngX) & ","
11770       Next
11780       If Right(strTmp01, 1) = "," Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
11790       .related_accountno = strTmp01
11800     End If
          ' ** Highlight the related accounts.
11810     .frmAccountProfile_RelAccts_Sub.Form.HighlightRelated arr_varORA, lngORAs, True  ' ** Form Procedure: frmAccountProfile_RelAccts_Sub.
11820   End With

EXITP:
11830   Exit Sub

ERRH:
11840   Select Case ERR.Number
        Case Else
11850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11860   End Select
11870   Resume EXITP

End Sub

Private Sub cmdSelectReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectReset_MouseMove"

11910   With Me
11920     If blnSelectReset_MouseDown = False Then
11930       Select Case blnSelectReset_Focus
            Case True
11940         .cmdSelectReset_raised_focus_dots_img.Visible = True
11950         .cmdSelectReset_raised_focus_img.Visible = False
11960       Case False
11970         .cmdSelectReset_raised_focus_img.Visible = True
11980         .cmdSelectReset_raised_focus_dots_img.Visible = False
11990       End Select
12000       .cmdSelectReset_raised_img.Visible = False
12010       .cmdSelectReset_raised_semifocus_dots_img.Visible = False
12020       .cmdSelectReset_sunken_focus_dots_img.Visible = False
12030       .cmdSelectReset_raised_img_dis.Visible = False
12040     End If
12050     If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
12060       Select Case blnSelectNone_Focus
            Case True
12070         .cmdSelectNone_raised_semifocus_dots_img.Visible = True
12080         .cmdSelectNone_raised_img.Visible = False
12090       Case False
12100         .cmdSelectNone_raised_img.Visible = True
12110         .cmdSelectNone_raised_semifocus_dots_img.Visible = False
12120       End Select
12130       .cmdSelectNone_raised_focus_img.Visible = False
12140       .cmdSelectNone_raised_focus_dots_img.Visible = False
12150       .cmdSelectNone_sunken_focus_dots_img.Visible = False
12160       .cmdSelectNone_raised_img_dis.Visible = False
12170     End If
12180   End With

EXITP:
12190   Exit Sub

ERRH:
12200   Select Case ERR.Number
        Case Else
12210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12220   End Select
12230   Resume EXITP

End Sub

Private Sub cmdSelectReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectReset_MouseUp"

12310   With Me
12320     .cmdSelectReset_raised_focus_dots_img.Visible = True
12330     .cmdSelectReset_raised_img.Visible = False
12340     .cmdSelectReset_raised_semifocus_dots_img.Visible = False
12350     .cmdSelectReset_raised_focus_img.Visible = False
12360     .cmdSelectReset_sunken_focus_dots_img.Visible = False
12370     .cmdSelectReset_raised_img_dis.Visible = False
12380     blnSelectReset_MouseDown = False
12390   End With

EXITP:
12400   Exit Sub

ERRH:
12410   Select Case ERR.Number
        Case Else
12420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12430   End Select
12440   Resume EXITP

End Sub

Private Sub cmdSelectReset_LostFocus()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectReset_LostFocus"

12510   With Me
12520     .cmdSelectReset_raised_img.Visible = True
12530     .cmdSelectReset_raised_semifocus_dots_img.Visible = False
12540     .cmdSelectReset_raised_focus_img.Visible = False
12550     .cmdSelectReset_raised_focus_dots_img.Visible = False
12560     .cmdSelectReset_sunken_focus_dots_img.Visible = False
12570     .cmdSelectReset_raised_img_dis.Visible = False
12580     blnSelectReset_Focus = False
12590   End With

EXITP:
12600   Exit Sub

ERRH:
12610   Select Case ERR.Number
        Case Else
12620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12630   End Select
12640   Resume EXITP

End Sub

Private Sub cmdSelectNone_GotFocus()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_GotFocus"

12710   With Me
12720     blnSelectNone_Focus = True
12730     .cmdSelectNone_raised_semifocus_dots_img.Visible = True
12740     .cmdSelectNone_raised_img.Visible = False
12750     .cmdSelectNone_raised_focus_img.Visible = False
12760     .cmdSelectNone_raised_focus_dots_img.Visible = False
12770     .cmdSelectNone_sunken_focus_dots_img.Visible = False
12780     .cmdSelectNone_raised_img_dis.Visible = False
12790   End With

EXITP:
12800   Exit Sub

ERRH:
12810   Select Case ERR.Number
        Case Else
12820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12830   End Select
12840   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseDown"

12910   With Me
12920     blnSelectNone_MouseDown = True
12930     .cmdSelectNone_sunken_focus_dots_img.Visible = True
12940     .cmdSelectNone_raised_img.Visible = False
12950     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
12960     .cmdSelectNone_raised_focus_img.Visible = False
12970     .cmdSelectNone_raised_focus_dots_img.Visible = False
12980     .cmdSelectNone_raised_img_dis.Visible = False
12990   End With

EXITP:
13000   Exit Sub

ERRH:
13010   Select Case ERR.Number
        Case Else
13020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13030   End Select
13040   Resume EXITP

End Sub

Private Sub cmdSelectNone_Click()

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_Click"

        Dim strTmp01 As String
        Dim lngX As Long, lngE As Long

13110   With Me

13120     lngRAs = 0&
13130     ReDim arr_varRA(RA_ELEMS, 0)
13140     strTmp01 = vbNullString

13150     If lngCRAs > 0 Then
            ' ** Add the closed accounts.
13160       For lngX = 0& To (lngCRAs - 1&)
13170         lngRAs = lngRAs + 1&
13180         lngE = lngRAs - 1&
13190         ReDim Preserve arr_varRA(RA_ELEMS, lngE)
13200         arr_varRA(RA_ACCT, lngE) = arr_varCRA(RA_ACCT, lngX)
13210         arr_varRA(RA_FOUND, lngE) = arr_varCRA(RA_FOUND, lngX)
13220         arr_varRA(RA_CLOSED, lngE) = arr_varCRA(RA_CLOSED, lngX)
13230         arr_varRA(RA_LST_ELEM, lngE) = arr_varCRA(RA_LST_ELEM, lngX)
13240         strTmp01 = strTmp01 & arr_varCRA(RA_ACCT, lngX) & ","
13250       Next
13260       strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
13270     End If

13280     .frmAccountProfile_RelAccts_Sub.Form.CRelArray_Set arr_varCRA, lngCRAs  ' ** Module Procedure: frmAccountProfile_RelAccts_Sub.

13290     If strTmp01 = vbNullString Then
13300       .related_accountno = Null
13310     Else
13320       .related_accountno = strTmp01
13330     End If

13340     .frmAccountProfile_RelAccts_Sub.Form.HighlightRelated arr_varRA, lngRAs  ' ** Module Procedure: frmAccountProfile_RelAccts_Sub.

13350     If strTmp01 <> vbNullString Then
13360       .cmdClearClosed.SetFocus
13370     End If

13380   End With

EXITP:
13390   Exit Sub

ERRH:
13400   Select Case ERR.Number
        Case Else
13410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13420   End Select
13430   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseMove"

13510   With Me
13520     If blnSelectNone_MouseDown = False Then
13530       Select Case blnSelectNone_Focus
            Case True
13540         .cmdSelectNone_raised_focus_dots_img.Visible = True
13550         .cmdSelectNone_raised_focus_img.Visible = False
13560       Case False
13570         .cmdSelectNone_raised_focus_img.Visible = True
13580         .cmdSelectNone_raised_focus_dots_img.Visible = False
13590       End Select
13600       .cmdSelectNone_raised_img.Visible = False
13610       .cmdSelectNone_raised_semifocus_dots_img.Visible = False
13620       .cmdSelectNone_sunken_focus_dots_img.Visible = False
13630       .cmdSelectNone_raised_img_dis.Visible = False
13640     End If
13650     If .cmdSelectReset_raised_focus_dots_img.Visible = True Or .cmdSelectReset_raised_focus_img.Visible = True Then
13660       Select Case blnSelectReset_Focus
            Case True
13670         .cmdSelectReset_raised_semifocus_dots_img.Visible = True
13680         .cmdSelectReset_raised_img.Visible = False
13690       Case False
13700         .cmdSelectReset_raised_img.Visible = True
13710         .cmdSelectReset_raised_semifocus_dots_img.Visible = False
13720       End Select
13730       .cmdSelectReset_raised_focus_img.Visible = False
13740       .cmdSelectReset_raised_focus_dots_img.Visible = False
13750       .cmdSelectReset_sunken_focus_dots_img.Visible = False
13760       .cmdSelectReset_raised_img_dis.Visible = False
13770     End If
13780   End With

EXITP:
13790   Exit Sub

ERRH:
13800   Select Case ERR.Number
        Case Else
13810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13820   End Select
13830   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseUp"

13910   With Me
13920     .cmdSelectNone_raised_focus_dots_img.Visible = True
13930     .cmdSelectNone_raised_img.Visible = False
13940     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
13950     .cmdSelectNone_raised_focus_img.Visible = False
13960     .cmdSelectNone_sunken_focus_dots_img.Visible = False
13970     .cmdSelectNone_raised_img_dis.Visible = False
13980     blnSelectNone_MouseDown = False
13990   End With

EXITP:
14000   Exit Sub

ERRH:
14010   Select Case ERR.Number
        Case Else
14020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14030   End Select
14040   Resume EXITP

End Sub

Private Sub cmdSelectNone_LostFocus()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_LostFocus"

14110   With Me
14120     .cmdSelectNone_raised_img.Visible = True
14130     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
14140     .cmdSelectNone_raised_focus_img.Visible = False
14150     .cmdSelectNone_raised_focus_dots_img.Visible = False
14160     .cmdSelectNone_sunken_focus_dots_img.Visible = False
14170     .cmdSelectNone_raised_img_dis.Visible = False
14180     blnSelectNone_Focus = False
14190   End With

EXITP:
14200   Exit Sub

ERRH:
14210   Select Case ERR.Number
        Case Else
14220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14230   End Select
14240   Resume EXITP

End Sub

Private Sub related_accountno_GotFocus()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_GotFocus"

14310   With Me
14320     .related_accountno.SelLength = 0
14330     .related_accountno.SelStart = 99
14340   End With

EXITP:
14350   Exit Sub

ERRH:
14360   Select Case ERR.Number
        Case Else
14370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14380   End Select
14390   Resume EXITP

End Sub

Private Sub related_accountno_KeyDown(KeyCode As Integer, Shift As Integer)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14420   intShiftDown = (Shift And acShiftMask) > 0
14430   intAltDown = (Shift And acAltMask) > 0
14440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14470       With Me
14480         intRetVal = 0
14490         .cmdPrintReport.SetFocus
14500       End With
14510     End Select
14520   End If

        ' ** Shift keys.
14530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14550       With Me
14560         intRetVal = 0
14570         .cmdSelectNone.SetFocus
14580       End With
14590     End Select
14600   End If

EXITP:
14610   KeyCode = intRetVal
14620   Exit Sub

ERRH:
14630   intRetVal = 0
14640   Select Case ERR.Number
        Case Else
14650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14660   End Select
14670   Resume EXITP

End Sub
