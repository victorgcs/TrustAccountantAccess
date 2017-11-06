Attribute VB_Name = "modStatementParamFuncs2"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modStatementParamFuncs2"

'VGC 09/05/2017: CHANGES!

' ** Array: arr_varStmt().
Private Const S_ELEMS1 As Integer = 12  ' ** Array's first-element UBound().
Private Const S_ELEMS2 As Integer = 4   ' ** Array's second-element UBound().
Private Const S_MID   As Integer = 0  'month_id
Private Const S_MSHT  As Integer = 1  'month_short
Private Const S_CNT   As Integer = 2  'cnt_smt
Private Const S_ACTNO As Integer = 3  'accountno
Private Const S_SNAM  As Integer = 4  'shortname

' ** Array: arr_varAcctFor().
Private Const F_ACTNO As Integer = 0
Private Const F_JCNT  As Integer = 1
Private Const F_ACNT  As Integer = 2
Private Const F_SUPP  As Integer = 3

' ** Array: arr_varAcctArch().
Private lngAcctArchs As Long, arr_varAcctArch As Variant
Private Const AR_ACTNO As Integer = 0
'Private Const AR_TDATE As Integer = 1
'Private Const AR_CNT   As Integer = 2

' ** cmbAccounts combo box constants:
Private Const CBX_A_ACTNO  As Integer = 0  ' ** accountno
Private Const CBX_A_DESC   As Integer = 1  ' ** Desc
Private Const CBX_A_PREDAT As Integer = 2  ' ** predate
Private Const CBX_A_SHORT  As Integer = 3  ' ** shortname
Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname
Private Const CBX_A_BALDAT As Integer = 5  ' ** BalanceDate (earliest [balance date])
Private Const CBX_A_HASREL As Integer = 6  ' ** HasRelated
Private Const CBX_A_CASNUM As Integer = 7  ' ** CaseNum
Private Const CBX_A_TRXDAT As Integer = 8  ' ** TransDate (earliest [transdate])

' ** cmbMonth combo box constants:
Private Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
Private Const CBX_MON_NAME  As Integer = 1  ' ** month_name
Private Const CBX_MON_SHORT As Integer = 2  ' ** month_short

Private blnIncludeCurrency As Boolean
' **

Public Sub ChkTrans_After_SP(blnIsOpen As Boolean, blnRunPriorStatement As Boolean, blnAcctNotSched As Boolean, datAssetListDate_Pref As Date, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctFors As Long, arr_varAcctFor As Variant, lngAcctArchs As Long, arr_varAcctArch As Variant, lngStmts As Long, arr_varStmt As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** ChkTrans_After_SP(
' **   blnIsOpen As Boolean, blnRunPriorStatement As Boolean, blnAcctNotSched As Boolean,
' **   datAssetListDate_Pref As Date, blnHasForEx As Boolean, blnHasForExThis As Boolean,
' **   lngAcctFors As Long, arr_varAcctFor As Variant, lngAcctArchs As Long, arr_varAcctArch As Variant,
' **   lngStmts As Long, arr_varStmt As Variant,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

100   On Error GoTo ERRH

        Const THIS_PROC As String = "ChkTrans_After_SP"

        Dim lngX As Long

110     With frm

120       DoCmd.Hourglass True
130       DoEvents

140       .ForEx_ChkScheduled_lbl.Visible = False

150       If .HasForeign = True Then
160         blnHasForEx = True
170         Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
180           Select Case IsNull(.cmbAccounts)
              Case True
190             .chkIncludeCurrency = False
200             .chkIncludeCurrency.Enabled = False
210             .chkIncludeCurrency.Locked = False
220             .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
230           Case False
240             .chkIncludeCurrency.Enabled = True
250             gblnHasForExThis = False
260             blnHasForExThis = False
270             For lngX = 0& To (lngAcctFors - 1&)
280               If arr_varAcctFor(F_ACTNO, lngX) = .cmbAccounts Then
290                 If arr_varAcctFor(F_JCNT, lngX) > 0 Or arr_varAcctFor(F_ACNT, lngX) > 0 Then
300                   gblnHasForExThis = True
310                   blnHasForExThis = True
320                 End If
330                 Exit For
340               End If
350             Next
360             Select Case blnHasForExThis
                Case True
370               .chkIncludeCurrency = True
380               .chkIncludeCurrency.Locked = True
390               .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
400             Case False
410               .chkIncludeCurrency.Locked = False
420               .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
430             End Select
440           End Select
450         Case .opgAccountNumber_optAll.OptionValue
460           .chkIncludeCurrency = True
470           .chkIncludeCurrency.Enabled = True
480           .chkIncludeCurrency.Locked = True
490           .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
500         End Select
510       End If

520       Btn_Enable_SP 1, blnIsOpen, blnRunPriorStatement, blnAcctNotSched, datAssetListDate_Pref, lngStmts, arr_varStmt, frm  ' ** Procedure: Below.

530       SetArchiveOption_SP .chkTransactions, .chkAssetList, lngAcctArchs, arr_varAcctArch, frm  ' ** Module Procedure: modStatementParamFuncs2.

540       DoEvents
550       DoCmd.Hourglass False

560     End With

EXITP:
570     Exit Sub

ERRH:
580     DoCmd.Hourglass False
590     THAT_PROC = THIS_PROC
600     That_Erl = Erl: That_Desc = ERR.description
610     frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
620     Resume EXITP

End Sub

Public Sub ChkAstList_After_SP(blnIsOpen As Boolean, blnRunPriorStatement As Boolean, blnAcctNotSched As Boolean, datAssetListDate_Pref As Date, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctFors As Long, arr_varAcctFor As Variant, lngAcctArchs As Long, arr_varAcctArch As Variant, lngStmts As Long, arr_varStmt As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** ChkAstList_After_SP(
' **   blnIsOpen As Boolean, blnRunPriorStatement As Boolean, blnAcctNotSched As Boolean,
' **   datAssetListDate_Pref As Date, blnHasForEx As Boolean, blnHasForExThis As Boolean,
' **   lngAcctFors As Long, arr_varAcctFor As Variant, lngAcctArchs As Long, arr_varAcctArch As Variant,
' **   lngStmts As Long, arr_varStmt As Variant,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

700   On Error GoTo ERRH

        Const THIS_PROC As String = "ChkAstList_After_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngX As Long

710     With frm

720       DoCmd.Hourglass True
730       DoEvents

740       .ForEx_ChkScheduled_lbl.Visible = False

750       Set dbs = CurrentDb
760       With dbs
            ' ** Empty tmpAssetList1.
770         Set qdf = .QueryDefs("qryStatementParameters_AssetList_09b")
780         qdf.Execute
790         Set qdf = Nothing
            ' ** Empty tmpAssetList2.
800         Set qdf = .QueryDefs("qryStatementParameters_AssetList_09c")
810         qdf.Execute
820         Set qdf = Nothing
            ' ** Empty tmpAssetList4.
830         Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_51")
840         qdf.Execute
850         Set qdf = Nothing
            ' ** Empty tmpAssetList5.
860         Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_52")
870         qdf.Execute
880         Set qdf = Nothing
890         .Close
900       End With
910       Set dbs = Nothing

920       If .HasForeign = True Then
930         blnHasForEx = True
940         Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
950           Select Case IsNull(.cmbAccounts)
              Case True
960             .chkIncludeCurrency = False
970             .chkIncludeCurrency.Enabled = False
980             .chkIncludeCurrency.Locked = False
990             .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
1000          Case False
1010            .chkIncludeCurrency.Enabled = True
1020            gblnHasForExThis = False
1030            blnHasForExThis = False
1040            For lngX = 0& To (lngAcctFors - 1&)
1050              If arr_varAcctFor(F_ACTNO, lngX) = .cmbAccounts Then
1060                If arr_varAcctFor(F_JCNT, lngX) > 0 Or arr_varAcctFor(F_ACNT, lngX) > 0 Then
1070                  gblnHasForExThis = True
1080                  blnHasForExThis = True
1090                End If
1100                Exit For
1110              End If
1120            Next
1130            Select Case blnHasForExThis
                Case True
1140              .chkIncludeCurrency = True
1150              .chkIncludeCurrency.Locked = True
1160              .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
1170            Case False
1180              .chkIncludeCurrency.Locked = False
1190              .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
1200            End Select
1210          End Select
1220        Case .opgAccountNumber_optAll.OptionValue
1230          .chkIncludeCurrency = True
1240          .chkIncludeCurrency.Enabled = True
1250          .chkIncludeCurrency.Locked = True
1260          .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
1270        End Select
1280      End If

1290      SetArchiveOption_SP .chkTransactions, .chkAssetList, lngAcctArchs, arr_varAcctArch, frm  ' ** Module Procedure: modStatementParamFuncs2.

1300      DoCmd.Hourglass True  ' ** Assure it's still going.
1310      DoEvents

1320      Btn_Enable_SP 2, blnIsOpen, blnRunPriorStatement, blnAcctNotSched, datAssetListDate_Pref, lngStmts, arr_varStmt, frm  ' ** Procedure: Below.

1330      DoEvents
1340      DoCmd.Hourglass False

1350    End With

EXITP:
1360    Set qdf = Nothing
1370    Set dbs = Nothing
1380    Exit Sub

ERRH:
1390    DoCmd.Hourglass False
1400    THAT_PROC = THIS_PROC
1410    That_Erl = Erl: That_Desc = ERR.description
1420    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1430    Resume EXITP

End Sub

Public Sub ChkStmt_After_SP(blnIsOpen As Boolean, blnRunPriorStatement As Boolean, blnAcctNotSched As Boolean, datAssetListDate_Pref As Date, lngAcctArchs As Long, arr_varAcctArch As Variant, lngAcctFors As Long, arr_varAcctFor As Variant, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngStmts As Long, arr_varStmt As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** ChkStmt_After_SP(
' **   blnIsOpen As Boolean, blnRunPriorStatement As Boolean, blnAcctNotSched As Boolean,
' **   datAssetListDate_Pref As Date, lngAcctArchs As Long, arr_varAcctArch As Variant,
' **   lngAcctFors As Long, arr_varAcctFor As Variant, blnHasForEx As Boolean, blnHasForExThis As Boolean,
' **   lngStmts As Long, arr_varStmt As Variant,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkStmt_After_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngX As Long

1510    With frm

1520      DoCmd.Hourglass True
1530      DoEvents

1540      .HasForeign_Sched = "NOT CHECKED"
1550      .ForEx_ChkScheduled_lbl.Visible = False

1560      If .HasForeign = True Then
1570        blnHasForEx = True
1580        Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
1590          Select Case IsNull(.cmbAccounts)
              Case True
1600            .chkIncludeCurrency = False
1610            .chkIncludeCurrency.Enabled = False
1620            .chkIncludeCurrency.Locked = False
1630            .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
1640          Case False
1650            .chkIncludeCurrency.Enabled = True
1660            gblnHasForExThis = False
1670            blnHasForExThis = False
1680            For lngX = 0& To (lngAcctFors - 1&)
1690              If arr_varAcctFor(F_ACTNO, lngX) = .cmbAccounts Then
1700                If arr_varAcctFor(F_JCNT, lngX) > 0 Or arr_varAcctFor(F_ACNT, lngX) > 0 Then
1710                  gblnHasForExThis = True
1720                  blnHasForExThis = True
1730                End If
1740                Exit For
1750              End If
1760            Next
1770            Select Case blnHasForExThis
                Case True
1780              .chkIncludeCurrency = True
1790              .chkIncludeCurrency.Locked = True
1800              .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
1810            Case False
1820              .chkIncludeCurrency.Locked = False
1830              .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
1840            End Select
1850          End Select
1860        Case .opgAccountNumber_optAll.OptionValue
1870          .chkIncludeCurrency = True
1880          .chkIncludeCurrency.Enabled = True
1890          .chkIncludeCurrency.Locked = True
1900          .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
1910        End Select
1920      End If

1930      Set dbs = CurrentDb
1940      With dbs
            ' ** Empty tmpAssetList1.
1950        Set qdf = .QueryDefs("qryStatementParameters_AssetList_09b")
1960        qdf.Execute
1970        Set qdf = Nothing
            ' ** Empty tmpAssetList2.
1980        Set qdf = .QueryDefs("qryStatementParameters_AssetList_09c")
1990        qdf.Execute
2000        Set qdf = Nothing
            ' ** Empty tmpAssetList4.
2010        Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_51")
2020        qdf.Execute
2030        Set qdf = Nothing
            ' ** Empty tmpAssetList5.
2040        Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_52")
2050        qdf.Execute
2060        Set qdf = Nothing
2070        .Close
2080      End With
2090      Set dbs = Nothing

2100      SetArchiveOption_SP .chkTransactions, .chkAssetList, lngAcctArchs, arr_varAcctArch, frm  ' ** Module Procedure: modStatementParamFuncs2.

2110      DoCmd.Hourglass True  ' ** Assure it's still going.
2120      DoEvents

2130      Btn_Enable_SP 3, blnIsOpen, blnRunPriorStatement, blnAcctNotSched, datAssetListDate_Pref, lngStmts, arr_varStmt, frm  ' ** Procedure: Below.

2140      DoEvents
2150      DoCmd.Hourglass False

2160    End With

EXITP:
2170    Set qdf = Nothing
2180    Set dbs = Nothing
2190    Exit Sub

ERRH:
2200    DoCmd.Hourglass False
2210    THAT_PROC = THIS_PROC
2220    That_Erl = Erl: That_Desc = ERR.description
2230    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2240    Resume EXITP

End Sub

Public Sub CmdBalTbl_Click_SP(blnContinue2 As Boolean, blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdBalTbl_Click_SP(
' **   blnContinue2 As Boolean, blnGoingToReport As Boolean, blnGTR_Emblem As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdBalTbl_Click_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strAccountNo As String, strDocName As String
        Dim lngRecs As Long
        Dim lngX As Long

2310    With frm

2320      DoCmd.Hourglass True
2330      DoEvents

2340      Set dbs = CurrentDb

          ' ** Empty tmpAccount.
2350      Set qdf = dbs.QueryDefs("qryStatementBalance_11_01")
2360      qdf.Execute
2370      Set qdf = Nothing
2380      DoEvents

          ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
          ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
          ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
          ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
          ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
          ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
          ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
          ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
2390      Set qdf = dbs.QueryDefs("qryStatementBalance_11_02")
2400      qdf.Execute
2410      Set qdf = Nothing
2420      DoEvents

          ' ** Empty tmpBalance.
2430      Set qdf = dbs.QueryDefs("qryStatementBalance_02")
2440      qdf.Execute
2450      Set qdf = Nothing

2460      Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
2470        If IsNull(.cmbAccounts) = False Then
2480          If Trim(.cmbAccounts) <> vbNullString Then
2490            strAccountNo = .cmbAccounts
                ' ** Append qryStatementBalance_03c (Balance table, with graphics) to tmpBalance, by specified [actno].
2500            Set qdf = dbs.QueryDefs("qryStatementBalance_04")
2510            With qdf.Parameters
2520              ![actno] = strAccountNo
2530            End With
2540            qdf.Execute
2550            Set qdf = Nothing
                ' ** Query Source; tblForm_Graphics, just frmStatementBalance.
2560            Set qdf = dbs.QueryDefs("qryStatementBalance_06b")
2570            Set rst1 = qdf.OpenRecordset
2580            Set rst2 = dbs.OpenRecordset("tmpBalance", dbOpenDynaset, dbConsistent)
2590            With rst2
2600              If .BOF = True And .EOF = True Then
                    ' ** Shouldn't even be here!
2610              Else
2620                .MoveLast
2630                lngRecs = .RecordCount
2640                .MoveFirst
2650                For lngX = 1& To lngRecs
2660                  .Edit
2670                  ![frmgfx_id] = rst1![frmgfx_id]
2680                  ![dbs_id] = rst1![dbs_id]
2690                  ![dbs_name] = rst1![dbs_name]
2700                  ![frm_id] = rst1![frm_id]
2710                  ![frm_name] = rst1![frm_name]
2720                  ![ctl_name_01] = rst1![ctl_name_01]
2730                  ![xadgfx_image_01] = rst1![xadgfx_image_01]
2740                  ![ctl_name_02] = rst1![ctl_name_02]
2750                  ![xadgfx_image_02] = rst1![xadgfx_image_02]
2760                  ![ctl_name_03] = rst1![ctl_name_03]
2770                  ![xadgfx_image_03] = rst1![xadgfx_image_03]
2780                  ![ctl_name_04] = rst1![ctl_name_04]
2790                  ![xadgfx_image_04] = rst1![xadgfx_image_04]
2800                  ![ctl_name_05] = rst1![ctl_name_05]
2810                  ![xadgfx_image_05] = rst1![xadgfx_image_05]
2820                  ![ctl_name_06] = rst1![ctl_name_06]
2830                  ![xadgfx_image_06] = rst1![xadgfx_image_06]
2840                  ![ctl_name_07] = rst1![ctl_name_07]
2850                  ![xadgfx_image_07] = rst1![xadgfx_image_07]
2860                  .Update
2870                  If lngX < lngRecs Then .MoveNext
2880                Next
2890              End If
2900              .Close
2910            End With
2920            rst1.Close
2930            Set rst1 = Nothing
2940            Set rst2 = Nothing
2950            Set qdf = Nothing
2960          Else
2970            blnContinue2 = False
2980            DoCmd.Hourglass False
2990            MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "G01")
3000          End If
3010        Else
3020          blnContinue2 = False
3030          DoCmd.Hourglass False
3040          MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "G02")
3050        End If
3060      Case .opgAccountNumber_optAll.OptionValue
3070        strAccountNo = vbNullString
3080        blnContinue2 = False
3090        DoCmd.Hourglass False
3100        MsgBox "The Edit Account Balance option is only available for single, specified accounts.", _
              vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "G03")
3110      End Select

3120      dbs.Close

3130      If blnContinue2 = True Then

3140        gblnSetFocus = True
3150        strDocName = "frmStatementBalance"
3160        DoCmd.OpenForm strDocName, , , , , , frm.Name & "~" & strAccountNo

3170        If blnGoingToReport = True Then
3180          Forms(strDocName).TimerInterval = 50&
3190          .TimerInterval = 0&
3200          .GoToReport_arw_sp_sbal_img.Visible = False
3210          blnGoingToReport = False
3220          blnGTR_Emblem = False
3230          .GTREmblem_Off  ' ** Form Procedure: frmStatementParameters.
3240        End If

3250      Else
3260        DoCmd.Hourglass False
3270      End If

3280    End With

EXITP:
3290    Set rst1 = Nothing
3300    Set rst2 = Nothing
3310    Set qdf = Nothing
3320    Set dbs = Nothing
3330    Exit Sub

ERRH:
3340    DoCmd.Hourglass False
3350    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
3360    Case Else
3370      THAT_PROC = THIS_PROC
3380      That_Erl = Erl: That_Desc = ERR.description
3390      frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3400    End Select
3410    Resume EXITP

End Sub

Public Sub OpgActNo_After_SP(blnResetStmts As Boolean, blnRunPriorStatement As Boolean, strRememberMe As String, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctArchs As Long, arr_varAcctArch As Variant, lngAcctFors As Long, arr_varAcctFor As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** OpgActNo_After_SP(
' **   blnResetStmts As Boolean, blnRunPriorStatement As Boolean, strRememberMe As String,
' **   blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctArchs As Long, arr_varAcctArch As Variant,
' **   lngAcctFors As Long, arr_varAcctFor As Variant,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "OpgActNo_After_SP"

        Dim blnTmp04 As Boolean
        Dim lngX As Long

3510    With frm

3520      DoCmd.Hourglass True
3530      DoEvents

3540      .GenericCurrencySymbol_lbl.Visible = False
3550      blnResetStmts = False

3560      Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
3570        .opgAccountNumber_optSpecified_lbl_box.Visible = True
3580        .opgAccountNumber_optAll_lbl_box.Visible = False
3590        .opgAccountSource.Enabled = True
3600        .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
3610        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
3620        .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
3630        .opgAccountSource_optName_lbl2_dim_hi.Visible = False
3640        .chkRememberMe.Enabled = True
3650        .chkRememberMe_lbl.Visible = True
3660        .chkRememberMe_lbl2_dim.Visible = False
3670        .chkRememberMe_lbl2_dim_hi.Visible = False
3680        .cmbAccounts.Enabled = True
3690        .cmbAccounts.BorderColor = CLR_LTBLU2
3700        .cmbAccounts.BackStyle = acBackStyleNormal
3710        .cmbAccounts.ForeColor = CLR_BLK
3720        .cmbAccounts.BackColor = CLR_WHT
3730        .cmbAccounts_lbl.ForeColor = CLR_WHT
3740        .cmbAccounts_lbl.BackStyle = acBackStyleNormal
3750        .cmbAccounts_lbl_box.Visible = False
3760        .cmbAccounts_lbl2.Visible = .chkStatements
3770        .cmbAccounts_lbl3.Caption = vbNullString
3780        .cmbAccounts_lbl3.Visible = .chkStatements
3790        .cmdBalanceTable.Enabled = True
3800        .cmdBalanceTable_raised_img.Visible = True
3810        .cmdBalanceTable_raised_semifocus_dots_img.Visible = False
3820        .cmdBalanceTable_raised_focus_img.Visible = False
3830        .cmdBalanceTable_raised_focus_dots_img.Visible = False
3840        .cmdBalanceTable_sunken_focus_dots_img.Visible = False
3850        .cmdBalanceTable_raised_img_dis.Visible = False
3860        DoEvents
3870        If .chkRememberMe = True And strRememberMe <> vbNullString Then
3880          .cmbAccounts = strRememberMe
3890          SetAccountLastStatement frm  ' ** Procedure: Below.
3900        End If
3910      Case .opgAccountNumber_optAll.OptionValue
3920        .opgAccountNumber_optSpecified_lbl_box.Visible = False
3930        .opgAccountNumber_optAll_lbl_box.Visible = True
3940        .opgAccountSource.Enabled = False
3950        .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
3960        .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
3970        .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
3980        .opgAccountSource_optName_lbl2_dim_hi.Visible = True
3990        .chkRememberMe.Enabled = False
4000        .chkRememberMe_lbl.Visible = False
4010        .chkRememberMe_lbl2_dim.Visible = True
4020        .chkRememberMe_lbl2_dim_hi.Visible = True
4030        .cmbAccounts = Null
4040        .cmbAccounts.Enabled = False
4050        .cmbAccounts.BorderColor = WIN_CLR_DISR
4060        .cmbAccounts.BackStyle = acBackStyleTransparent
4070        .cmbAccounts_lbl.BackStyle = acBackStyleTransparent  ' ** When you do this, it often leaves a partial blue border visible,
4080        .cmbAccounts_lbl_box.Visible = True                  ' ** which I've covered with the cmbAccounts_lbl_box.
4090        .cmbAccounts_lbl2.Visible = False
4100        .cmbAccounts_lbl3.Visible = False
4110        .cmdBalanceTable.Enabled = False
4120        .cmdBalanceTable_raised_img_dis.Visible = True
4130        .cmdBalanceTable_raised_img.Visible = False
4140        .cmdBalanceTable_raised_semifocus_dots_img.Visible = False
4150        .cmdBalanceTable_raised_focus_img.Visible = False
4160        .cmdBalanceTable_raised_focus_dots_img.Visible = False
4170        .cmdBalanceTable_sunken_focus_dots_img.Visible = False
4180        DoEvents
4190      End Select
4200      .chkStatements_lbl3.Caption = Format(DLookup("[Statement_Date]", "Statement Date"), "mm/dd/yyyy")  ' ** Default.
4210      SetArchiveOption_SP .chkTransactions, .chkAssetList, lngAcctArchs, arr_varAcctArch, frm  ' ** Module Procedure: modStatementParamFuncs2.

4220      If .HasForeign = True Then
4230        blnHasForEx = True
4240        blnTmp04 = True
4250        If .chkStatements = True Then
              ' ** This only applies when dealing with scheduled accounts.
4260          If IsNull(.HasForeign_Sched) = True Then
                ' ** Shouldn't be; it has a DefaultValue!
4270          Else
4280            If .HasForeign_Sched = vbNullString Then
                  ' ** Shouldn't be; it has a DefaultValue!
4290            Else
4300              If .HasForeign_Sched = "NOT CHECKED" Then
                    ' ** Proceed below.
4310              ElseIf .HasForeign_Sched = "SOME" Then
                    ' ** Proceed below.
4320              ElseIf .HasForeign_Sched = "NONE" Then
                    ' ** No scheduled accounts have foreign currency.
4330                blnTmp04 = False
4340              End If
4350            End If
4360          End If
4370        End If  ' ** chkStatements.

4380        If blnTmp04 = True Then
4390          Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
4400            Select Case IsNull(.cmbAccounts)
                Case True
4410              .chkIncludeCurrency = False
4420              .chkIncludeCurrency.Enabled = False
4430              .chkIncludeCurrency.Locked = False
4440              .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
4450            Case False
4460              .chkIncludeCurrency.Enabled = True
4470              gblnHasForExThis = False
4480              blnHasForExThis = False
4490              For lngX = 0& To (lngAcctFors - 1&)
4500                If arr_varAcctFor(F_ACTNO, lngX) = .cmbAccounts Then
4510                  If arr_varAcctFor(F_JCNT, lngX) > 0 Or arr_varAcctFor(F_ACNT, lngX) > 0 Then
4520                    gblnHasForExThis = True
4530                    blnHasForExThis = True
4540                  End If
4550                  Exit For
4560                End If
4570              Next
4580              Select Case blnHasForExThis
                  Case True
4590                .chkIncludeCurrency = True
4600                .chkIncludeCurrency.Locked = True
4610                .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
4620              Case False
4630                .chkIncludeCurrency.Locked = False
4640                .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
4650              End Select
4660            End Select
4670          Case .opgAccountNumber_optAll.OptionValue
4680            .chkIncludeCurrency = True
4690            .chkIncludeCurrency.Enabled = True
4700            .chkIncludeCurrency.Locked = True
4710            .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
4720          End Select
4730        End If  ' ** blnTmp04.
4740      End If

4750      If .chkStatements = True And .opgAccountNumber = .opgAccountNumber_optAll.OptionValue Then
4760        If .cmdTransactionsPreview.Enabled = False And .cmdAssetListPreview.Enabled = False Then
              ' ** User went from All to Specified and chose a non-scheduled account, then back to All.
4770          .chkStatements.SetFocus
4780          blnResetStmts = True
4790        End If
4800      End If

4810      SetStatementOptions frm, blnRunPriorStatement  ' ** Procedure: Below.

4820      If blnResetStmts = True Then
4830        .TimerInterval = 100&
4840      End If

4850      DoCmd.Hourglass False

4860    End With

EXITP:
4870    Exit Sub

ERRH:
4880    DoCmd.Hourglass False
4890    THAT_PROC = THIS_PROC
4900    That_Erl = Erl: That_Desc = ERR.description
4910    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4920    Resume EXITP

End Sub

Public Sub CmbAccts_After_SP(blnContinue As Boolean, blnAcctNotSched As Boolean, blnRunPriorStatement As Boolean, strRememberMe As String, blnAfterFired As Boolean, lngStmts As Long, arr_varStmt As Variant, lngStmtCnt As Long, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctFors As Long, arr_varAcctFor As Variant, lngAcctArchs As Long, arr_varAcctArch As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmbAccts_After_SP(
' **   blnContinue As Boolean, blnAcctNotSched As Boolean, blnRunPriorStatement As Boolean,
' **   strRememberMe As String, blnAfterFired As Boolean, lngStmts As Long, arr_varStmt As Variant,
' **   lngStmtCnt As Long, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctFors As Long,
' **   arr_varAcctFor As Variant, lngAcctArchs As Long, arr_varAcctArch As Variant,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "CmbAccts_After_SP"

        Dim strStmtFld As String, strAccountNo As String, strMsg As String
        Dim datLastStmt_ThisAcct As Date, datLastStmt_AllAccts As Date, datStmtRequested As Date
        Dim lngMonthID As Long
        Dim lngX As Long
        Dim varTmp00 As Variant

        ' ** cmbMonth combo box constants:
        Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
        Const CBX_MON_NAME  As Integer = 1  ' ** month_name
        Const CBX_MON_SHORT As Integer = 2  ' ** month_short

        ' ** Array: arr_varStmt().
        'Const S_ELEMS1 As Integer = 12  ' ** Array's first-element UBound().
        'Const S_ELEMS2 As Integer = 4   ' ** Array's second-element UBound().
        'Const S_MID    As Integer = 0   ' ** month_id
        'Const S_MSHT   As Integer = 1   ' ** month_short
        Const S_CNT    As Integer = 2   ' ** cnt_smt
        Const S_ACTNO  As Integer = 3   ' ** accountno
        'Const S_SNAM   As Integer = 4   ' ** shortname

5010    With frm

5020      If IsNull(.cmbAccounts) = False Then

5030        DoCmd.Hourglass True
5040        DoEvents

5050        blnAcctNotSched = False

5060        strAccountNo = .cmbAccounts
5070        gstrAccountNo = strAccountNo
5080        .GenericCurrencySymbol_lbl.Visible = False

5090        If IsEmpty(arr_varStmt) = True Then
5100          arr_varStmt = AcctSched_Load  ' ** Module Function: modStatementParamFuncs1.
              'Runs qrys to load arrays.
5110          lngStmts = UBound(arr_varStmt, 1)
5120        End If

5130        If .chkStatements = True Then
5140          If IsNull(.cmbMonth) = True Then
5150            .cmbMonth = "December"  ' ** Default to December.
5160            .cmbMonth_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
                'Calls Month_AfterUpdate_SP()
5170            DoEvents
5180          End If
5190          lngMonthID = .cmbMonth.Column(CBX_MON_ID)
5200          lngStmtCnt = arr_varStmt(lngMonthID, S_CNT, 0)
5210          If lngStmtCnt = 0& Then
5220            blnContinue = False
5230          Else
5240            blnContinue = False
5250            For lngX = 0& To (lngStmtCnt - 1&)
5260              If arr_varStmt(lngMonthID, S_ACTNO, lngX) = strAccountNo Then
5270                blnContinue = True
5280                Exit For
5290              End If
5300            Next
5310          End If
5320        End If

5330        If blnContinue = False Then
5340          If lngStmts = 0& Then
5350            strMsg = "There are no accounts scheduled for statements in " & .cmbMonth.Column(CBX_MON_NAME) & "."
5360          Else
5370            strMsg = "This account is not scheduled for a statement in " & .cmbMonth.Column(CBX_MON_NAME) & "."
5380          End If
5390          Beep
5400          DoCmd.Hourglass False
5410          MsgBox strMsg, vbInformation + vbOKOnly, "Nothing To Do"
5420        End If  ' ** blnContinue.

5430        DoCmd.Hourglass True
5440        DoEvents

5450        If blnContinue = True Then

5460          If .HasForeign = True Then
5470            blnHasForEx = True
5480            Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
5490              .chkIncludeCurrency.Enabled = True
5500              gblnHasForExThis = False
5510              blnHasForExThis = False
5520              For lngX = 0& To (lngAcctFors - 1&)
5530                If arr_varAcctFor(F_ACTNO, lngX) = strAccountNo Then
5540                  If arr_varAcctFor(F_JCNT, lngX) > 0 Or arr_varAcctFor(F_ACNT, lngX) > 0 Then
5550                    gblnHasForExThis = True
5560                    blnHasForExThis = True
5570                    .GenericCurrencySymbol_lbl.Visible = True
5580                  End If
5590                  Exit For
5600                End If
5610              Next
5620              Select Case blnHasForExThis
                  Case True
5630                .chkIncludeCurrency = True
5640                .chkIncludeCurrency.Locked = True
5650                .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
5660              Case False
5670                .chkIncludeCurrency.Locked = False
5680                .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
5690              End Select
5700            Case .opgAccountNumber_optAll.OptionValue
5710              .chkIncludeCurrency = True
5720              .chkIncludeCurrency.Enabled = True
5730              .chkIncludeCurrency.Locked = True
5740              .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
5750            End Select
5760          End If

5770          blnRunPriorStatement = False

5780          If .chkTransactions = True Then
5790            SetArchiveOption_SP .chkTransactions, .chkAssetList, lngAcctArchs, arr_varAcctArch, frm  ' ** Module Procedure: modStatementParamFuncs2.
                'Calls GetTPP()
5800          ElseIf .chkAssetList = True Then
5810            SetArchiveOption_SP .chkTransactions, .chkAssetList, lngAcctArchs, arr_varAcctArch, frm  ' ** Module Procedure: modStatementParamFuncs2.
5820          ElseIf .chkStatements = True Then
5830            SetArchiveOption_SP .chkTransactions, .chkAssetList, lngAcctArchs, arr_varAcctArch, frm  ' ** Module Procedure: modStatementParamFuncs2.
5840            SetAccountLastStatement frm  ' ** Procedure: Below.
                'Runs Qrys.
5850            If Trim(.cmbAccounts_lbl3.Caption) <> vbNullString Then
5860              datLastStmt_ThisAcct = CDate(.cmbAccounts_lbl3.Caption)
5870            End If
5880            varTmp00 = DLookup("[Statement_Date]", "Statement Date")
5890            If IsNull(varTmp00) = False Then
5900              datLastStmt_AllAccts = CDate(varTmp00)
5910            End If
5920            If IsNull(.StatementsYear) = False And IsNull(.cmbMonth.Column(CBX_MON_ID)) = False Then
5930              datStmtRequested = DateSerial(CLng(.StatementsYear), CLng(.cmbMonth.Column(CBX_MON_ID)), 31)
5940              strStmtFld = "smt" & .cmbMonth.Column(CBX_MON_SHORT)
5950            End If
5960            varTmp00 = DLookup("[" & strStmtFld & "]", "account", "[accountno] = '" & strAccountNo & "'")
5970            If IsNull(varTmp00) = False Then
5980              If CBool(varTmp00) = True Then
5990                If (datLastStmt_ThisAcct < datLastStmt_AllAccts) And (datStmtRequested = datLastStmt_AllAccts) Then
                      ' ** Missed one!
6000                  blnRunPriorStatement = True
6010                  .cmdPrintStatement_Single.Caption = "Print Single Statement"
6020                  .cmdPrintStatement_Single.ControlTipText = "Print Single" & vbCrLf & "Statement - Ctrl+S"
6030                  .cmdPrintStatement_Single.StatusBarText = "Print Single Statement - Ctrl+S"
6040                  .cmdPrintStatement_Summary.Enabled = False
6050                End If
6060              End If
6070            End If
6080            If blnRunPriorStatement = False Then
6090              If .cmdPrintStatement_Single.Caption = "Print Single Statement" Then
6100                .cmdPrintStatement_Single.Caption = "Reprint Single Statement"
6110                .cmdPrintStatement_Single.ControlTipText = "Reprint Single" & vbCrLf & "Statement - Ctrl+S"
6120                .cmdPrintStatement_Single.StatusBarText = "Reprint Single Statement - Ctrl+S"
6130              End If
6140              If .cmdPrintStatement_Summary.Enabled = False Then
6150                .cmdPrintStatement_Summary.Enabled = True
6160              End If
6170            End If
6180          End If

6190          Select Case .chkRememberMe
              Case True
6200            strRememberMe = strAccountNo
6210          Case False
6220            strRememberMe = vbNullString
6230          End Select

6240          SetRelatedOption frm  ' ** Module Procedure: modStatementParamFuncs2.
              'Sets Ctls.

6250        End If  ' ** blnContinue.

6260        Select Case blnContinue
            Case True
6270          If .chkTransactions = True Then
6280            .cmdTransactionsPreview.Enabled = True
6290            .cmdTransactionsPrint.Enabled = True
6300            .cmdTransactionsWord.Enabled = True
6310            .cmdTransactionsExcel.Enabled = True
6320          ElseIf .chkAssetList = True Then
6330            .cmdAssetListPreview.Enabled = True
6340            .cmdAssetListPrint.Enabled = True
6350            .cmdAssetListWord.Enabled = True
6360            .cmdAssetListExcel.Enabled = True
6370          Else
6380            .cmdTransactionsPreview.Enabled = True
6390            .cmdTransactionsPrint.Enabled = True
6400            .cmdTransactionsWord.Enabled = True
6410            .cmdTransactionsExcel.Enabled = True
6420            .cmdAssetListPreview.Enabled = True
6430            .cmdAssetListPrint.Enabled = True
6440            .cmdAssetListWord.Enabled = True
6450            .cmdAssetListExcel.Enabled = True
6460            .cmdPrintStatement_Single.Enabled = True
6470            .cmdPrintStatement_Summary.Enabled = True
6480          End If
6490          blnAcctNotSched = False
6500        Case False
6510          .cmdTransactionsPreview.Enabled = False
6520          .cmdTransactionsPrint.Enabled = False
6530          .cmdTransactionsWord.Enabled = False
6540          .cmdTransactionsExcel.Enabled = False
6550          .cmdAssetListPreview.Enabled = False
6560          .cmdAssetListPrint.Enabled = False
6570          .cmdAssetListWord.Enabled = False
6580          .cmdAssetListExcel.Enabled = False
6590          .cmdPrintStatement_Single.Enabled = False
6600          .cmdPrintStatement_Summary.Enabled = False
6610          blnAcctNotSched = True
6620        End Select

6630        blnAfterFired = True

6640        DoCmd.Hourglass False

6650      End If

6660    End With

EXITP:
6670    Exit Sub

ERRH:
6680    DoCmd.Hourglass False
6690    THAT_PROC = THIS_PROC
6700    That_Erl = Erl: That_Desc = ERR.description
6710    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6720    Resume EXITP

End Sub

Public Sub CmdTransPreview_Click_SP(blnContinue As Boolean, blnAllStatements As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String, strFileName As String, strReportName As String, blnHasForEx As Boolean, blnHasForExThis As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdTransPreview_Click_SP(
' **   blnContinue As Boolean, blnAllStatements As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String,
' **   strFileName As String, strReportName As String, blnHasForEx As Boolean, blnHasForExThis As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdTransPreview_Click_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

        ' ** Array: arr_varDoc().
        'Const D_ACTNO As Integer = 0
        'Const D_MONID As Integer = 1

6810    With frm

6820      blnRetVal = True

6830      If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
6840        blnRetVal = False
6850        MsgBox "You must select an account to continue," & vbCrLf & _
              "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "I01")
6860      Else
6870        If .chkStatements = True Then
6880          If IsNull(.cmbMonth) Then
6890            blnRetVal = False
6900            MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                  (Left(("Entry Required" & Space(55)), 55) & "I02")
6910            .cmbMonth.SetFocus
6920          Else
6930            If .cmbMonth = vbNullString Then
6940              blnRetVal = False
6950              MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                    (Left(("Entry Required" & Space(55)), 55) & "I03")
6960              .cmbMonth.SetFocus
6970            Else
6980              If IsNull(.StatementsYear) = True Then
6990                blnRetVal = False
7000                MsgBox "You must enter a report year to continue.", vbInformation + vbOKOnly, _
                      (Left(("Entry Required" & Space(55)), 55) & "I04")
7010                .StatementsYear.SetFocus
7020              End If
7030            End If
7040          End If
7050        Else
7060          If FirstDate_SP(frm) = False Then  ' ** Function: Below.
7070            blnRetVal = False
7080            MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "I05")
7090          End If
7100        End If

7110        If blnRetVal = True Then

7120          DoCmd.Hourglass True
7130          DoEvents

7140          If .opgAccountNumber = .opgAccountNumber_optAll.OptionValue Then
7150            blnAllStatements = True
7160          End If

              'If blnFromStmts = False Then
              'AND IT'S A SINGLE TRANS RPT, THEN SHOW THE 'NO DATA' MSG!

              ' ** Execute the common code.
7170          blnRetVal = BuildTransactionInfo_SP(frm, strFileName, strReportName, blnAllStatements, blnContinue, blnHasForEx, blnHasForExThis, blnFromStmts)  ' ** Module Function: modStatementParamFuncs1.
7180          If blnContinue = True And blnRetVal = True Then
7190            DoCmd.Maximize
7200            DoCmd.RunCommand acCmdFitToWindow
7210          End If

7220        End If  ' ** blnRetVal.
7230      End If

7240      DoCmd.Hourglass False

7250    End With

EXITP:
7260    Set rst = Nothing
7270    Set qdf = Nothing
7280    Set dbs = Nothing
7290    Exit Sub

ERRH:
7300    DoCmd.Hourglass False
7310    THAT_PROC = THIS_PROC
7320    That_Erl = Erl: That_Desc = ERR.description
7330    frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7340    Resume EXITP

End Sub

Public Sub CmdTransPrint_Click_SP(blnContinue As Boolean, blnAllStatements As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String, strFileName As String, strReportName As String, blnHasForEx As Boolean, blnHasForExThis As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdTransPrint_Click_SP(
' **   blnContinue As Boolean, blnAllStatements As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String,
' **   strFileName As String, strReportName As String, blnHasForEx As Boolean, blnHasForExThis As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdTransPrint_Click_SP"

        Dim strSQL As String
        Dim blnRetVal As Boolean

7410    With frm

7420      blnRetVal = True

7430      If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
7440        MsgBox "You must select an account to continue," & vbCrLf & _
              "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "J01")
7450      Else

7460        If .chkStatements = True Then
7470          If IsNull(.cmbMonth) Then
7480            blnRetVal = False
7490            MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                  (Left(("Entry Required" & Space(55)), 55) & "J02")
7500            .cmbMonth.SetFocus
7510          Else
7520            If .cmbMonth = vbNullString Then
7530              blnRetVal = False
7540              MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                    (Left(("Entry Required" & Space(55)), 55) & "J03")
7550              .cmbMonth.SetFocus
7560            Else
7570              If IsNull(.StatementsYear) = True Then
7580                blnRetVal = False
7590                MsgBox "You must enter a report year to continue.", vbInformation + vbOKOnly, _
                      (Left(("Entry Required" & Space(55)), 55) & "J04")
7600                .StatementsYear.SetFocus
7610              End If
7620            End If
7630          End If
7640        Else
7650          If FirstDate_SP(frm) = False Then  ' ** Function: Below.
7660            blnRetVal = False
7670            MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "J05")
7680          End If
7690        End If

7700        If blnRetVal = True Then

7710          DoCmd.Hourglass True
7720          DoEvents

              ' ** Execute the common code.
7730          blnRetVal = BuildTransactionInfo_SP(frm, strFileName, strReportName, blnAllStatements, blnContinue, blnHasForEx, blnHasForExThis, blnFromStmts)  ' ** Module Function: modStatementParamFuncs1.

7740          If blnContinue = True And blnRetVal = True Then
7750            DoCmd.SetWarnings False
7760            If gblnDev_Debug = True Or (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
7770              If IsLoaded(strReportName, acReport) = True Then  ' ** Module Function: modFileUtilities.
7780                DoCmd.Close acReport, strReportName
7790              End If
7800              DoCmd.OpenReport strReportName, acViewPreview
7810            Else  ' ** Normal.
7820              If IsLoaded(strReportName, acReport) = False Then  ' ** Module Function: modFileUtilities.
7830                DoCmd.OpenReport strReportName, acViewPreview  ' ** This gets the caption changed!
7840              End If
                  '##GTR_Ref: rptTransaction_Statement_SortDate
                  '##GTR_Ref: rptTransaction_Statement_SortType
                  '##GTR_Ref: rptTransaction_Statement_ForEx_SortDate
                  '##GTR_Ref: rptTransaction_Statement_ForEx_SortType
7850              DoCmd.OpenReport strReportName, acViewNormal
7860            End If
7870            DoCmd.SetWarnings True
                'NOW MAKE SURE IT'S CLOSED!
7880          Else
7890            If blnContinue = True And blnRetVal = False And .chkStatements = True Then
                  ' ** Must have been no transactions when printing statements.
7900              strSQL = "SELECT DISTINCTROW account.shortname, account.legalname, account.accountno," & CoInfo & " " & _
                    "FROM account " & _
                    "WHERE account.accountno = '" & Trim(.cmbAccounts) & "';"
7910              strSQL = StringReplace(strSQL, "'' As ", "Null As ")  ' ** Module Function: modStringFuncs.
7920              CurrentDb.QueryDefs("qryStatementParameters_19").SQL = strSQL
7930              DoCmd.SetWarnings False
7940              Select Case blnHasForEx
                  Case True
7950                Select Case blnHasForExThis
                    Case True
7960                  strReportName = "rptTransaction_Statement_ForEx_NoData"
7970                Case False
7980                  Select Case .chkIncludeCurrency
                      Case True
7990                    strReportName = "rptTransaction_Statement_ForEx_NoData"
8000                  Case False
8010                    strReportName = "rptTransaction_Statement_NoData"
8020                  End Select
8030                End Select
8040              Case False
8050                strReportName = "rptTransaction_Statement_NoData"
8060              End Select
8070              If gblnDev_Debug = True Or (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
8080                If IsLoaded(strReportName, acReport) = True Then  ' ** Module Function: modFileUtilities.
8090                  DoCmd.Close acReport, strReportName
8100                End If
8110                DoCmd.OpenReport strReportName, acViewPreview, , , , Trim(.cmbAccounts)
8120              Else
8130                If IsLoaded(strReportName, acReport) = False Then  ' ** Module Function: modFileUtilities.
8140                  DoCmd.OpenReport strReportName, acViewPreview, , , , Trim(.cmbAccounts)  ' ** This gets the caption changed!
8150                End If
8160                DoCmd.OpenReport strReportName, acViewNormal, , , , Trim(.cmbAccounts)
8170              End If
8180              DoCmd.SetWarnings True
                  'NOW MAKE SURE IT'S CLOSED!
8190            End If  ' ** blnContinue, blnRetVal, chkStatements.
8200          End If  ' ** blnContinue, blnRetVal.

8210        End If  ' ** blnRetVal.

8220      End If

8230      DoEvents
8240      If IsLoaded(strReportName, acReport) = True Then
8250        If gblnDev_Debug = False And ((CurrentUser <> "Superuser") Or (CurrentUser = "Superuser" And .chkAsDev = False)) Then
8260          DoCmd.Close acReport, strReportName
8270          DoEvents
8280        End If
8290      End If

8300      DoCmd.Hourglass False

8310    End With

EXITP:
8320    Exit Sub

ERRH:
8330    DoCmd.Hourglass False
8340    DoCmd.SetWarnings True
8350    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
8360      blnContinue = False
8370      If Reports.Count > 0 Then
8380        DoCmd.Close acReport, Reports(0).Name   ' ** Close report in preview.
8390      End If
8400    Case Else
8410      THAT_PROC = THIS_PROC
8420      That_Erl = Erl: That_Desc = ERR.description
8430      frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8440    End Select
8450    Resume EXITP

End Sub

Public Sub CmdTransWord_Click_SP(blnContinue As Boolean, blnAllStatements As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String, strFileName As String, strReportName As String, blnHasForEx As Boolean, blnHasForExThis As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdTransWord_Click_SP(
' **   blnContinue As Boolean, blnAllStatements As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String,
' **   strFileName As String, strReportName As String, blnHasForEx As Boolean, blnHasForExThis As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "CmdTransWord_Click_SP"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim blnRetVal As Boolean

        ' ** cmbAccounts combo box constants:
        Const CBX_A_ACTNO  As Integer = 0  ' ** accountno
        'Const CBX_A_DESC   As Integer = 1  ' ** Desc
        'Const CBX_A_PREDAT As Integer = 2  ' ** predate
        'Const CBX_A_SHORT  As Integer = 3  ' ** shortname
        'Const CBX_A_LEGAL  As Integer = 4  ' ** legalname
        'Const CBX_A_BALDAT As Integer = 5  ' ** BalanceDate (earliest [balance date])
        'Const CBX_A_HASREL As Integer = 6  ' ** HasRelated
        'Const CBX_A_CASNUM As Integer = 7  ' ** CaseNum
        'Const CBX_A_TRXDAT As Integer = 8  ' ** TransDate (earliest [transdate])

8510    With frm

8520      blnRetVal = True: strFileName = vbNullString

8530      If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
8540        MsgBox "You must select an account to continue," & vbCrLf & _
              "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "K01")
8550      Else

8560        If .chkStatements = True Then
8570          If IsNull(.cmbMonth) Then
8580            blnRetVal = False
8590            MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                  (Left(("Entry Required" & Space(55)), 55) & "K02")
8600            .cmbMonth.SetFocus
8610          Else
8620            If .cmbMonth = vbNullString Then
8630              blnRetVal = False
8640              MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                    (Left(("Entry Required" & Space(55)), 55) & "K03")
8650              .cmbMonth.SetFocus
8660            Else
8670              If IsNull(.StatementsYear) = True Then
8680                blnRetVal = False
8690                MsgBox "You must enter a report year to continue.", vbInformation + vbOKOnly, _
                      (Left(("Entry Required" & Space(55)), 55) & "K04")
8700                .StatementsYear.SetFocus
8710              End If
8720            End If
8730          End If
8740        Else
8750          If FirstDate_SP(frm) = False Then  ' ** Function: Below.
8760            blnRetVal = False
8770            MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "K05")
8780          End If
8790        End If

8800        If blnRetVal = True Then

8810          DoCmd.Hourglass True
8820          DoEvents

8830          strRptCap = vbNullString: strRptPathFile = vbNullString

              'If .chkStatements = True, then just those scheduled!

              ' ** Execute the common code.
8840          blnRetVal = BuildTransactionInfo_SP(frm, strFileName, strReportName, blnAllStatements, blnContinue, blnHasForEx, blnHasForExThis, blnFromStmts, "Word")  ' ** Module Function: modStatementParamFuncs1.

8850          If blnRetVal = True Then
8860            If blnContinue = True Then

8870              If IsNull(.UserReportPath) = True Then
8880                strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
8890              Else
8900                strRptPath = .UserReportPath
8910              End If
8920              Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
8930                strRptCap = "rptTransaction_Statement_" & .cmbAccounts.Column(CBX_A_ACTNO) & "_"
8940              Case .opgAccountNumber_optAll.OptionValue
8950                strRptCap = "rptTransaction_Statement_All_"
8960              End Select
8970              strRptCap = StringReplace(strRptCap, "/", "_")  ' ** Module Function: modStringFuncs.
8980              If .chkTransactions = True Then
8990                strRptCap = strRptCap & "_" & Format(.TransDateStart, "yymmdd") & "-" & Format(.TransDateEnd, "yymmdd")
9000              ElseIf .chkStatements = True Then
9010                strRptCap = strRptCap & "_" & Format(.DateStart, "yymmdd") & "-" & Format(.DateEnd, "yymmdd")
9020              End If

9030              If gblnDev_Debug = True Or (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
9040                Select Case .opgOrderBy
                    Case .opgOrderBy_optDate.OptionValue
9050                  Select Case blnHasForEx
                      Case True
9060                    Select Case .opgAccountNumber
                        Case .opgAccountNumber_optSpecified.OptionValue
9070                      Select Case blnHasForExThis
                          Case True
9080                        strRpt = "rptTransaction_Statement_ForEx_SortDate"
9090                      Case False
9100                        Select Case .chkIncludeCurrency
                            Case True
9110                          strRpt = "rptTransaction_Statement_ForEx_SortDate"
9120                        Case False
9130                          strRpt = "rptTransaction_Statement_SortDate"
9140                        End Select
9150                      End Select
9160                    Case .opgAccountNumber_optAll.OptionValue
9170                      strRpt = "rptTransaction_Statement_ForEx_SortDate"
9180                    End Select
9190                  Case False
9200                    strRpt = "rptTransaction_Statement_SortDate"
9210                  End Select
9220                Case .opgOrderBy_optType.OptionValue
9230                  Select Case blnHasForEx
                      Case True
9240                    Select Case .opgAccountNumber
                        Case .opgAccountNumber_optSpecified.OptionValue
9250                      Select Case blnHasForExThis
                          Case True
9260                        strRpt = "rptTransaction_Statement_ForEx_SortType"
9270                      Case False
9280                        Select Case .chkIncludeCurrency
                            Case True
9290                          strRpt = "rptTransaction_Statement_ForEx_SortType"
9300                        Case False
9310                          strRpt = "rptTransaction_Statement_SortType"
9320                        End Select
9330                      End Select
9340                    Case .opgAccountNumber_optAll.OptionValue
9350                      strRpt = "rptTransaction_Statement_ForEx_SortType"
9360                    End Select
9370                  Case False
9380                    strRpt = "rptTransaction_Statement_Sorttype"
9390                  End Select
9400                End Select
9410                DoCmd.OpenReport strRpt, acViewPreview
9420              Else

9430                strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

9440                If strRptPathFile <> vbNullString Then
9450                  Select Case .opgOrderBy
                      Case .opgOrderBy_optDate.OptionValue
9460                    Select Case blnHasForEx
                        Case True
9470                      Select Case .opgAccountNumber
                          Case .opgAccountNumber_optSpecified.OptionValue
9480                        Select Case blnHasForExThis
                            Case True
9490                          strRpt = "rptTransaction_Statement_ForEx_SortDate"
9500                        Case False
9510                          Select Case .chkIncludeCurrency
                              Case True
9520                            strRpt = "rptTransaction_Statement_ForEx_SortDate"
9530                          Case False
9540                            strRpt = "rptTransaction_Statement_SortDate"
9550                          End Select
9560                        End Select
9570                      Case .opgAccountNumber_optAll.OptionValue
9580                        strRpt = "rptTransaction_Statement_ForEx_SortDate"
9590                      End Select
9600                    Case False
9610                      strRpt = "rptTransaction_Statement_SortDate"
9620                    End Select
9630                  Case .opgOrderBy_optType.OptionValue
9640                    Select Case blnHasForEx
                        Case True
9650                      Select Case .opgAccountNumber
                          Case .opgAccountNumber_optSpecified.OptionValue
9660                        Select Case blnHasForExThis
                            Case True
9670                          strRpt = "rptTransaction_Statement_ForEx_SortType"
9680                        Case False
9690                          Select Case .chkIncludeCurrency
                              Case True
9700                            strRpt = "rptTransaction_Statement_ForEx_SortType"
9710                          Case False
9720                            strRpt = "rptTransaction_Statement_SortType"
9730                          End Select
9740                        End Select
9750                      Case .opgAccountNumber_optAll.OptionValue
9760                        strRpt = "rptTransaction_Statement_ForEx_SortType"
9770                      End Select
9780                    Case False
9790                      strRpt = "rptTransaction_Statement_SortType"
9800                    End Select
9810                  End Select
9820                  DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
9830                  .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
9840                End If

9850              End If  ' ** gblnDev_Debug.

9860            End If  ' ** blnContinue.
9870          End If  ' ** blnRetVal.

9880        End If  ' ** blnRetVal.
9890      End If

9900      DoCmd.Hourglass False

9910    End With

EXITP:
9920    Exit Sub

ERRH:
9930    DoCmd.Hourglass False
9940    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
9950      blnContinue = False
9960    Case Else
9970      THAT_PROC = THIS_PROC
9980      That_Erl = Erl: That_Desc = ERR.description
9990      frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10000   End Select
10010   Resume EXITP

End Sub

Public Sub CmdAstListPreview_Click_SP(blnContinue As Boolean, datAssetListDate As Date, blnPrintAnnualStatement As Boolean, blnAllStatements As Boolean, blnNoDataAll As Boolean, blnRollbackNeeded As Boolean, strFirstDateMsg As String, strReportName As String, blnHasForExClick As Boolean, blnIncludeCurrency As Boolean, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctFors As Long, arr_varAcctFor As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdAstListPreview_Click_SP(
' **   blnContinue As Boolean, datAssetListDate As Date, blnPrintAnnualStatement As Boolean,
' **   blnAllStatements As Boolean, blnNoDataAll As Boolean, blnRollbackNeeded As Boolean,
' **   strFirstDateMsg As String, strReportName As String, blnHasForExClick As Boolean,
' **   blnIncludeCurrency As Boolean, blnHasForEx As Boolean, blnHasForExThis As Boolean,
' **   lngAcctFors As Long, arr_varAcctFor As Variant,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "CmdAstListPreview_Click_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intOptGrpAcctNum As Integer
        Dim blnPriceHistory As Boolean
        Dim strDocName As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

10110   With frm

10120     If blnHasForExClick = False Then

10130       blnContinue = True
10140       blnRetVal = True

10150       .chkNoAssets = False
10160       .chkNoAssets_lbl.FontBold = False
10170       .chkNoAssets_All = False
10180       .chkNoAssets_All_lbl.FontBold = False

10190       If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
10200         blnContinue = False
10210         MsgBox "You must select an account to continue," & vbCrLf & _
                "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "M01")
10220       Else

10230         If .chkAssetList = True Then
10240           If FirstDate_SP(frm) = False Then  ' ** Function: Below.
10250             blnRetVal = False
10260             MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "M02")
10270           End If
10280         End If

10290         If blnRetVal = True Then

10300           DoCmd.Hourglass True
10310           DoEvents

10320           intOptGrpAcctNum = .opgAccountNumber
10330           gblnCombineAssets = .chkCombineCash.Value

10340           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
10350             gstrAccountNo = .cmbAccounts
10360           Case .opgAccountNumber_optAll.OptionValue
10370             gstrAccountNo = "All"
10380           End Select
10390           blnIncludeCurrency = .chkIncludeCurrency
10400           datAssetListDate = .AssetListDate

10410           Set dbs = CurrentDb
10420           With dbs
                  ' ** Empty tmpAssetList1.
10430             Set qdf = .QueryDefs("qryStatementParameters_AssetList_09b")
10440             qdf.Execute
10450             Set qdf = Nothing
                  ' ** Empty tmpAssetList2.
10460             Set qdf = .QueryDefs("qryStatementParameters_AssetList_09c")
10470             qdf.Execute
10480             Set qdf = Nothing
                  ' ** Empty tmpAssetList4.
10490             Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_51")
10500             qdf.Execute
10510             Set qdf = Nothing
                  ' ** Empty tmpAssetList5.
10520             Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_52")
10530             qdf.Execute
10540             Set qdf = Nothing
10550             .Close
10560           End With
10570           Set dbs = Nothing

10580           If .chkForeignExchange = True And .chkIncludeCurrency = True Then
10590             .currentDate = Null
                  'blnPriceHistory = PricingHistory(datAssetListDate)  ' ** Function: Below.
                  ' ** blnPriceHistory indicates whether current pricing or pricing history should be used.
                  ' ** It ONLY applies to foreign exchange, since regular reports don't require that info.
                  '.UsePriceHistory = blnPriceHistory
                  ' ** Since UsePriceHistory determines whether foreign currency columns should
                  ' ** be used, and the current pricing should also be in pricing history,
                  ' ** I'm just going to say it should always use pricing history.
10600             blnPriceHistory = True
10610             .UsePriceHistory = blnPriceHistory
10620           End If

10630           If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
10640             If blnHasForEx = True And .UsePriceHistory = True Then
                    ' ** If all the rest of this code is using the foreign currency tables and queries,
                    ' ** but the user unchecked the box because this particular account has no foreign currency,
                    ' ** the report will end up looking in the wrong tables for the data!
10650               Select Case blnIncludeCurrency
                    Case True
10660                 gblnHasForExThis = False
10670                 blnHasForExThis = False
10680                 gblnSwitchTo = False
10690                 For lngX = 0& To (lngAcctFors - 1&)
10700                   If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
10710                     If arr_varAcctFor(F_ACNT, lngX) > 0 Then
10720                       gblnHasForExThis = True
10730                       blnHasForExThis = True
10740                     End If
10750                     gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)
10760                     Exit For
10770                   End If
10780                 Next
10790                 Select Case gblnHasForExThis
                      Case True
10800                   If gblnSwitchTo = True Then
                          ' ** Turn it off since they now do have foreign currencies.
10810                     Set dbs = CurrentDb
10820                     With dbs
                            ' ** Update tblCurrency_Account for curracct_supress = False, by specified [actno].
10830                       Set qdf = .QueryDefs("qryCurrency_17_02")
10840                       With qdf.Parameters
10850                         ![actno] = gstrAccountNo
10860                       End With
10870                       qdf.Execute
10880                       Set qdf = Nothing
10890                       .Close
10900                     End With
10910                     Set dbs = Nothing
10920                   End If
10930                 Case False
10940                   If gblnSwitchTo = True Then
                          ' ** If they've specified to suppress, then turn chkIncludeCurrency off.
10950                     blnIncludeCurrency = False
10960                     .chkIncludeCurrency = False
10970                     .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
10980                   End If
10990                 End Select
11000               Case False
11010                 gblnHasForExThis = False
11020                 blnHasForExThis = False
11030                 gblnSwitchTo = False
11040                 For lngX = 0& To (lngAcctFors - 1&)
11050                   If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
11060                     If arr_varAcctFor(F_ACNT, lngX) > 0 Then
11070                       gblnHasForExThis = True
11080                       blnHasForExThis = True
11090                     End If
11100                     gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)
11110                     Exit For
11120                   End If
11130                 Next
11140                 Select Case gblnHasForExThis
                      Case True
                        ' ** This account does have foreign currencies, and
                        ' ** the user shouldn't have been able to turn it off.
11150                   blnIncludeCurrency = True
11160                   .chkIncludeCurrency = True
11170                   .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
11180                 Case False
11190                   Select Case gblnSwitchTo
                        Case True
11200                     gblnMessage = True
11210                   Case False
11220                     strDocName = "frmStatementParameters_ForEx"
11230                     gblnSetFocus = True
11240                     gblnMessage = True  ' ** False return means cancel.
11250                     gblnSwitchTo = True  ' ** False return means show ForEx, don't supress.
11260                     DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & gstrAccountNo
11270                   End Select
11280                   Select Case gblnMessage
                        Case True
11290                     Select Case gblnSwitchTo
                          Case True
                            ' ** Let blnIncludeCurrency remain False.
11300                       .UsePriceHistory = False
11310                     Case False
                            ' ** Turn it back on then.
11320                       blnIncludeCurrency = True
11330                       .chkIncludeCurrency = True
11340                       .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
11350                     End Select
11360                     gblnMessage = False
11370                     gblnSwitchTo = False
11380                   Case False
                          ' ** Cancel this button.
11390                     blnRetVal = False
11400                     DoCmd.Hourglass False
11410                   End Select
11420                 End Select
11430               End Select
11440             Else
                    ' ** When it's the current date, there's no need to use pricing history,
                    ' ** but we still should ask about suppressing the columns.
11450               If blnHasForEx = True Then

11460               End If
11470             End If
11480           End If

11490           If blnRetVal = True Then
                  ' ** chkNoAssets_All means some accounts have no asset trans or no trans at all.
11500             blnRetVal = BuildAssetListInfo_SP(frm, blnContinue, datAssetListDate, blnPrintAnnualStatement, blnAllStatements, blnNoDataAll, blnRollbackNeeded)  ' ** Module Function: modStatementParamFuncs1.
11510             If blnRetVal = True Then
11520               Select Case blnIncludeCurrency
                    Case True
11530                 strReportName = "rptAssetList_ForEx"
11540               Case False
11550                 strReportName = "rptAssetList"
11560               End Select
11570               If strReportName = "rptAssetList_ForEx" Then
11580                 ForExRptSub_Load frm, blnRollbackNeeded  ' ** Module Procedure: modStatementParamFuncs1.
11590               End If
11600               DoCmd.OpenReport strReportName, acViewPreview
11610               DoCmd.Maximize
11620               DoCmd.RunCommand acCmdFitToWindow
11630               If intOptGrpAcctNum <> .opgAccountNumber Then
11640                 .opgAccountNumber = intOptGrpAcctNum  '#Covered.
11650                 .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
11660               End If
11670             End If  ' ** blnRetVal.
11680           End If  ' ** blnRetVal.

11690         End If
11700       End If

11710     Else
11720       blnHasForExClick = False
11730       DoCmd.Hourglass False
11740     End If

11750   End With

EXITP:
11760   Set qdf = Nothing
11770   Set dbs = Nothing
11780   Exit Sub

ERRH:
11790   DoCmd.Hourglass False
11800   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
11810   Case Else
11820     THAT_PROC = THIS_PROC
11830     That_Erl = Erl: That_Desc = ERR.description
11840     frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11850   End Select
11860   Resume EXITP

End Sub

Public Sub CmdAstListPrint_Click_SP(blnContinue As Boolean, blnFromStmts As Boolean, blnPrintAnnualStatement As Boolean, blnAllStatements As Boolean, blnRollbackNeeded As Boolean, blnNoDataAll As Boolean, datAssetListDate As Date, strFirstDateMsg As String, strReportName As String, blnHasForExClick As Boolean, blnIncludeCurrency As Boolean, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctFors As Long, arr_varAcctFor As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdAstListPrint_Click_SP(
' **   blnContinue As Boolean, blnFromStmts As Boolean, blnPrintAnnualStatement As Boolean,
' **   blnAllStatements As Boolean, blnRollbackNeeded As Boolean, blnNoDataAll As Boolean,
' **   datAssetListDate As Date, strFirstDateMsg As String, strReportName As String,
' **   blnHasForExClick As Boolean, blnIncludeCurrency As Boolean, blnHasForEx As Boolean,
' **   blnHasForExThis As Boolean, lngAcctFors As Long, arr_varAcctFor As Variant,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "CmdAstListPrint_Click_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intOptGrpAcctNum As Integer
        Dim blnPriceHistory As Boolean
        Dim strDocName As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

11910   With frm

11920     If blnHasForExClick = False Then

11930       blnContinue = True
11940       blnRetVal = True
11950       blnNoDataAll = False

11960       If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
11970         blnContinue = False
11980         MsgBox "You must select an account to continue," & vbCrLf & _
                "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "N01")
11990       Else

12000         If .chkAssetList = True Then
12010           If FirstDate_SP(frm) = False Then  ' ** Function: Below.
12020             blnRetVal = False
12030             MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "N02")
12040           End If
12050         End If

12060         If blnRetVal = True Then

12070           intOptGrpAcctNum = .opgAccountNumber
12080           gblnCombineAssets = .chkCombineCash.Value

12090           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
12100             gstrAccountNo = .cmbAccounts
12110           Case .opgAccountNumber_optAll.OptionValue
12120             gstrAccountNo = "All"
12130           End Select
12140           blnIncludeCurrency = .chkIncludeCurrency
12150           datAssetListDate = .AssetListDate

12160           Select Case blnFromStmts
                Case True
12170             If blnIncludeCurrency = True Then
12180               .currentDate = Null
12190               blnPriceHistory = PricingHistory(datAssetListDate)  ' ** Function: Below.
                    ' ** blnPriceHistory indicates whether current pricing or pricing history should be used.
                    ' ** It ONLY applies to foreign exchange, since regular reports don't require that info.
12200               .UsePriceHistory = blnPriceHistory
12210             End If
12220           Case False
12230             If .chkForeignExchange = True And blnIncludeCurrency = True Then
12240               .currentDate = Null
12250               blnPriceHistory = PricingHistory(datAssetListDate)  ' ** Function: Below.
                    ' ** blnPriceHistory indicates whether current pricing or pricing history should be used.
                    ' ** It ONLY applies to foreign exchange, since regular reports don't require that info.
12260               .UsePriceHistory = blnPriceHistory
12270             End If
12280           End Select  ' ** blnFromStmts.

12290           If blnFromStmts = False Then  ' ** Since blnIncludeCurrency has already been set.
12300             If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
12310               If blnHasForEx = True And .UsePriceHistory = True Then
                      ' ** If all the rest of this code is using the foreign currency tables and queries,
                      ' ** but the user unchecked the box because this particular account has no foreign currency,
                      ' ** the report will end up looking in the wrong tables for the data!
12320                 Select Case blnIncludeCurrency
                      Case True
12330                   gblnHasForExThis = False
12340                   blnHasForExThis = False
12350                   gblnSwitchTo = False
12360                   For lngX = 0& To (lngAcctFors - 1&)
12370                     If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
12380                       If arr_varAcctFor(F_ACNT, lngX) > 0 Then
12390                         gblnHasForExThis = True
12400                         blnHasForExThis = True
12410                       End If
12420                       gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)
12430                       Exit For
12440                     End If
12450                   Next
12460                   Select Case gblnHasForExThis
                        Case True
12470                     If gblnSwitchTo = True Then
                            ' ** Turn it off since they now do have foreign currencies.
12480                       Set dbs = CurrentDb
12490                       With dbs
                              ' ** Update tblCurrency_Account for curracct_supress = False, by specified [actno].
12500                         Set qdf = .QueryDefs("qryCurrency_17_02")
12510                         With qdf.Parameters
12520                           ![actno] = gstrAccountNo
12530                         End With
12540                         qdf.Execute
12550                         Set qdf = Nothing
12560                         .Close
12570                       End With
12580                       Set dbs = Nothing
12590                     End If
12600                   Case False
12610                     If gblnSwitchTo = True Then
                            ' ** If they've specified to suppress, then turn chkIncludeCurrency off.
12620                       blnIncludeCurrency = False
12630                       .chkIncludeCurrency = False
12640                       .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
12650                     End If
12660                   End Select
12670                 Case False
12680                   gblnHasForExThis = False
12690                   blnHasForExThis = False
12700                   gblnSwitchTo = False
12710                   For lngX = 0& To (lngAcctFors - 1&)
12720                     If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
12730                       If arr_varAcctFor(F_ACNT, lngX) > 0 Then
12740                         gblnHasForExThis = True
12750                         blnHasForExThis = True
12760                       End If
12770                       gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)
12780                       Exit For
12790                     End If
12800                   Next
12810                   Select Case gblnHasForExThis
                        Case True
                          ' ** This account does have foreign currencies, and
                          ' ** the user shouldn't have been able to turn it off.
12820                     blnIncludeCurrency = True
12830                     .chkIncludeCurrency = True
12840                     .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
12850                   Case False
12860                     Select Case gblnSwitchTo
                          Case True
12870                       gblnMessage = True
12880                     Case False
12890                       strDocName = "frmStatementParameters_ForEx"
12900                       gblnSetFocus = True
12910                       gblnMessage = True  ' ** False return means cancel.
12920                       gblnSwitchTo = True  ' ** False return means show ForEx, don't supress.
12930                       DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & gstrAccountNo
12940                     End Select
12950                     Select Case gblnMessage
                          Case True
12960                       Select Case gblnSwitchTo
                            Case True
                              ' ** Let blnIncludeCurrency remain False.
12970                         .UsePriceHistory = False
12980                       Case False
                              ' ** Turn it back on then.
12990                         blnIncludeCurrency = True
13000                         .chkIncludeCurrency = True
13010                         .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
13020                       End Select
13030                       gblnMessage = False
13040                       gblnSwitchTo = False
13050                     Case False
                            ' ** Cancel this button.
13060                       blnRetVal = False
13070                       DoCmd.Hourglass False
13080                     End Select
13090                   End Select
13100                 End Select
13110               End If
13120             End If
13130           End If  ' ** blnFromStmts.

13140           If blnRetVal = True Then

13150             blnRetVal = BuildAssetListInfo_SP(frm, blnContinue, datAssetListDate, blnPrintAnnualStatement, blnAllStatements, blnNoDataAll, blnRollbackNeeded)  ' ** Module Function: modStatementParamFuncs1.

13160             If blnContinue = True And blnRetVal = True Then
13170               If blnHasForEx = True And .UsePriceHistory = True Then
13180                 Select Case blnIncludeCurrency
                      Case True
13190                   strReportName = "rptAssetList_ForEx"
13200                 Case False
13210                   strReportName = "rptAssetList"
13220                 End Select
13230               Else
13240                 strReportName = "rptAssetList"
13250               End If
13260               If strReportName = "rptAssetList_ForEx" Then
13270                 ForExRptSub_Load frm, blnRollbackNeeded  ' ** Module Procedure: modStatementParamFuncs1.
13280               End If
13290               If gblnDev_Debug = True Or (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
13300                 If IsLoaded(strReportName, acReport) = True Then  ' ** Module Function: modFileUtilities.
13310                   DoCmd.Close acReport, strReportName
13320                 End If
13330                 DoCmd.OpenReport strReportName, acViewPreview
13340               Else  ' ** Normal.
13350                 DoCmd.OpenReport strReportName, acViewPreview  ' ** This gets the caption changed!
                      '##GTR_Ref: rptAssetList
                      '##GTR_Ref: rptAssetList_ForEx
13360                 DoCmd.OpenReport strReportName, acViewNormal
13370               End If
                    'MAKE SURE IT'S CLOSED!
13380               If intOptGrpAcctNum <> .opgAccountNumber Then
13390                 .opgAccountNumber = intOptGrpAcctNum  '#Covered.
13400                 .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
13410               End If
13420             ElseIf blnNoDataAll = True Then
13430               If blnAllStatements = True Then
13440                 Select Case blnIncludeCurrency
                      Case True
13450                   strReportName = "rptAssetList_ForEx_NoData"
13460                 Case False
13470                   strReportName = "rptAssetList_NoData"
13480                 End Select
13490                 If gblnDev_Debug = True Or (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
13500                   If IsLoaded(strReportName, acReport) = True Then  ' ** Module Function: modFileUtilities.
13510                     DoCmd.Close acReport, strReportName
13520                   End If
13530                   DoCmd.OpenReport strReportName, acViewPreview
13540                 Else  ' ** Normal.
13550                   DoCmd.OpenReport strReportName, acViewPreview  ' ** This gets the caption changed!
13560                   DoCmd.OpenReport strReportName, acViewNormal
13570                 End If
                      'MAKE SURE IT'S CLOSED!
13580                 If intOptGrpAcctNum <> .opgAccountNumber Then
13590                   .opgAccountNumber = intOptGrpAcctNum  '#Covered.
13600                   .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
13610                 End If
13620               End If
13630             End If

13640           End If  ' ** blnRetVal.

13650         End If
13660       End If

13670       DoEvents
13680       If IsLoaded(strReportName, acReport) = True Then
13690         If gblnDev_Debug = False And ((CurrentUser <> "Superuser") Or (CurrentUser = "Superuser" And .chkAsDev = False)) Then
13700           DoCmd.Close acReport, strReportName
13710           DoEvents
13720         End If
13730       End If

13740       blnNoDataAll = False

13750     Else
13760       blnHasForExClick = False
13770       DoCmd.Hourglass False
13780     End If

13790   End With

EXITP:
13800   Set qdf = Nothing
13810   Set dbs = Nothing
13820   Exit Sub

ERRH:
13830   DoCmd.Hourglass False
13840   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
13850   Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
13860     blnContinue = False
13870   Case Else
13880     THAT_PROC = THIS_PROC
13890     That_Erl = Erl: That_Desc = ERR.description
13900     frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13910   End Select
13920   Resume EXITP

End Sub

Public Sub CmdAstListWord_Click_SP(blnContinue As Boolean, blnAllStatements As Boolean, blnPrintAnnualStatement As Boolean, blnNoDataAll As Boolean, blnRollbackNeeded As Boolean, datAssetListDate As Date, strFirstDateMsg As String, blnHasForExClick As Boolean, blnIncludeCurrency As Boolean, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctFors As Long, arr_varAcctFor As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdAstListWord_Click_SP(
' **   blnContinue As Boolean, blnAllStatements As Boolean, blnPrintAnnualStatement As Boolean, blnNoDataAll As Boolean,
' **   blnRollbackNeeded As Boolean, datAssetListDate As Date, strFirstDateMsg As String, blnHasForExClick As Boolean,
' **   blnIncludeCurrency As Boolean, blnHasForEx As Boolean, blnHasForExThis As Boolean, lngAcctFors As Long,
' **   arr_varAcctFor As Variant, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "CmdAstListWord_Click_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim intOptGrpAcctNum As Integer
        Dim blnPriceHistory As Boolean
        Dim strDocName As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** cmbAccounts combo box constants:
        Const CBX_A_ACTNO  As Integer = 0  ' ** accountno
        'Const CBX_A_DESC   As Integer = 1  ' ** Desc
        'Const CBX_A_PREDAT As Integer = 2  ' ** predate
        'Const CBX_A_SHORT  As Integer = 3  ' ** shortname
        'Const CBX_A_LEGAL  As Integer = 4  ' ** legalname
        'Const CBX_A_BALDAT As Integer = 5  ' ** BalanceDate (earliest [balance date])
        'Const CBX_A_HASREL As Integer = 6  ' ** HasRelated
        'Const CBX_A_CASNUM As Integer = 7  ' ** CaseNum
        'Const CBX_A_TRXDAT As Integer = 8  ' ** TransDate (earliest [transdate])

14010   With frm

14020     If blnHasForExClick = False Then

14030       blnContinue = True
14040       blnRetVal = True

14050       If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
14060         blnContinue = False
14070         MsgBox "You must select an account to continue," & vbCrLf & _
                "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "O01")
14080       Else

14090         If .chkAssetList = True Then
14100           If FirstDate_SP(frm) = False Then  ' ** Function: Below.
14110             blnRetVal = False
14120             MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "O02")
14130           End If
14140         End If

14150         If blnRetVal = True Then

14160           intOptGrpAcctNum = .opgAccountNumber
14170           gblnCombineAssets = .chkCombineCash.Value
14180           strRptCap = vbNullString: strRptPathFile = vbNullString

14190           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
14200             gstrAccountNo = .cmbAccounts
14210           Case .opgAccountNumber_optAll.OptionValue
14220             gstrAccountNo = "All"
14230           End Select
14240           blnIncludeCurrency = .chkIncludeCurrency
14250           datAssetListDate = .AssetListDate

14260           If .chkForeignExchange = True Then
14270             .currentDate = Null
14280             blnPriceHistory = PricingHistory(datAssetListDate)  ' ** Function: Below.
                  ' ** blnPriceHistory indicates whether current pricing or pricing history should be used.
                  ' ** It ONLY applies to foreign exchange, since regular reports don't require that info.
14290             .UsePriceHistory = blnPriceHistory
14300           End If

14310           If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
14320             If blnHasForEx = True And .UsePriceHistory = True Then
                    ' ** If all the rest of this code is using the foreign currency tables and queries,
                    ' ** but the user unchecked the box because this particular account has no foreign currency,
                    ' ** the report will end up looking in the wrong tables for the data!
14330               Select Case blnIncludeCurrency
                    Case True
14340                 gblnHasForExThis = False
14350                 blnHasForExThis = False
14360                 gblnSwitchTo = False
14370                 For lngX = 0& To (lngAcctFors - 1&)
14380                   If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
14390                     If arr_varAcctFor(F_ACNT, lngX) > 0 Then
14400                       gblnHasForExThis = True
14410                       blnHasForExThis = True
14420                     End If
14430                     gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)
14440                     Exit For
14450                   End If
14460                 Next
14470                 Select Case gblnHasForExThis
                      Case True
14480                   If gblnSwitchTo = True Then
                          ' ** Turn it off since they now do have foreign currencies.
14490                     Set dbs = CurrentDb
14500                     With dbs
                            ' ** Update tblCurrency_Account for curracct_supress = False, by specified [actno].
14510                       Set qdf = .QueryDefs("qryCurrency_17_02")
14520                       With qdf.Parameters
14530                         ![actno] = gstrAccountNo
14540                       End With
14550                       qdf.Execute
14560                       Set qdf = Nothing
14570                       .Close
14580                     End With
14590                     Set dbs = Nothing
14600                   End If
14610                 Case False
14620                   If gblnSwitchTo = True Then
                          ' ** If they've specified to suppress, then turn chkIncludeCurrency off.
14630                     blnIncludeCurrency = False
14640                     .chkIncludeCurrency = False
14650                     .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
14660                   End If
14670                 End Select
14680               Case False
14690                 gblnHasForExThis = False
14700                 blnHasForExThis = False
14710                 gblnSwitchTo = False
14720                 For lngX = 0& To (lngAcctFors - 1&)
14730                   If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
14740                     If arr_varAcctFor(F_ACNT, lngX) > 0 Then
14750                       gblnHasForExThis = True
14760                       blnHasForExThis = True
14770                     End If
14780                     gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)
14790                     Exit For
14800                   End If
14810                 Next
14820                 Select Case gblnHasForExThis
                      Case True
                        ' ** This account does have foreign currencies, and
                        ' ** the user shouldn't have been able to turn it off.
14830                   blnIncludeCurrency = True
14840                   .chkIncludeCurrency = True
14850                   .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
14860                 Case False
14870                   Select Case gblnSwitchTo
                        Case True
14880                     gblnMessage = True
14890                   Case False
14900                     strDocName = "frmStatementParameters_ForEx"
14910                     gblnSetFocus = True
14920                     gblnMessage = True  ' ** False return means cancel.
14930                     gblnSwitchTo = True  ' ** False return means show ForEx, don't supress.
14940                     DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & gstrAccountNo
14950                   End Select
14960                   Select Case gblnMessage
                        Case True
14970                     Select Case gblnSwitchTo
                          Case True
                            ' ** Let blnIncludeCurrency remain False.
14980                       .UsePriceHistory = False
14990                     Case False
                            ' ** Turn it back on then.
15000                       blnIncludeCurrency = True
15010                       .chkIncludeCurrency = True
15020                       .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
15030                     End Select
15040                     gblnMessage = False
15050                     gblnSwitchTo = False
15060                   Case False
                          ' ** Cancel this button.
15070                     blnRetVal = False
15080                     DoCmd.Hourglass False
15090                   End Select
15100                 End Select
15110               End Select
15120             End If
15130           End If

15140           If blnRetVal = True Then

                  ' ** Execute the common code.
15150             If BuildAssetListInfo_SP(frm, blnContinue, datAssetListDate, blnPrintAnnualStatement, blnAllStatements, blnNoDataAll, blnRollbackNeeded) = True Then  ' ** Module Function: modStatementParamFuncs1.
15160               If blnContinue = True Then

15170                 If IsNull(.UserReportPath) = True Then
15180                   strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
15190                 Else
15200                   strRptPath = .UserReportPath
15210                 End If
15220                 Select Case .opgAccountNumber
                      Case .opgAccountNumber_optSpecified.OptionValue
15230                   strRptCap = "rptAssetList_" & .cmbAccounts.Column(CBX_A_ACTNO) & "_"
15240                 Case .opgAccountNumber_optAll.OptionValue
15250                   strRptCap = "rptAssetList_All_"
15260                 End Select
15270                 strRptCap = StringReplace(strRptCap, "/", "_")  ' ** Module Function: modStringFuncs.
15280                 If .chkAssetList = True Then
15290                   strRptCap = strRptCap & "_" & Format(.AssetListDate, "yymmdd")
15300                 ElseIf .chkStatements = True Then
15310                   strRptCap = strRptCap & "_" & Format(.DateEnd, "yymmdd")
15320                 End If

15330                 Select Case blnIncludeCurrency
                      Case True
15340                   strRpt = "rptAssetList_ForEx"
15350                 Case False
15360                   strRpt = "rptAssetList"
15370                 End Select
15380                 If strRpt = "rptAssetList_ForEx" Then
15390                   ForExRptSub_Load frm, blnRollbackNeeded  ' ** Module Procedure: modStatementParamFuncs1.
15400                 End If

15410                 If gblnDev_Debug = True Or (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
15420                   DoCmd.OpenReport strRpt, acViewPreview
15430                 Else

15440                   strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

15450                   If strRptPathFile <> vbNullString Then
15460                     DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
15470                     .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
15480                   End If

15490                   If IsLoaded("rptAssetList", acReport) = True Then  ' ** Module Function: modFileUtilities.
15500                     DoCmd.Close acReport, "rptAssetList"
15510                   End If

15520                 End If  ' ** gblnDev_Debug.

15530                 If intOptGrpAcctNum <> .opgAccountNumber Then
15540                   .opgAccountNumber = intOptGrpAcctNum  '#Covered.
15550                   .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
15560                 End If

15570               End If  ' ** blnContinue.
15580             End If  ' ** BuildAssetListInfo_SP.

15590           End If  ' ** blnRetVal.

15600         End If
15610       End If

15620     Else
15630       blnHasForExClick = False
15640       DoCmd.Hourglass False
15650     End If

15660   End With

EXITP:
15670   Set qdf = Nothing
15680   Set dbs = Nothing
15690   Exit Sub

ERRH:
15700   DoCmd.Hourglass False
15710   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
15720     blnContinue = False
15730   Case Else
15740     THAT_PROC = THIS_PROC
15750     That_Erl = Erl: That_Desc = ERR.description
15760     frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15770   End Select
15780   Resume EXITP

End Sub

Public Sub CmdPrintStmtAll_Click_SP(blnContinue As Boolean, blnAllStatements As Boolean, blnPrintStatements As Boolean, blnSingleStatement As Boolean, blnRunPriorStatement As Boolean, blnFromStmts As Boolean, blnAcctsSchedRpt As Boolean, datFirstDate As Date, strFirstDateMsg As String, blnGTR_Emblem As Boolean, blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnWasGTR As Boolean, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdPrintStmtAll_Click_SP(
' **   blnContinue As Boolean, blnAllStatements As Boolean, blnPrintStatements As Boolean,
' **   blnSingleStatement As Boolean, blnRunPriorStatement As Boolean, blnFromStmts As Boolean,
' **   blnAcctsSchedRpt As Boolean, datFirstDate As Date, strFirstDateMsg As String, blnGTR_Emblem As Boolean,
' **   blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnWasGTR As Boolean,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "CmdPrintStmtAll_Click_SP"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strCurrentDate As String
        Dim blnRetVal As Boolean

15810   With frm

15820     blnRetVal = True

15830     Set dbs = CurrentDb

          ' ** See if there is a current masterasset date (the same for all records, so just get one).
15840     Set rst = dbs.OpenRecordset("SELECT Max(MasterAsset.currentDate) AS MaxDate FROM MasterAsset;", dbOpenSnapshot)
15850     If rst.BOF = True And rst.EOF = True Then
15860       blnRetVal = False
15870       Select Case gblnGoToReport
            Case True
15880         .TimerInterval = 0&
15890         .GoToReport_arw_sp_printall_img.Visible = False
15900         .cmdPrintStatement_Single.Visible = True
15910         blnGoingToReport2 = False
15920         blnGoingToReport = False
15930         gblnGoToReport = False
15940         blnGTR_Emblem = False
15950         .GTREmblem_Off  ' ** Form Procedure: frmStatementParameters.
15960         Beep
15970         DoCmd.Hourglass False
15980         MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                "Assets must be priced to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
15990       Case False
16000         MsgBox "Assets must be priced in order to run statements.", vbInformation + vbOKOnly, _
                (Left(("Missing An Asset Current Date" & Space(55)), 55) & "Q01")
16010       End Select
16020       rst.Close
16030       dbs.Close
16040     Else

16050       If FirstDate_SP(frm) = False Then  ' ** Function: Below.
16060         blnRetVal = False
16070         Select Case gblnGoToReport
              Case True
16080           .TimerInterval = 0&
16090           .GoToReport_arw_sp_printall_img.Visible = False
16100           .cmdPrintStatement_Single.Visible = True
16110           blnGoingToReport2 = False
16120           blnGoingToReport = False
16130           gblnGoToReport = False
16140           blnGTR_Emblem = False
16150           .GTREmblem_Off  ' ** Form Procedure: frmStatementParameters.
16160           Beep
16170           DoCmd.Hourglass False
16180           MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                  "There is insufficient data to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
16190         Case False
16200           MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "Q02")
16210         End Select
16220       End If

16230       If blnRetVal = True Then

16240         rst.MoveFirst
16250         If IsNull(rst![MaxDate]) Then
16260           MsgBox "Assets must be priced in order to run statements.", vbInformation + vbOKOnly, _
                  (Left(("Missing An Asset Current Date" & Space(55)), 55) & "Q03")
16270           rst.Close
16280           dbs.Close
16290         Else
16300           strCurrentDate = Trim(CStr(rst![MaxDate]))
16310           rst.Close
16320           dbs.Close
16330           Set rst = Nothing
16340           Set dbs = Nothing

                ' ** There is, so print all of the pieces.
16350           blnAllStatements = True
16360           Statements_Print frm, blnPrintStatements, blnAllStatements, blnSingleStatement, _
                  blnRunPriorStatement, blnAcctsSchedRpt, datFirstDate, blnContinue, blnFromStmts, _
                  blnGoingToReport, blnGoingToReport2, blnGTR_Emblem, blnWasGTR  ' ** Module Function: modStatementParamFuncs1.

16370           If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
16380             .opgAccountNumber = .opgAccountNumber_optAll.OptionValue  '#Covered.
16390             .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
16400           End If

16410           blnAllStatements = False

16420         End If
16430       End If
16440     End If

16450   End With

EXITP:
16460   Set rst = Nothing
16470   Set dbs = Nothing
16480   Exit Sub

ERRH:
16490   DoCmd.Hourglass False
16500   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
16510   Case Else
16520     THAT_PROC = THIS_PROC
16530     That_Erl = Erl: That_Desc = ERR.description
16540     frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16550   End Select
16560   Resume EXITP

End Sub

Public Sub CmdPrintStmtSingle_Click_SP(blnContinue As Boolean, blnPrintStatements As Boolean, blnAllStatements As Boolean, blnSingleStatement As Boolean, blnRunPriorStatement As Boolean, blnAcctsSchedRpt As Boolean, datFirstDate As Date, blnFromStmts As Boolean, blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, blnWasGTR As Boolean, strFirstDateMsg As String, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdPrintStmtSingle_Click_SP(
' **   blnContinue As Boolean, blnPrintStatements As Boolean, blnAllStatements As Boolean, blnSingleStatement As Boolean,
' **   blnRunPriorStatement As Boolean, blnAcctsSchedRpt As Boolean, datFirstDate As Date, blnFromStmts As Boolean,
' **   blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, blnWasGTR As Boolean,
' **   strFirstDateMsg As String, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "CmdPrintStmtSingle_Click_SP"

        Dim blnRetVal As Boolean

        ' ** cmbMonth combo box constants:
        Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
        'Const CBX_MON_NAME  As Integer = 1  ' ** month_name
        'Const CBX_MON_SHORT As Integer = 2  ' ** month_short

16610   With frm

16620     blnRetVal = True

16630     If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
16640       MsgBox "You must select an account to continue," & vbCrLf & _
              "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "R01")
16650     Else

16660       If FirstDate_SP(frm) = False Then  ' ** Function: Below.
16670         blnRetVal = False
16680         MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "R02")
16690       End If

16700       If blnRetVal = True Then

16710         If glngMonthID = 0& Then
16720           If IsNull(.cmbMonth) = True Then
16730             .cmbMonth = "December"
16740           End If
16750           glngMonthID = .cmbMonth.Column(CBX_MON_ID)
16760         End If

16770         blnContinue = True  ' ** Unless user cancels.
16780         blnSingleStatement = True
16790         Statements_Print frm, blnPrintStatements, blnAllStatements, blnSingleStatement, _
                blnRunPriorStatement, blnAcctsSchedRpt, datFirstDate, blnContinue, blnFromStmts, _
                blnGoingToReport, blnGoingToReport2, blnGTR_Emblem, blnWasGTR  ' ** Module Function: modStatementParamFuncs1.

16800         If blnAcctsSchedRpt = False Then

16810           blnRunPriorStatement = False

16820           .cmbAccounts.Enabled = True
16830           .cmbAccounts.Locked = False
16840           .cmbAccounts.ForeColor = CLR_BLK
16850           .cmbAccounts.BackColor = CLR_WHT
16860           .cmbAccounts.BorderColor = CLR_LTBLU2
16870           .cmbAccounts.BackStyle = acBackStyleNormal
16880           .cmbAccounts_lbl.ForeColor = CLR_WHT
16890           .cmbAccounts_lbl.BackStyle = acBackStyleNormal
16900           .opgAccountSource.Enabled = True
16910           .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
16920           .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
16930           .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
16940           .opgAccountSource_optName_lbl2_dim_hi.Visible = False
16950           .chkRememberMe.Enabled = True
16960           .chkRememberMe_lbl.Visible = True
16970           .chkRememberMe_lbl2_dim.Visible = False
16980           .chkRememberMe_lbl2_dim_hi.Visible = False
16990           .cmdPrintStatement_Single.Caption = "Reprint Single Statement"
17000           .cmdPrintStatement_Single.ControlTipText = "Reprint Single" & vbCrLf & "Statement - Ctrl+S"
17010           .cmdPrintStatement_Single.StatusBarText = "Reprint Single Statement - Ctrl+S"
17020           .cmdPrintStatement_Summary.Enabled = True

17030         End If  ' ** blnAcctsSchedRpt.

17040       End If

17050     End If

17060   End With

EXITP:
17070   Exit Sub

ERRH:
17080   DoCmd.Hourglass False
17090   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
17100   Case Else
17110     THAT_PROC = THIS_PROC
17120     That_Erl = Erl: That_Desc = ERR.description
17130     frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17140   End Select
17150   Resume EXITP

End Sub

Public Sub CmdPrintStmtSum_Click_SP(blnContinue As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String, THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form)
' **
' ** CmdPrintStmtSum_Click_SP(
' **   blnContinue As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String,
' **   THAT_PROC As String, That_Erl As Long, That_Desc As String, frm As Access.Form
' ** )

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "CmdPrintStmtSum_Click_SP"

        Dim intOpgAccountNumber As Integer
        Dim blnContinue2 As Boolean, blnNoAccount As Boolean
        Dim blnRetVal As Boolean

17210   With frm

17220     DoCmd.Hourglass True
17230     DoEvents

17240     blnContinue2 = True
17250     blnRetVal = True
17260     blnNoAccount = False

17270     If IsNull(.cmbMonth) = True Then
17280       blnContinue2 = False
17290       DoCmd.Hourglass False
17300       MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, (Left(("Entry Required" & Space(55)), 55) & "S01")
17310       .cmbMonth.SetFocus
17320     Else
17330       If .cmbMonth = vbNullString Then
17340         blnContinue2 = False
17350         DoCmd.Hourglass False
17360         MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, (Left(("Entry Required" & Space(55)), 55) & "S02")
17370         .cmbMonth.SetFocus
17380       Else
17390         intOpgAccountNumber = 0
17400         If .cmbAccounts.Enabled = True Then
17410           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
17420             If IsNull(.cmbAccounts) = True Then
17430               blnNoAccount = True
17440             Else
17450               If .cmbAccounts = vbNullString Then
17460                 blnNoAccount = True
17470               Else
17480                 intOpgAccountNumber = .opgAccountNumber_optSpecified.OptionValue
17490               End If
17500             End If
17510           Case .opgAccountNumber_optAll.OptionValue
17520             intOpgAccountNumber = .opgAccountNumber_optAll.OptionValue
17530           End Select
17540         Else
17550           intOpgAccountNumber = .opgAccountNumber_optAll.OptionValue
17560         End If
17570         If blnNoAccount = True Then
17580           blnContinue2 = False
17590           DoCmd.Hourglass False
17600           MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "T01")
17610         End If
17620       End If
17630     End If

17640     If FirstDate_SP(frm) = False Then  ' ** Function: Below.
17650       blnRetVal = False
17660       DoCmd.Hourglass False
17670       MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "R02")
17680     End If

17690     If blnRetVal = True Then
17700       If blnContinue2 = True Then
17710         blnFromStmts = True
17720         .cmdSummaryPrint  ' ** Form Procedure: frmStatementParameters.
17730         blnFromStmts = False

              ' ** qryMaxBalDates is written in the SetDateSpecificSQL() function, which is called by:
              ' **   BuildTransactionInfo_SP()
              ' **     SetDateSpecificSQL(Me.cmbAccounts, "StatementTransactions", THIS_NAME)
              ' **   BuildAssetListInfo_SP()
              ' **     SetDateSpecificSQL(Me.cmbAccounts, "Statements", THIS_NAME)
              ' ** It is set depending on the parameters supplied:
              ' **   SetDateSpecificSQL(strAccountNo As String, strOption As String, strActiveFormName As String,
              ' **     Optional varStartDate As Variant, Optional varEndDate As Variant, Optional varIsArchive As Variant)
              ' ** Since neither of the above send a varEndDate, it uses the 2nd of the options:
              ' **   1. strEndDate = Forms(strActiveFormName)!TransDateEnd
              ' **   2. strEndDate = Forms(strActiveFormName)!DateEnd
              ' **   3. strEndDate = Format(CDate(varEndDate), "mm/dd/yyyy")
              ' ** Of the 2 field names used in qryMaxBalDates, Statements uses the 1st.
              ' **   1. strMaxDateFld = "MaxOfbalance date"
              ' **   2. strMaxDateFld = "balance date"
              ' ** SQL of qryMaxBalDates as generated in SetDateSpecificSQL():
              ' **   SELECT Balance.accountno As accountno, Max(Balance.[balance date]) AS [" & strMaxDateFld & "] " & _
              ' **     "FROM Balance " & _
              ' **     "WHERE (((Balance.[balance date]) < #" & strEndDate & "#)) " & _
              ' **     "GROUP BY Balance.accountno;
              ' ** Resulting in:
              ' **   SELECT Balance.accountno As accountno, Max(Balance.[balance date]) AS [MaxOfbalance date]
              ' **     FROM Balance
              ' **     WHERE (((Balance.[balance date]) < FormRef('BalanceDate')))
              ' **     GROUP BY Balance.accountno;

17740       End If

17750     End If

17760     DoCmd.Hourglass False

17770   End With

EXITP:
17780   Exit Sub

ERRH:
17790   DoCmd.Hourglass False
17800   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
17810   Case Else
17820     THAT_PROC = THIS_PROC
17830     That_Erl = Erl: That_Desc = ERR.description
17840     frm.Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17850   End Select
17860   Resume EXITP

End Sub

Public Sub Btn_Enable_SP(intMode As Integer, blnIsOpen As Boolean, blnRunPriorStatement As Boolean, blnAcctNotSched As Boolean, datAssetListDate_Pref As Date, lngStmts As Long, arr_varStmt As Variant, frm As Access.Form)
' ** Called by:
' **   modStatementParamFuncs2:
' **     ChkTrans_After_SP()
' **     ChkAstList_After_SP()
' **     ChkStmt_After_SP()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "Btn_Enable_SP"

        'CLR_DISABLED_FG
        'CLR_DISABLED_BG

17910   With frm

17920     Select Case intMode
          Case 1  ' ** Transactions.

17930       Select Case .chkTransactions
            Case True
              ' ** Checked.
17940         .chkAssetList.Value = False
17950         .chkStatements = False
17960         .PrintAnnual_chk = False
17970         .cmdAnnualStatement.Enabled = False
17980         .cmdAnnualStatement_raised_img_dis.Visible = True
17990         .cmdAnnualStatement_raised_img.Visible = False
18000         .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
18010         .cmdAnnualStatement_raised_focus_img.Visible = False
18020         .cmdAnnualStatement_raised_focus_dots_img.Visible = False
18030         .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
18040         .cmbAccounts_lbl2.Visible = False
18050         .cmbAccounts_lbl3.Visible = False
18060         .TransDateStart.Enabled = True
18070         .TransDateStart.Locked = False
18080         .TransDateStart.BorderColor = CLR_LTBLU2
18090         .TransDateStart.ForeColor = CLR_BLK
18100         .TransDateStart.BackColor = CLR_WHT
18110         .TransDateStart_lbl.ForeColor = CLR_WHT
18120         .TransDateStart_lbl.BackStyle = acBackStyleNormal
18130         .TransDateStart_lbl_box.Visible = False
18140         .TransDateStart_lbl_box.BorderColor = MY_CLR_LTBGE
18150         .TransDateEnd.Enabled = True
18160         .TransDateEnd.Locked = False
18170         .TransDateEnd.BorderColor = CLR_LTBLU2
18180         .TransDateEnd.ForeColor = CLR_BLK
18190         .TransDateEnd.BackColor = CLR_WHT
18200         .TransDateEnd_lbl.ForeColor = CLR_WHT
18210         .TransDateEnd_lbl.BackStyle = acBackStyleNormal
18220         .TransDateEnd_lbl_box.Visible = False
18230         .TransDateEnd_lbl_box.BorderColor = MY_CLR_LTBGE
18240         .chkRememberDates_Trans.Enabled = True
18250         .chkRememberDates_Trans_lbl.Visible = True
18260         .chkRememberDates_Trans_lbl2_dim.Visible = False
18270         .chkRememberDates_Trans_lbl2_dim_hi.Visible = False
18280         .AssetListDate.Enabled = False
18290         .AssetListDate.Locked = True
18300         .AssetListDate.BorderColor = WIN_CLR_DISR
18310         .AssetListDate.ForeColor = CLR_LTGRY
18320         .AssetListDate.BackColor = MY_CLR_MDBGE
18330         .AssetListDate_lbl.ForeColor = WIN_CLR_DISF
18340         .AssetListDate_lbl.BackStyle = acBackStyleTransparent
18350         .AssetListDate_lbl_box.BorderColor = MY_CLR_LTBGE
18360         .AssetListDate_lbl_box.Visible = True
18370         .chkRememberDates_Asset.Enabled = False
18380         .chkRememberDates_Asset_lbl.Visible = False
18390         .chkRememberDates_Asset_lbl2_dim.Visible = True
18400         .chkRememberDates_Asset_lbl2_dim_hi.Visible = True
18410         .cmbMonth.Enabled = False
18420         .cmbMonth.Locked = True
18430         .cmbMonth.BorderColor = WIN_CLR_DISR
18440         .cmbMonth.ForeColor = CLR_LTGRY
18450         .cmbMonth.BackColor = MY_CLR_MDBGE
18460         .cmbMonth_lbl.ForeColor = WIN_CLR_DISF
18470         .cmbMonth_lbl.BackStyle = acBackStyleTransparent
18480         .cmbMonth_lbl_box.BorderColor = MY_CLR_LTBGE
18490         .cmbMonth_lbl_box.Visible = True
18500         .StatementsYear.Enabled = False
18510         .StatementsYear.Locked = True
18520         .StatementsYear.BorderColor = WIN_CLR_DISR
18530         .StatementsYear.ForeColor = CLR_LTGRY
18540         .StatementsYear.BackColor = MY_CLR_MDBGE
18550         .chkStatements_lbl2.ForeColor = CLR_GRY
18560         .chkStatements_lbl3.ForeColor = CLR_GRY
18570         DoCmd.Hourglass True  ' ** Assure it's still going.
18580         DoEvents
18590         .DateEnd = Null
18600         .opgAccountNumber.Enabled = True
18610         .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue  '#Covered.
18620         .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
18630         .opgOrderBy.Enabled = True
18640         .chkRelatedAccounts.Enabled = False
18650         .chkRelatedAccounts = False
18660         .chkRelatedAccounts_AfterUpdate  ' ** Procedure: Below.
18670         SetRelatedOption frm  ' ** Procedure: Above.
18680         .chkCombineCash.Enabled = True
18690         .chkCombineCash_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
18700         .cmdTransactions_lbl_dim.Visible = False
18710         .cmdTransactions_lbl.ForeColor = CLR_WHT
18720         .cmdTransactions_lbl.BackStyle = acBackStyleNormal
18730         .cmdTransactions_lbl_box_dim.Visible = False
18740         .cmdTransactions_lbl_hline01_dim.Visible = False
18750         .cmdTransactions_lbl_hline02_dim.Visible = False
18760         .cmdTransactions_lbl_hline03_dim.Visible = False
18770         .cmdTransactions_lbl_vline01_dim.Visible = False
18780         .cmdTransactions_lbl_vline02_dim.Visible = False
18790         .cmdTransactions_lbl_vline03_dim.Visible = False
18800         .cmdTransactions_lbl_vline04_dim.Visible = False
18810         DoEvents
18820         .cmdTransactionsPreview.Enabled = True
18830         .cmdTransactionsPrint.Enabled = True
18840         .cmdTransactionsWord.Enabled = True
      #If NoExcel Then
18850         .cmdTransactionsExcel.Enabled = False
      #Else
18860         .cmdTransactionsExcel.Enabled = True
      #End If
18870         .cmdAssetList_lbl_dim.Visible = True
18880         .cmdAssetList_lbl.ForeColor = WIN_CLR_DISF
18890         .cmdAssetList_lbl.BackStyle = acBackStyleTransparent
18900         .cmdAssetList_lbl_box_dim.Visible = True
18910         .cmdAssetList_lbl_hline01_dim.Visible = True
18920         .cmdAssetList_lbl_hline02_dim.Visible = True
18930         .cmdAssetList_lbl_hline03_dim.Visible = True
18940         .cmdAssetList_lbl_vline01_dim.Visible = True
18950         .cmdAssetList_lbl_vline02_dim.Visible = True
18960         .cmdAssetList_lbl_vline03_dim.Visible = True
18970         .cmdAssetList_lbl_vline04_dim.Visible = True
18980         DoCmd.Hourglass True  ' ** Assure it's still going.
18990         DoEvents
19000         .cmdAssetListPreview.Enabled = False
19010         .cmdAssetListPrint.Enabled = False
19020         .cmdAssetListWord.Enabled = False
      #If NoExcel Then
19030         .cmdAssetListExcel.Enabled = False
      #Else
19040         .cmdAssetListExcel.Enabled = False
      #End If
19050 On Error Resume Next
19060         .TransDateStart.SetFocus
19070 On Error GoTo ERRH
19080         SetStatementOptions frm, blnRunPriorStatement  ' ** Procedure: Below.
19090         Calendar_Set_SP frm, True, True, False  ' ** Procedure: Below.
19100       Case False
              ' ** Unchecked.
19110         SetStatementOptions frm, blnRunPriorStatement  ' ** Procedure: Below.
              ' ** If the user manually unchecks the box, move the check elsewhere.
19120         If blnIsOpen = False Then
19130           .chkAssetList = True
19140           .chkAssetList_AfterUpdate  ' ** Form Procedure: frmStatementParameters
19150         End If
19160       End Select

19170     Case 2  ' ** Asset List.

19180       Select Case .chkAssetList
            Case True
              ' ** Checked.
19190         .chkTransactions = False
19200         .chkStatements = False
19210         .PrintAnnual_chk = False
19220         .cmdAnnualStatement.Enabled = False
19230         .cmdAnnualStatement_raised_img_dis.Visible = True
19240         .cmdAnnualStatement_raised_img.Visible = False
19250         .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
19260         .cmdAnnualStatement_raised_focus_img.Visible = False
19270         .cmdAnnualStatement_raised_focus_dots_img.Visible = False
19280         .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
19290         .cmbAccounts_lbl2.Visible = False
19300         .cmbAccounts_lbl3.Visible = False
19310         .chkRememberDates_Asset.Enabled = True
19320         .chkRememberDates_Asset_lbl.Visible = True
19330         .chkRememberDates_Asset_lbl2_dim.Visible = False
19340         .chkRememberDates_Asset_lbl2_dim_hi.Visible = False
19350         .TransDateStart.Enabled = False
19360         .TransDateStart.Locked = True
19370         .TransDateStart.BorderColor = WIN_CLR_DISR
19380         .TransDateStart.ForeColor = CLR_LTGRY
19390         .TransDateStart.BackColor = MY_CLR_MDBGE
19400         .TransDateStart_lbl.ForeColor = WIN_CLR_DISF
19410         .TransDateStart_lbl.BackStyle = acBackStyleTransparent
19420         .TransDateStart_lbl_box.BorderColor = MY_CLR_LTBGE
19430         .TransDateStart_lbl_box.Visible = True
19440         .TransDateEnd.Enabled = False
19450         .TransDateEnd.Locked = True
19460         .TransDateEnd.BorderColor = WIN_CLR_DISR
19470         .TransDateEnd.ForeColor = CLR_LTGRY
19480         .TransDateEnd.BackColor = MY_CLR_MDBGE
19490         .TransDateEnd_lbl.ForeColor = WIN_CLR_DISF
19500         .TransDateEnd_lbl.BackStyle = acBackStyleTransparent
19510         .TransDateEnd_lbl_box.BorderColor = MY_CLR_LTBGE
19520         .TransDateEnd_lbl_box.Visible = True
19530         .chkRememberDates_Trans.Enabled = False
19540         .chkRememberDates_Trans_lbl.Visible = False
19550         .chkRememberDates_Trans_lbl2_dim.Visible = True
19560         .chkRememberDates_Trans_lbl2_dim_hi.Visible = True
19570         .AssetListDate.Enabled = True
19580         .AssetListDate.Locked = False
19590         .AssetListDate.BorderColor = CLR_LTBLU2
19600         .AssetListDate.ForeColor = CLR_BLK
19610         .AssetListDate.BackColor = CLR_WHT
19620         .AssetListDate_lbl.ForeColor = CLR_WHT
19630         .AssetListDate_lbl.BackStyle = acBackStyleNormal
19640         .AssetListDate_lbl_box.Visible = False
19650         .AssetListDate_lbl_box.BorderColor = MY_CLR_LTBGE
19660         .cmbMonth.Enabled = False
19670         .cmbMonth.Locked = True
19680         .cmbMonth.BorderColor = WIN_CLR_DISR
19690         .cmbMonth.ForeColor = CLR_LTGRY
19700         .cmbMonth.BackColor = MY_CLR_MDBGE
19710         .cmbMonth_lbl.ForeColor = WIN_CLR_DISF
19720         .cmbMonth_lbl.BackStyle = acBackStyleTransparent
19730         .cmbMonth_lbl_box.BorderColor = MY_CLR_LTBGE
19740         .cmbMonth_lbl_box.Visible = True
19750         .StatementsYear.Enabled = False
19760         .StatementsYear.Locked = True
19770         .StatementsYear.BorderColor = WIN_CLR_DISR
19780         .StatementsYear.ForeColor = CLR_LTGRY
19790         .StatementsYear.BackColor = MY_CLR_MDBGE
19800         .chkStatements_lbl2.ForeColor = CLR_GRY
19810         .chkStatements_lbl3.ForeColor = CLR_GRY
19820         DoCmd.Hourglass True  ' ** Assure it's still going.
19830         DoEvents
19840         .DateEnd = Now()
19850         .opgAccountNumber.Enabled = True
19860         .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue  '#Covered.
19870         .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
19880         .opgOrderBy.Enabled = False
19890         .chkRelatedAccounts.Enabled = True
19900         .chkRelatedAccounts = False
19910         .chkRelatedAccounts_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
19920         SetRelatedOption frm  ' ** Procedure: Above.
19930         .chkCombineCash.Enabled = True
19940         .chkCombineCash_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
19950         .cmdTransactions_lbl_dim.Visible = True
19960         .cmdTransactions_lbl.ForeColor = WIN_CLR_DISF
19970         .cmdTransactions_lbl.BackStyle = acBackStyleTransparent
19980         .cmdTransactions_lbl_box_dim.Visible = True
19990         .cmdTransactions_lbl_hline01_dim.Visible = True
20000         .cmdTransactions_lbl_hline02_dim.Visible = True
20010         .cmdTransactions_lbl_hline03_dim.Visible = True
20020         .cmdTransactions_lbl_vline01_dim.Visible = True
20030         .cmdTransactions_lbl_vline02_dim.Visible = True
20040         .cmdTransactions_lbl_vline03_dim.Visible = True
20050         .cmdTransactions_lbl_vline04_dim.Visible = True
20060         DoEvents
20070         .cmdTransactionsPreview.Enabled = False
20080         .cmdTransactionsPrint.Enabled = False
20090         .cmdTransactionsWord.Enabled = False
      #If NoExcel Then
20100         .cmdTransactionsExcel.Enabled = False
      #Else
20110         .cmdTransactionsExcel.Enabled = False
      #End If
20120         .cmdAssetList_lbl_dim.Visible = False
20130         .cmdAssetList_lbl.ForeColor = CLR_WHT
20140         .cmdAssetList_lbl.BackStyle = acBackStyleNormal
20150         .cmdAssetList_lbl_box_dim.Visible = False
20160         .cmdAssetList_lbl_hline01_dim.Visible = False
20170         .cmdAssetList_lbl_hline02_dim.Visible = False
20180         .cmdAssetList_lbl_hline03_dim.Visible = False
20190         .cmdAssetList_lbl_vline01_dim.Visible = False
20200         .cmdAssetList_lbl_vline02_dim.Visible = False
20210         .cmdAssetList_lbl_vline03_dim.Visible = False
20220         .cmdAssetList_lbl_vline04_dim.Visible = False
20230         DoEvents
20240         .cmdAssetListPreview.Enabled = True
20250         .cmdAssetListPrint.Enabled = True
20260         .cmdAssetListWord.Enabled = True
      #If NoExcel Then
20270         .cmdAssetListExcel.Enabled = False
      #Else
20280         .cmdAssetListExcel.Enabled = True
      #End If
20290         DoCmd.Hourglass True  ' ** Assure it's still going.
20300         If IsNull(.AssetListDate) = True Then
                ' ** Populate it with today's date.
20310           .AssetListDate = Date
20320           .DateEnd = .AssetListDate
20330         Else
20340           If .AssetListDate = vbNullString Then
                  ' ** Populate it with today's date.
20350             .AssetListDate = Date
20360             .DateEnd = .AssetListDate
20370           Else
20380             If .AssetListDate <> Date Then
20390               If datAssetListDate_Pref <> 0 Then
                      ' ** Leave it be.
20400               Else
                      ' ** Switch it back if changed by Statements.
20410                 .AssetListDate = Date
20420                 DoEvents
20430               End If
20440             End If
20450             .DateEnd = .AssetListDate
20460           End If
20470         End If
20480         DoEvents
20490 On Error Resume Next
20500         .AssetListDate.SetFocus
20510 On Error GoTo ERRH
20520         SetStatementOptions frm, blnRunPriorStatement  ' ** Procedure: Below.
20530         Calendar_Set_SP frm, False, False, True  ' ** Procedure: Below.
20540       Case False
20550         SetStatementOptions frm, blnRunPriorStatement  ' ** Procedure: Below.
              ' ** If the user manually unchecks the box, move the check elsewhere.
20560         If blnIsOpen = False Then
20570           .chkStatements = True
20580           .chkStatements_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
20590         End If
20600       End Select

20610     Case 3  ' ** Statements.

20620       Select Case .chkStatements
            Case True
              ' ** Checked.
20630         .chkAssetList = False
20640         .chkTransactions = False
20650         .PrintAnnual_chk = False
20660         .cmdAnnualStatement.Enabled = True
20670         .cmdAnnualStatement_raised_img.Visible = True
20680         .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
20690         .cmdAnnualStatement_raised_focus_img.Visible = False
20700         .cmdAnnualStatement_raised_focus_dots_img.Visible = False
20710         .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
20720         .cmdAnnualStatement_raised_img_dis.Visible = False
20730         .cmbAccounts_lbl2.Visible = False
20740         .cmbAccounts_lbl3.Visible = False
20750         .TransDateStart.Enabled = False
20760         .TransDateStart.Locked = True
20770         .TransDateStart.BorderColor = WIN_CLR_DISR
20780         .TransDateStart.ForeColor = CLR_LTGRY
20790         .TransDateStart.BackColor = MY_CLR_MDBGE
20800         .TransDateStart_lbl.ForeColor = WIN_CLR_DISF
20810         .TransDateStart_lbl.BackStyle = acBackStyleTransparent
20820         .TransDateStart_lbl_box.BorderColor = MY_CLR_LTBGE
20830         .TransDateStart_lbl_box.Visible = True
20840         .TransDateEnd.Enabled = False
20850         .TransDateEnd.Locked = True
20860         .TransDateEnd.BorderColor = WIN_CLR_DISR
20870         .TransDateEnd.ForeColor = CLR_LTGRY
20880         .TransDateEnd.BackColor = MY_CLR_MDBGE
20890         .TransDateEnd_lbl.ForeColor = WIN_CLR_DISF
20900         .TransDateEnd_lbl.BackStyle = acBackStyleTransparent
20910         .TransDateEnd_lbl_box.BorderColor = MY_CLR_LTBGE
20920         .TransDateEnd_lbl_box.Visible = True
20930         .chkRememberDates_Trans.Enabled = False
20940         .chkRememberDates_Trans_lbl.Visible = False
20950         .chkRememberDates_Trans_lbl2_dim.Visible = True
20960         .chkRememberDates_Trans_lbl2_dim_hi.Visible = True
20970         .AssetListDate.Enabled = False
20980         .AssetListDate.Locked = True
20990         .AssetListDate.BorderColor = WIN_CLR_DISR
21000         .AssetListDate.ForeColor = CLR_LTGRY
21010         .AssetListDate.BackColor = MY_CLR_MDBGE
21020         .AssetListDate_lbl.ForeColor = WIN_CLR_DISF
21030         .AssetListDate_lbl.BackStyle = acBackStyleTransparent
21040         .AssetListDate_lbl_box.BorderColor = MY_CLR_LTBGE
21050         .AssetListDate_lbl_box.Visible = True
21060         .chkRememberDates_Asset.Enabled = False
21070         .chkRememberDates_Asset_lbl.Visible = False
21080         .chkRememberDates_Asset_lbl2_dim.Visible = True
21090         .chkRememberDates_Asset_lbl2_dim_hi.Visible = True
21100         .cmbMonth.Enabled = True
21110         .cmbMonth.Locked = False
21120         .cmbMonth.BorderColor = CLR_LTBLU2
21130         .cmbMonth.ForeColor = CLR_BLK
21140         .cmbMonth.BackColor = CLR_WHT
21150         .cmbMonth_lbl.ForeColor = CLR_WHT
21160         .cmbMonth_lbl.BackStyle = acBackStyleNormal
21170         .cmbMonth_lbl_box.Visible = False
21180         .cmbMonth_lbl_box.BorderColor = MY_CLR_LTBGE
21190         .StatementsYear.Enabled = True
21200         .StatementsYear.Locked = False
21210         .StatementsYear.BorderColor = CLR_LTBLU2
21220         .StatementsYear.ForeColor = CLR_BLK
21230         .StatementsYear.BackColor = CLR_WHT
21240         .chkStatements_lbl2.ForeColor = CLR_BLK
21250         .chkStatements_lbl3.ForeColor = CLR_BLK
21260         .DateEnd = Null
21270         DoCmd.Hourglass True  ' ** Assure it's still going.
21280         DoEvents
21290         .opgAccountNumber.Enabled = True
21300         .opgAccountNumber = .opgAccountNumber_optAll.OptionValue  '#Covered.
21310         .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
21320         .opgOrderBy.Enabled = True
21330         .chkRelatedAccounts.Enabled = False
21340         .chkRelatedAccounts = False
21350         .chkRelatedAccounts_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
21360         SetRelatedOption frm  ' ** Procedure: Above.
21370         .chkCombineCash.Enabled = True
21380         .chkCombineCash_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
21390         .cmdTransactions_lbl_dim.Visible = False
21400         .cmdTransactions_lbl.ForeColor = CLR_WHT
21410         .cmdTransactions_lbl.BackStyle = acBackStyleNormal
21420         .cmdTransactions_lbl_box_dim.Visible = False
21430         .cmdTransactions_lbl_hline01_dim.Visible = False
21440         .cmdTransactions_lbl_hline02_dim.Visible = False
21450         .cmdTransactions_lbl_hline03_dim.Visible = False
21460         .cmdTransactions_lbl_vline01_dim.Visible = False
21470         .cmdTransactions_lbl_vline02_dim.Visible = False
21480         .cmdTransactions_lbl_vline03_dim.Visible = False
21490         .cmdTransactions_lbl_vline04_dim.Visible = False
21500         DoCmd.Hourglass True  ' ** Assure it's still going.
21510         DoEvents
21520         .cmdTransactionsPreview.Enabled = True
21530         .cmdTransactionsPrint.Enabled = True
21540         .cmdTransactionsWord.Enabled = True
      #If NoExcel Then
21550         .cmdTransactionsExcel.Enabled = False
      #Else
21560         .cmdTransactionsExcel.Enabled = True
      #End If
21570         .cmdAssetList_lbl_dim.Visible = False
21580         .cmdAssetList_lbl.ForeColor = CLR_WHT
21590         .cmdAssetList_lbl.BackStyle = acBackStyleNormal
21600         .cmdAssetList_lbl_box_dim.Visible = False
21610         .cmdAssetList_lbl_hline01_dim.Visible = False
21620         .cmdAssetList_lbl_hline02_dim.Visible = False
21630         .cmdAssetList_lbl_hline03_dim.Visible = False
21640         .cmdAssetList_lbl_vline01_dim.Visible = False
21650         .cmdAssetList_lbl_vline02_dim.Visible = False
21660         .cmdAssetList_lbl_vline03_dim.Visible = False
21670         .cmdAssetList_lbl_vline04_dim.Visible = False
21680         DoEvents
21690         .cmdAssetListPreview.Enabled = True
21700         .cmdAssetListPrint.Enabled = True
21710         .cmdAssetListWord.Enabled = True
      #If NoExcel Then
21720         .cmdAssetListExcel.Enabled = False
      #Else
21730         .cmdAssetListExcel.Enabled = True
      #End If
21740 On Error Resume Next
21750         .cmbMonth.SetFocus
21760 On Error GoTo ERRH
21770         SetStatementOptions frm, blnRunPriorStatement  ' ** Procedure: Below.
21780         Calendar_Set_SP frm, False, False, False  ' ** Procedure: Below.
21790         If IsEmpty(arr_varStmt) = True Then
21800           arr_varStmt = AcctSched_Load  ' ** Module Function: modStatementParamFuncs1.
21810           lngStmts = UBound(arr_varStmt, 1)
21820         End If
21830         blnAcctNotSched = False
21840       Case False
              ' ** Unchecked.
21850         SetStatementOptions frm, blnRunPriorStatement  ' ** Procedure: Below.
              ' ** If the user manually unchecks the box, move the check elsewhere.
21860         If blnIsOpen = False Then
21870           .chkTransactions = True
21880           .chkTransactions_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
21890         End If
21900       End Select

21910     End Select

21920   End With

EXITP:
21930   Exit Sub

ERRH:
21940   Select Case ERR.Number
        Case Else
21950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21960   End Select
21970   Resume EXITP

End Sub

Public Sub Calendar_Handler_SP(strProc As String, blnFocus As Boolean, blnMouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_SP"

        Dim datStartDate As Date, datEndDate As Date
        Dim strCalled As String, strAction As String
        Dim intPos01 As Integer
        Dim blnRetVal As Boolean

22010   With frm

22020     intPos01 = InStr(strProc, "_")
22030     strCalled = Left(strProc, (intPos01 - 1))
22040     strAction = Mid(strProc, (intPos01 + 1))

22050     Select Case strCalled
          Case "cmdCalendar1"

22060       Select Case strAction
            Case "GotFocus"
22070         blnFocus = True
22080         .cmdCalendar1_raised_semifocus_dots_img.Visible = True
22090         .cmdCalendar1_raised_img.Visible = False
22100         .cmdCalendar1_raised_focus_img.Visible = False
22110         .cmdCalendar1_raised_focus_dots_img.Visible = False
22120         .cmdCalendar1_sunken_focus_dots_img.Visible = False
22130         .cmdCalendar1_raised_img_dis.Visible = False
22140       Case "MouseDown"
22150         blnMouseDown = True
22160         .cmdCalendar1_sunken_focus_dots_img.Visible = True
22170         .cmdCalendar1_raised_img.Visible = False
22180         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
22190         .cmdCalendar1_raised_focus_img.Visible = False
22200         .cmdCalendar1_raised_focus_dots_img.Visible = False
22210         .cmdCalendar1_raised_img_dis.Visible = False
22220       Case "MouseMove"
22230         If blnMouseDown = False Then
22240           Select Case blnFocus
                Case True
22250             .cmdCalendar1_raised_focus_dots_img.Visible = True
22260             .cmdCalendar1_raised_focus_img.Visible = False
22270           Case False
22280             .cmdCalendar1_raised_focus_img.Visible = True
22290             .cmdCalendar1_raised_focus_dots_img.Visible = False
22300           End Select
22310           .cmdCalendar1_raised_img.Visible = False
22320           .cmdCalendar1_raised_semifocus_dots_img.Visible = False
22330           .cmdCalendar1_sunken_focus_dots_img.Visible = False
22340           .cmdCalendar1_raised_img_dis.Visible = False
22350         End If
22360       Case "MouseUp"
22370         .cmdCalendar1_raised_focus_dots_img.Visible = True
22380         .cmdCalendar1_raised_img.Visible = False
22390         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
22400         .cmdCalendar1_raised_focus_img.Visible = False
22410         .cmdCalendar1_sunken_focus_dots_img.Visible = False
22420         .cmdCalendar1_raised_img_dis.Visible = False
22430         blnMouseDown = False
22440       Case "LostFocus"
22450         .cmdCalendar1_raised_img.Visible = True
22460         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
22470         .cmdCalendar1_raised_focus_img.Visible = False
22480         .cmdCalendar1_raised_focus_dots_img.Visible = False
22490         .cmdCalendar1_sunken_focus_dots_img.Visible = False
22500         .cmdCalendar1_raised_img_dis.Visible = False
22510         blnFocus = False
22520       Case "Click"
22530         datStartDate = Date
22540         datEndDate = 0
22550         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
22560         If blnRetVal = True Then
22570           .TransDateStart = datStartDate
22580         Else
22590           .TransDateStart = CDate(Format(Date, "mm/dd/yyyy"))
22600         End If
22610         .TransDateStart.SetFocus
22620       End Select

22630     Case "cmdCalendar2"

22640       Select Case strAction
            Case "GotFocus"
22650         blnFocus = True
22660         .cmdCalendar2_raised_semifocus_dots_img.Visible = True
22670         .cmdCalendar2_raised_img.Visible = False
22680         .cmdCalendar2_raised_focus_img.Visible = False
22690         .cmdCalendar2_raised_focus_dots_img.Visible = False
22700         .cmdCalendar2_sunken_focus_dots_img.Visible = False
22710         .cmdCalendar2_raised_img_dis.Visible = False
22720       Case "MouseDown"
22730         blnMouseDown = True
22740         .cmdCalendar2_sunken_focus_dots_img.Visible = True
22750         .cmdCalendar2_raised_img.Visible = False
22760         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
22770         .cmdCalendar2_raised_focus_img.Visible = False
22780         .cmdCalendar2_raised_focus_dots_img.Visible = False
22790         .cmdCalendar2_raised_img_dis.Visible = False
22800       Case "MouseMove"
22810         If blnMouseDown = False Then
22820           Select Case blnFocus
                Case True
22830             .cmdCalendar2_raised_focus_dots_img.Visible = True
22840             .cmdCalendar2_raised_focus_img.Visible = False
22850           Case False
22860             .cmdCalendar2_raised_focus_img.Visible = True
22870             .cmdCalendar2_raised_focus_dots_img.Visible = False
22880           End Select
22890           .cmdCalendar2_raised_img.Visible = False
22900           .cmdCalendar2_raised_semifocus_dots_img.Visible = False
22910           .cmdCalendar2_sunken_focus_dots_img.Visible = False
22920           .cmdCalendar2_raised_img_dis.Visible = False
22930         End If
22940       Case "MouseUp"
22950         .cmdCalendar2_raised_focus_dots_img.Visible = True
22960         .cmdCalendar2_raised_img.Visible = False
22970         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
22980         .cmdCalendar2_raised_focus_img.Visible = False
22990         .cmdCalendar2_sunken_focus_dots_img.Visible = False
23000         .cmdCalendar2_raised_img_dis.Visible = False
23010         blnMouseDown = False
23020       Case "LostFocus"
23030         .cmdCalendar2_raised_img.Visible = True
23040         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
23050         .cmdCalendar2_raised_focus_img.Visible = False
23060         .cmdCalendar2_raised_focus_dots_img.Visible = False
23070         .cmdCalendar2_sunken_focus_dots_img.Visible = False
23080         .cmdCalendar2_raised_img_dis.Visible = False
23090         blnFocus = False
23100       Case "Click"
23110         datStartDate = Date
23120         datEndDate = 0
23130         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
23140         If blnRetVal = True Then
23150           .TransDateEnd = datStartDate
23160         Else
23170           .TransDateEnd = CDate(Format(Date, "mm/dd/yyyy"))
23180         End If
23190         .TransDateEnd.SetFocus
23200       End Select

23210     Case "cmdCalendar3"
23220       Select Case strAction
            Case "GotFocus"
23230         blnFocus = True
23240         .cmdCalendar3_raised_semifocus_dots_img.Visible = True
23250         .cmdCalendar3_raised_img.Visible = False
23260         .cmdCalendar3_raised_focus_img.Visible = False
23270         .cmdCalendar3_raised_focus_dots_img.Visible = False
23280         .cmdCalendar3_sunken_focus_dots_img.Visible = False
23290         .cmdCalendar3_raised_img_dis.Visible = False
23300       Case "MouseDown"
23310         blnMouseDown = True
23320         .cmdCalendar3_sunken_focus_dots_img.Visible = True
23330         .cmdCalendar3_raised_img.Visible = False
23340         .cmdCalendar3_raised_semifocus_dots_img.Visible = False
23350         .cmdCalendar3_raised_focus_img.Visible = False
23360         .cmdCalendar3_raised_focus_dots_img.Visible = False
23370         .cmdCalendar3_raised_img_dis.Visible = False
23380       Case "MouseMove"
23390         If blnMouseDown = False Then
23400           Select Case blnFocus
                Case True
23410             .cmdCalendar3_raised_focus_dots_img.Visible = True
23420             .cmdCalendar3_raised_focus_img.Visible = False
23430           Case False
23440             .cmdCalendar3_raised_focus_img.Visible = True
23450             .cmdCalendar3_raised_focus_dots_img.Visible = False
23460           End Select
23470           .cmdCalendar3_raised_img.Visible = False
23480           .cmdCalendar3_raised_semifocus_dots_img.Visible = False
23490           .cmdCalendar3_sunken_focus_dots_img.Visible = False
23500           .cmdCalendar3_raised_img_dis.Visible = False
23510         End If
23520       Case "MouseUp"
23530         .cmdCalendar3_raised_focus_dots_img.Visible = True
23540         .cmdCalendar3_raised_img.Visible = False
23550         .cmdCalendar3_raised_semifocus_dots_img.Visible = False
23560         .cmdCalendar3_raised_focus_img.Visible = False
23570         .cmdCalendar3_sunken_focus_dots_img.Visible = False
23580         .cmdCalendar3_raised_img_dis.Visible = False
23590         blnMouseDown = False
23600       Case "LostFocus"
23610         .cmdCalendar3_raised_img.Visible = True
23620         .cmdCalendar3_raised_semifocus_dots_img.Visible = False
23630         .cmdCalendar3_raised_focus_img.Visible = False
23640         .cmdCalendar3_raised_focus_dots_img.Visible = False
23650         .cmdCalendar3_sunken_focus_dots_img.Visible = False
23660         .cmdCalendar3_raised_img_dis.Visible = False
23670         blnFocus = False
23680       Case "Click"
23690         datStartDate = Date
23700         datEndDate = 0
23710         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
23720         If blnRetVal = True Then
23730           .AssetListDate = datStartDate
23740         Else
23750           .AssetListDate = CDate(Format(Date, "mm/dd/yyyy"))
23760         End If
23770         .AssetListDate.SetFocus
23780       End Select

23790     End Select
23800   End With

EXITP:
23810   Exit Sub

ERRH:
23820   Select Case ERR.Number
        Case Else
23830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23840   End Select
23850   Resume EXITP

End Sub

Public Sub Calendar_Set_SP(frm As Access.Form, blnAble1 As Boolean, blnAble2 As Boolean, blnAble3 As Boolean)

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Set_SP"

23910   With frm
23920     Select Case blnAble1
          Case True
23930       .cmdCalendar1.Enabled = True
23940       .cmdCalendar1_raised_img.Visible = True
23950       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
23960       .cmdCalendar1_raised_focus_img.Visible = False
23970       .cmdCalendar1_raised_focus_dots_img.Visible = False
23980       .cmdCalendar1_sunken_focus_dots_img.Visible = False
23990       .cmdCalendar1_raised_img_dis.Visible = False
24000     Case False
24010       .cmdCalendar1.Enabled = False
24020       .cmdCalendar1_raised_img_dis.Visible = True
24030       .cmdCalendar1_raised_img.Visible = False
24040       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
24050       .cmdCalendar1_raised_focus_img.Visible = False
24060       .cmdCalendar1_raised_focus_dots_img.Visible = False
24070       .cmdCalendar1_sunken_focus_dots_img.Visible = False
24080     End Select
24090     Select Case blnAble2
          Case True
24100       .cmdCalendar2.Enabled = True
24110       .cmdCalendar2_raised_img.Visible = True
24120       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
24130       .cmdCalendar2_raised_focus_img.Visible = False
24140       .cmdCalendar2_raised_focus_dots_img.Visible = False
24150       .cmdCalendar2_sunken_focus_dots_img.Visible = False
24160       .cmdCalendar2_raised_img_dis.Visible = False
24170     Case False
24180       .cmdCalendar2.Enabled = False
24190       .cmdCalendar2_raised_img_dis.Visible = True
24200       .cmdCalendar2_raised_img.Visible = False
24210       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
24220       .cmdCalendar2_raised_focus_img.Visible = False
24230       .cmdCalendar2_raised_focus_dots_img.Visible = False
24240       .cmdCalendar2_sunken_focus_dots_img.Visible = False
24250     End Select
24260     Select Case blnAble3
          Case True
24270       .cmdCalendar3.Enabled = True
24280       .cmdCalendar3_raised_img.Visible = True
24290       .cmdCalendar3_raised_semifocus_dots_img.Visible = False
24300       .cmdCalendar3_raised_focus_img.Visible = False
24310       .cmdCalendar3_raised_focus_dots_img.Visible = False
24320       .cmdCalendar3_sunken_focus_dots_img.Visible = False
24330       .cmdCalendar3_raised_img_dis.Visible = False
24340     Case False
24350       .cmdCalendar3.Enabled = False
24360       .cmdCalendar3_raised_img_dis.Visible = True
24370       .cmdCalendar3_raised_img.Visible = False
24380       .cmdCalendar3_raised_semifocus_dots_img.Visible = False
24390       .cmdCalendar3_raised_focus_img.Visible = False
24400       .cmdCalendar3_raised_focus_dots_img.Visible = False
24410       .cmdCalendar3_sunken_focus_dots_img.Visible = False
24420     End Select
24430   End With

EXITP:
24440   Exit Sub

ERRH:
24450   Select Case ERR.Number
        Case Else
24460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24470   End Select
24480   Resume EXITP

End Sub

Public Sub EmptyTable_Tmp_SP()

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTable_Tmp_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

        ' ** Let's find somewhere else for these.
24510   Set dbs = CurrentDb
24520   With dbs
          ' ** Empty AssetList.
24530     Set qdf = .QueryDefs("qryStatementParameters_AssetList_09a")
24540     qdf.Execute
24550     Set qdf = Nothing
24560     DoEvents
          ' ** Empty AssetList2.
24570     Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_50")
24580     qdf.Execute
24590     Set qdf = Nothing
24600     DoEvents
          ' ** Empty tmpAssetList1.
24610     Set qdf = .QueryDefs("qryStatementParameters_AssetList_09b")
24620     qdf.Execute
24630     Set qdf = Nothing
24640     DoEvents
          ' ** Empty tmpAssetList2.
24650     Set qdf = .QueryDefs("qryStatementParameters_AssetList_09c")
24660     qdf.Execute
24670     Set qdf = Nothing
24680     DoEvents
          ' ** Empty tmpAssetList4.
24690     Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_51")
24700     qdf.Execute
24710     Set qdf = Nothing
24720     DoEvents
          ' ** Empty tmpAssetList5.
24730     Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_52")
24740     qdf.Execute
24750     Set qdf = Nothing
24760     DoEvents
          ' ** Empty tmpAccountInfo.
24770     Set qdf = .QueryDefs("qryStatementParameters_AssetList_09d")
24780     qdf.Execute
24790     Set qdf = Nothing
24800     DoEvents
          ' ** Empty tmpAccountInfo2.
24810     Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_53")
24820     qdf.Execute
24830     Set qdf = Nothing
24840     DoEvents
          ' ** Empty tmpRelatedAccount_01.
24850     Set qdf = .QueryDefs("qryStatementParameters_AssetList_21")
24860     qdf.Execute
24870     Set qdf = Nothing
24880     DoEvents
          ' ** Empty tmpRelatedAccount_02.
24890     Set qdf = .QueryDefs("qryStatementParameters_AssetList_22")
24900     qdf.Execute
24910     Set qdf = Nothing
24920     DoEvents
          ' ** Empty tmpRelatedAccount_03.
24930     Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_54")
24940     qdf.Execute
24950     Set qdf = Nothing
24960     DoEvents
24970     .Close
24980   End With
24990   Set dbs = Nothing

EXITP:
25000   Set qdf = Nothing
25010   Set dbs = Nothing
25020   Exit Sub

ERRH:
25030   Select Case ERR.Number
        Case Else
25040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25050   End Select
25060   Resume EXITP

End Sub

Public Sub SetArchiveOption_SP(blnTransEnabled As Boolean, blnAssetEnabled As Boolean, lngAcctArchs As Long, arr_varAcctArch As Variant, frm As Access.Form)
' ** Disable and change caption if there are no archived Ledger records.

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "SetArchiveOption_SP"

        Dim blnArchOpts As Boolean
        Dim lngTpp As Long
        Dim lngX As Long

        'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
25110   lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

25120   With frm
25130     .opgAccountNumber_optAll.Enabled = True
          ' ** Only enable if there are transactions in LedgerArchive.
25140     If lngAcctArchs > 0 Then

25150       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Specified account.
25160         blnArchOpts = False
25170         If Nz(.cmbAccounts.Column(CBX_A_ACTNO), vbNullString) <> vbNullString Then
                ' ***********************************************
                ' ** Array: arr_varAcctArch()
                ' **
                ' **   Field  Element  Name         Constant
                ' **   =====  =======  ===========  ===========
                ' **     1       0     accountno    AR_ACTNO
                ' **     2       1     transdate    AR_TDATE
                ' **     3       2     cnt          AR_CNT
                ' **
                ' ***********************************************
25180           For lngX = 0& To (lngAcctArchs - 1&)
25190             If arr_varAcctArch(AR_ACTNO, lngX) = .cmbAccounts.Column(CBX_A_ACTNO) Then
                    ' ** This account's got archived transactions, so normal settings.
25200               blnArchOpts = True
25210               Exit For
25220             End If
25230           Next
25240         Else
                ' ** Nothing selected yet, so normal settings.
25250           blnArchOpts = True
25260         End If
25270       Case .opgAccountNumber_optAll.OptionValue
              ' ** All accounts, normal settings.
25280         blnArchOpts = True
25290       End Select
            ' ** Normal settings.
25300       Select Case blnArchOpts
            Case True

25310         If .chkIncludeArchive_Trans_lbl.Caption <> "Include Arc&hive" Then
25320           .chkIncludeArchive_Trans_lbl.Caption = "Include Arc&hive"
25330           .chkIncludeArchive_Trans_lbl.FontBold = False
25340         End If
25350         If .chkArchiveOnly_Trans_lbl.Caption <> "Archived Transactions Only" Then
25360           .chkArchiveOnly_Trans_lbl.Caption = "Archived Transactions Only"
25370           .chkArchiveOnly_Trans_lbl.FontBold = False
25380         End If

25390         If .chkIncludeArchive_Asset_lbl.Caption <> "Include" Then
25400           .chkIncludeArchive_Asset_lbl.Caption = "Include"
25410           .chkIncludeArchive_Asset_lbl2.Caption = "Archive"
25420           .chkIncludeArchive_Asset_lbl2_dim_hi.Caption = "Archive"
25430           .chkIncludeArchive_Asset = True  ' ** No longer optional.
25440           .chkIncludeArchive_Asset_lbl.FontBold = True
25450           .chkIncludeArchive_Asset_lbl2.FontBold = True
25460           .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = True
25470         End If

25480         Select Case blnTransEnabled
              Case True
                ' ** Archived.
25490           .chkIncludeArchive_Trans.Enabled = True   ' ** Enable if transactions selected.
25500           Select Case .chkIncludeArchive_Trans
                Case True
25510             .chkIncludeArchive_Trans_lbl.FontBold = True
25520           Case False
25530             .chkIncludeArchive_Trans_lbl.FontBold = False
25540           End Select
25550           .chkArchiveOnly_Trans.Enabled = True
25560           Select Case .chkArchiveOnly_Trans
                Case True
25570             .chkArchiveOnly_Trans_lbl.FontBold = True
25580           Case False
25590             .chkArchiveOnly_Trans_lbl.FontBold = False
25600           End Select
25610         Case False
                ' ** Regular.
25620           .chkIncludeArchive_Trans.Enabled = False  ' ** Disable if transactions not selected.
25630           Select Case .chkIncludeArchive_Trans
                Case True
25640             .chkIncludeArchive_Trans_lbl.FontBold = True
25650           Case False
25660             .chkIncludeArchive_Trans_lbl.FontBold = False
25670           End Select
25680           .chkArchiveOnly_Trans.Enabled = False
25690           Select Case .chkArchiveOnly_Trans
                Case True
25700             .chkArchiveOnly_Trans_lbl.FontBold = True
25710           Case False
25720             .chkArchiveOnly_Trans_lbl.FontBold = False
25730           End Select
25740         End Select  ' ** blnTransEnabled.

25750         Select Case blnAssetEnabled
              Case True
                ' ** Include.
25760           .chkIncludeArchive_Asset.Enabled = False   ' ** No longer optional.
25770           .chkIncludeArchive_Asset_lbl2.ForeColor = WIN_CLR_DISF
25780           .chkIncludeArchive_Asset_lbl2_dim_hi.Visible = True
25790           .chkIncludeArchive_Asset = True   ' ** No longer optional.
25800           Select Case .chkIncludeArchive_Asset
                Case True
25810             .chkIncludeArchive_Asset_lbl.FontBold = True
25820             .chkIncludeArchive_Asset_lbl2.FontBold = True
25830             .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = True
25840           Case False
25850             .chkIncludeArchive_Asset_lbl.FontBold = False
25860             .chkIncludeArchive_Asset_lbl2.FontBold = False
25870             .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = False
25880           End Select
25890         Case False
                ' ** Regular.
25900           .chkIncludeArchive_Asset.Enabled = False  ' ** Disable if asset list not selected.
25910           .chkIncludeArchive_Asset_lbl.FontBold = True
25920           .chkIncludeArchive_Asset_lbl2.FontBold = True
25930           .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = True
25940           .chkIncludeArchive_Asset_lbl2.ForeColor = WIN_CLR_DISF
25950           .chkIncludeArchive_Asset_lbl2_dim_hi.Visible = True
25960           .chkIncludeArchive_Asset = True  ' ** No longer optional.
25970         End Select  ' ** blnAssetEnabled.

25980       Case False

25990         If .chkIncludeArchive_Trans.Enabled = True Then
26000           .chkIncludeArchive_Trans = False
26010           .chkIncludeArchive_Trans_lbl.FontBold = False
26020           .chkIncludeArchive_Trans.Enabled = False
26030         End If
26040         If .chkIncludeArchive_Trans_lbl.Caption <> "No Arc&hived Transactions" Then
26050           .chkIncludeArchive_Trans_lbl.Caption = "No Arc&hived Transactions"
26060         End If
26070         If .chkArchiveOnly_Trans.Enabled = True Then
26080           .chkArchiveOnly_Trans = False
26090           .chkArchiveOnly_Trans_lbl.FontBold = False
26100           .chkArchiveOnly_Trans.Enabled = False
26110         End If
26120         If .chkArchiveOnly_Trans_lbl.Caption <> "For This Account" Then
26130           .chkArchiveOnly_Trans_lbl.Caption = "For This Account"
26140         End If

26150         If .chkIncludeArchive_Asset.Enabled = True Then
26160           .chkIncludeArchive_Asset.Enabled = False
26170           .chkIncludeArchive_Asset_lbl.FontBold = False
26180           .chkIncludeArchive_Asset_lbl2.FontBold = False
26190           .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = False
26200           .chkIncludeArchive_Asset_lbl2.ForeColor = WIN_CLR_DISF
26210           .chkIncludeArchive_Asset_lbl2_dim_hi.Visible = True
26220         End If
26230         If .chkIncludeArchive_Asset_lbl.Caption <> "No Archived Transactions" Then
26240           .chkIncludeArchive_Asset_lbl.Caption = "No Archived Transactions"
26250           .chkIncludeArchive_Asset_lbl2.Caption = "For This Account"
26260           .chkIncludeArchive_Asset_lbl2_dim_hi.Caption = "For This Account"
26270         End If
26280         .chkIncludeArchive_Asset = False
26290         .chkIncludeArchive_Asset_lbl.FontBold = False
26300         .chkIncludeArchive_Asset_lbl2.FontBold = False
26310         .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = False

26320       End Select  ' ** blnArchOpts.

26330     Else

26340       If .chkIncludeArchive_Trans.Enabled = True Then
26350         .chkIncludeArchive_Trans = False
26360         .chkIncludeArchive_Trans_lbl.FontBold = False
26370         .chkIncludeArchive_Trans.Enabled = False
26380       End If
26390       If .chkArchiveOnly_Trans.Enabled = True Then
26400         .chkArchiveOnly_Trans = False
26410         .chkArchiveOnly_Trans_lbl.FontBold = False
26420         .chkArchiveOnly_Trans.Enabled = False
26430       End If
26440       If .chkIncludeArchive_Trans_lbl.Caption <> "There Are No" Then
26450         .chkIncludeArchive_Trans_lbl.Caption = "There Are No"
26460       End If
26470       If .chkArchiveOnly_Trans_lbl.Caption <> "Archived Transactions" Then
26480         .chkArchiveOnly_Trans_lbl.Caption = "Archived Transactions"
26490       End If

26500       If .chkIncludeArchive_Asset.Enabled = True Then
26510         .chkIncludeArchive_Asset.Enabled = False
26520         .chkIncludeArchive_Asset_lbl.FontBold = False
26530         .chkIncludeArchive_Asset_lbl2.FontBold = False
26540         .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = False
26550         .chkIncludeArchive_Asset_lbl2.ForeColor = WIN_CLR_DISF
26560         .chkIncludeArchive_Asset_lbl2_dim_hi.Visible = True
26570       End If
26580       If .chkIncludeArchive_Asset_lbl.Caption <> "There Are No" Then
26590         .chkIncludeArchive_Asset_lbl.Caption = "There Are No"
26600         .chkIncludeArchive_Asset_lbl2.Caption = "Archived Transactions"
26610         .chkIncludeArchive_Asset_lbl2_dim_hi.Caption = "Archived Transactions"
26620         .chkIncludeArchive_Asset_lbl2_dim_hi.Visible = True
26630       End If
26640       .chkIncludeArchive_Asset = False
26650       .chkIncludeArchive_Asset_lbl.FontBold = False
26660       .chkIncludeArchive_Asset_lbl2.FontBold = False
26670       .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = False

26680     End If  ' ** lngAcctArchs.
26690   End With

EXITP:
26700   Exit Sub

ERRH:
26710   Select Case ERR.Number
        Case Else
26720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26730   End Select
26740   Resume EXITP

End Sub

Public Sub UpdateBalanceTable1(frm As Access.Form, blnFromStmts As Boolean, blnTmp01 As Boolean)
' ** Procedures for updating and/or adding to the Balance table.
' ** Skipped for 'Annual Statement'.
' ** Called by:
' **   Statements_Print(), Above.

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateBalanceTable1"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strSQL As String
        Dim dblAccountValue As Double
        Dim datEndDate As Date
        Dim lngRecs As Long
        Dim blnAccountSpecific As Boolean
        Dim lngX As Long

26810   With frm

26820     blnIncludeCurrency = blnTmp01

          '**********************************************************************************
          ' ** Figure out whether we are dealing with one specific account or All accounts.
          '**********************************************************************************
26830     If .cmbAccounts.Enabled = True Then
26840       blnAccountSpecific = True  ' ** This isn't used!
26850     Else
26860       blnAccountSpecific = False
26870     End If
26880     datEndDate = .DateEnd
26890     gdatEndDate = datEndDate
26900     glngMonthID = .cmbMonth.Column(CBX_MON_ID)

          '***************************************************************************
          ' ** This code will generate the values needed to update the Balance table
          ' ** with the current calculations.
          '***************************************************************************

          'QUESTION:
          'qryMaxBalDates IS SIMPLY LOOKS FOR THE LAST BALANCE DATE PRIOR TO THE CURRENTLY-ASKED-FOR END DATE: Balance_Date < DateEnd.
          'SO FOR A 2009 STATEMENT, IT WILL RETURN: 12/31/2008.
          'THE LastBalanceDate FIELD BELOW THEN ADDS 1 DAY TO THAT, RESULTING IN: 01/01/2009.
          'THE transdate RETURNED BELOW ALSO ADDS 1 DAY TO THE qryMaxBalDates DATE, THUS RETURNING: transdate >= 01/01/2009.
          'WHAT EFFECT DOES THE LastBalanceDate FIELD HAVE, BEING THAT IT'S THE 1ST OF THE YEAR?

26910     Set dbs = CurrentDb

          ' ** Reset the SQL for qryTransRangeTotals to work with this date range.
26920     Select Case blnFromStmts
          Case True
26930       Select Case blnIncludeCurrency
            Case True
              ' ** Parameter Query: qryStatementParameters_Balance_ ...
26940       Case False
              ' ** qryStatementParameters_Balance_05_04 (xx).
26950         Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_05_05")
26960         strSQL = qdf.SQL
26970         Set qdf = Nothing
26980         dbs.QueryDefs("qryTransRangeTotals").SQL = strSQL
26990       End Select
27000     Case False
27010       Select Case .chkForeignExchange
            Case True
              ' ** Parameter Query: qryStatementParameters_Balance_07.
27020       Case False
              ' ** qryStatementParameters_Balance_05  THIS IS NOT WHAT'S IN qryStatementParameters_Balance_05!
              'strSQL = "SELECT DISTINCTROW ledger.accountno, Sum(ledger.icash) AS CurrentIcash, " & _
              '  "Sum(ledger.pcash) AS CurrentPcash, Sum(ledger.cost) AS CurrentCost, " & _
              '  "Format(CDate(Format([qryMaxBalDates].[MaxOfbalance date], 'mm/dd/yyyy')) + 1, 'mm/dd/yyyy') As LastBalanceDate " & _
              '  "FROM (account INNER JOIN (ledger LEFT JOIN masterasset ON ledger.assetno = masterasset.assetno) " & _
              '  "ON account.accountno = ledger.accountno) INNER JOIN qryMaxBalDates ON account.accountno = qryMaxBalDates.accountno " & _
              '  "WHERE (((ledger.transdate) >= CDate(Format([qryMaxBalDates].[MaxOfbalance date], 'mm/dd/yyyy')) + 1 " & _
              '  "AND (ledger.transdate) <= #" & Format(datEndDate, "mm/dd/yyyy") & "#)) " & _
              '  "GROUP BY ledger.accountno, Format(CDate(Format([qryMaxBalDates].[MaxOfbalance date], 'mm/dd/yyyy')) + 1, 'mm/dd/yyyy');"
              ' ** qryStatementParameters_Balance_05_04 (xx).
27030         Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_05_05")
27040         strSQL = qdf.SQL
27050         Set qdf = Nothing
27060         dbs.QueryDefs("qryTransRangeTotals").SQL = strSQL
27070       End Select
27080     End Select  ' ** blnFromStmts.

          'dbs.QueryDefs("qryStatementParameters_18").SQL = strSQL
          ' ** Queries this refers to may have been changed by frmCourReportMenu_CA.
          ' ** Uses:
          ' **   qryMaxBalDates: rewritten in SetDateSpecificSQL().
          ' **   qrySumIncreases: rewritten in SetDateSpecificSQL().
          ' **   qrySumDecreases: rewritten in SetDateSpecificSQL().
          ' **   qryTransRangeTotals: rewritten above.
          ' **   qryCurrentTotalMarketValue: uses qryAssetList, rewritten in SetDateSpecificSQL().
          ' **   qryQualifyingAccountsForStatement: rewritten above.

27090     Select Case blnFromStmts
          Case True
27100       Select Case blnIncludeCurrency
            Case True
              ' ** Select Query: qryStatementParameters_Balance_ ...
27110       Case False
              ' ** Account, linked to ActiveAssets, MasterAsset.
27120         Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_05_06")
27130         strSQL = qdf.SQL
27140         Set qdf = Nothing
27150         dbs.QueryDefs("qryAssetList").SQL = strSQL
27160       End Select
27170     Case False
27180       Select Case .chkForeignExchange
            Case True
              ' ** Select Query: qryStatementParameters_Balance_07.
27190       Case False
              ' ** qryAssetList_01_05.
              ' ** VGC 11/25/2009: «='90',-1,1)» to «='90',1,1)».
              ' ** VGC 12/04/2010: Added legalname.
              'strSQL = "SELECT ActiveAssets.assetno, " & _
              '  "masterasset.description AS MasterAssetDescription, " & _
              '  "masterasset.due, masterasset.rate, Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost])) AS TotalCost, " & _
              '  "Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface])) * " & _
              '  "IIf([assettype].[assettype] = '90',1,1) AS TotalShareface, account.accountno, account.shortname, " & _
              '  "account.legalname, assettype.assettype, assettype_description, " & _
              '  "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
              '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & "
              'strSQL = strSQL & "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc, " & _
              '  "account.icash, account.pcash, masterasset.currentDate, " & _
              '  "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]) AS MarketValueX, " & _
              '  "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]) AS MarketValueCurrentX, " & _
              '  "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) AS YieldX, " & CoInfo & " "
              'strSQL = strSQL & "FROM account LEFT JOIN ((masterasset RIGHT JOIN ActiveAssets ON masterasset.assetno = ActiveAssets.assetno) " & _
              '  "LEFT JOIN assettype ON masterasset.assettype = assettype.assettype) ON account.accountno = ActiveAssets.accountno " & _
              '  "GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, account.accountno, " & _
              '  "account.shortname, account.legalname, assettype.assettype, assettype_description, " & _
              '  "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & "
              'strSQL = strSQL & "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
              '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))), account.icash, " & _
              '  "account.pcash, IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]), " & _
              '  "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]), " & _
              '  "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]), account.accountno, masterasset.currentDate;"
              'strSQL = StringReplace(strSQL, "'' As ", "Null As ")  ' ** Module Function: modStringFuncs.
              ' ** Account, linked to ActiveAssets, MasterAsset.
27200         Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_05_06")
27210         strSQL = qdf.SQL
27220         Set qdf = Nothing
27230         dbs.QueryDefs("qryAssetList").SQL = strSQL
27240       End Select
27250     End Select  ' ** blnFromStmts.

          ' ** Empty tmpUpdatedValues.
27260     Set qdf = dbs.QueryDefs("qryStatementParameters_17")
27270     qdf.Execute
27280     Set qdf = Nothing
          ' ** Empty tmpUpdatedValues2.
27290     Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_09")
27300     qdf.Execute
27310     Set qdf = Nothing

27320     Select Case blnFromStmts
          Case True
27330       Select Case blnIncludeCurrency
            Case True
              ' ** Append qryStatementParameters_Balance_20_11 (xx) to tmpUpdatedValues2; Ledger, LedgerArchive.
27340         Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_20_12")
27350       Case False
              ' ** Append Account, linked to qryStatementParameters_20, and
              ' ** add'l tables, to tmpUpdatedValues; Ledger, LedgerArchive.
27360         Set qdf = dbs.QueryDefs("qryStatementParameters_18_08")  ' ** Ledger, LedgerArchive.
27370       End Select
27380     Case False
27390       Select Case .chkForeignExchange
            Case True
              ' ** Append qryStatementParameters_Balance_10_11 (xx) to tmpUpdatedValues2; From .._18.
27400         Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_10_12")
27410         With qdf.Parameters
27420           ![baldat] = datEndDate
27430         End With
27440       Case False
              ' ** Append Account, linked to add'l tables, to tmpUpdatedValues.
              'Set qdf = dbs.QueryDefs("qryStatementParameters_18")    ' ** Ledger Only.
              ' ** Append Account, linked to qryStatementParameters_20, and
              ' ** add'l tables, to tmpUpdatedValues; Ledger, LedgerArchive.
27450         Set qdf = dbs.QueryDefs("qryStatementParameters_18_08")  ' ** Ledger, LedgerArchive.
              ' ** Queries this refers to may have been changed by frmCourReportMenu_CA.
              ' ** Uses:
              ' **   qryMaxBalDates
              ' **   qrySumIncreases             'SET IN modUtilities.SetDateSpecificSQL() IN strQry_SumInc VARIABLE!
              ' **   qrySumDecreases             'SET IN modUtilities.SetDateSpecificSQL() IN strQry_SumDec VARIABLE!
              ' **   qryTransRangeTotals         'SET ABOVE!
              ' **   qryCurrentTotalMarketValue  'THIS IS BASED ON qryAssetList AS SET ABOVE!
              ' **   qryStatementParameters_20   'THIS IS THE SCHEDULED ACCOUNTS QUERY!
27460       End Select
27470     End Select
27480     qdf.Execute
27490     Set qdf = Nothing

          ' ** From qryStatementParameters_18:
          ' **   Pcash: [qryTransRangeTotals].[CurrentPcash]
          ' **   CurrentPcash: IIf(IsNull([qryTransRangeTotals].[CurrentPcash]),0,[qryTransRangeTotals].[CurrentPcash])+[balance].[pcash]
          ' **   Icash: [qryTransRangeTotals].[CurrentIcash]
          ' **   CurrentICash: IIf(IsNull([qryTransRangeTotals].[CurrentIcash]),0,[qryTransRangeTotals].[CurrentIcash])+[balance].[icash]
          ' **   Cost: [qryTransRangeTotals].[CurrentCost]
          ' **   CurrentCost: IIf(IsNull([qryTransRangeTotals].[CurrentCost]),0,[qryTransRangeTotals].[CurrentCost])+[balance].[cost]
          ' **   CurrentTotalMarketValue: ([qryCurrentTotalMarketValue].[TotalMarketValue]+[qryCurrentTotalMarketValue].[IcashAndPcash])
          ' **   PreviousTotalMarketValue: [PreviousTotalMarketValue].[TotalMarketValue]

          ' ** From qryCurrentTotalMarketValue:
          ' **   TotalMarketValue: Sum([qryAssetList].[TotalShareface]*[qryAssetList].[MarketValueCurrentX])
          ' **   IcashAndPcash: IIf(IsNull([qryAssetList].[icash]),0,[qryAssetList].[icash])+IIf(IsNull([qryAssetList].[pcash]),0,[qryAssetList].[pcash])

          ' ** From qryAssetList:
          ' **   icash: [account].[icash]
          ' **   pcash: [account].[pcash]
          ' **   MarketValueCurrentX: IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent])
          ' **   TotalShareface: Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface]))*IIf([assettype].[assettype]='90',1,1)
          ' **   TotalCost: Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost]))

          ' ** Below:
          ' **   dblAccountValue = IIf(IsNull(rst![CurrentTotalMarketValue]), 0, rst![CurrentTotalMarketValue])

          ' ** This code will go through the recordset of values returned by the queries and
          ' ** udpate the balance table with its results.
27500     Select Case blnFromStmts
          Case True
27510       Select Case blnIncludeCurrency
            Case True
27520         Set rst = dbs.OpenRecordset("tmpUpdatedValues2")
27530       Case False
27540         Set rst = dbs.OpenRecordset("tmpUpdatedValues")
27550       End Select
27560     Case False
27570       Select Case .chkForeignExchange
            Case True
27580         Set rst = dbs.OpenRecordset("tmpUpdatedValues2")
27590       Case False
27600         Set rst = dbs.OpenRecordset("tmpUpdatedValues")
27610       End Select
27620     End Select  ' ** blnFromStmts.
27630     If rst.BOF = True And rst.EOF = True Then
            ' ** In your dreams!
27640     Else
27650       rst.MoveLast
27660       lngRecs = rst.RecordCount
27670       rst.MoveFirst
            ' ** Loop through the values until EOF is reached.
27680       For lngX = 1& To lngRecs
27690         Select Case blnFromStmts
              Case True
27700           Select Case blnIncludeCurrency
                Case True
27710             dblAccountValue = IIf(IsNull(rst![CurrentTotalMarketValue_usd]), 0, rst![CurrentTotalMarketValue])
27720             UpdateBalanceTable2 rst![accountno], Format(datEndDate, "mm/dd/yyyy"), _
                    IIf(IsNull(rst![CurrentIcash_usd]), 0, rst![CurrentIcash_usd]), _
                    IIf(IsNull(rst![CurrentPcash_usd]), 0, rst![CurrentPcash_usd]), _
                    IIf(IsNull(rst![CurrentCost_usd]), 0, rst![CurrentCost_usd]), _
                    CStr(dblAccountValue), CStr(dblAccountValue)  ' ** Procedure: Below.
27730           Case False
27740             dblAccountValue = IIf(IsNull(rst![CurrentTotalMarketValue]), 0, rst![CurrentTotalMarketValue])
27750             UpdateBalanceTable2 rst![accountno], Format(datEndDate, "mm/dd/yyyy"), IIf(IsNull(rst![CurrentIcash]), 0, rst![CurrentIcash]), _
                    IIf(IsNull(rst![CurrentPcash]), 0, rst![CurrentPcash]), IIf(IsNull(rst![CurrentCost]), 0, rst![CurrentCost]), _
                    CStr(dblAccountValue), CStr(dblAccountValue)  ' ** Procedure: Below.
27760           End Select
27770         Case False
27780           Select Case .chkForeignExchange
                Case True
27790             dblAccountValue = IIf(IsNull(rst![CurrentTotalMarketValue_usd]), 0, rst![CurrentTotalMarketValue])
27800             UpdateBalanceTable2 rst![accountno], Format(datEndDate, "mm/dd/yyyy"), _
                    IIf(IsNull(rst![CurrentIcash_usd]), 0, rst![CurrentIcash_usd]), _
                    IIf(IsNull(rst![CurrentPcash_usd]), 0, rst![CurrentPcash_usd]), _
                    IIf(IsNull(rst![CurrentCost_usd]), 0, rst![CurrentCost_usd]), _
                    CStr(dblAccountValue), CStr(dblAccountValue)  ' ** Procedure: Below.
27810           Case False
27820             dblAccountValue = IIf(IsNull(rst![CurrentTotalMarketValue]), 0, rst![CurrentTotalMarketValue])
27830             UpdateBalanceTable2 rst![accountno], Format(datEndDate, "mm/dd/yyyy"), IIf(IsNull(rst![CurrentIcash]), 0, rst![CurrentIcash]), _
                    IIf(IsNull(rst![CurrentPcash]), 0, rst![CurrentPcash]), IIf(IsNull(rst![CurrentCost]), 0, rst![CurrentCost]), _
                    CStr(dblAccountValue), CStr(dblAccountValue)  ' ** Procedure: Below.
27840           End Select
27850         End Select  ' ** blnFromStmts.
              ' ** UpdateBalanceTable2(strAccountNumber, strEndDate, strIcash, strPcash, strCost, strTotalMarketValue, strAccountValue)
              ' **   strAccountNumber    = rst![accountno]
              ' **   strEndDate          = Me.DateEnd
              ' **   strIcash            = rst![CurrentIcash]
              ' **   strPcash            = rst![CurrentPcash]
              ' **   strCost             = rst![CurrentCost]
              ' **   strTotalMarketValue = rst![CurrentTotalMarketValue]  |  SAME!
              ' **   strAccountValue     = rst![CurrentTotalMarketValue]  |
              ' ** rst![CurrentTotalMarketValue] =
              ' **   Sum([TotalShareface]*[MarketValueCurrentX]) + IIf(IsNull([icash]),0,[icash])+IIf(IsNull([pcash]),0,[pcash])
              ' **   From qryAssetList.
27860         If lngX < lngRecs Then rst.MoveNext
27870       Next
27880     End If
27890     rst.Close
27900     dbs.Close

27910     blnTmp01 = blnIncludeCurrency

27920   End With

EXITP:
27930   Set rst = Nothing
27940   Set qdf = Nothing
27950   Set dbs = Nothing
27960   Exit Sub

ERRH:
27970   Select Case ERR.Number
        Case Else
27980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27990   End Select
28000   Resume EXITP

End Sub

Public Sub UpdateBalanceTable2(strAccountNo As String, strEndDate As String, strIcash As String, strPcash As String, strCost As String, strTotalMarketValue As String, strAccountValue As String)
' ** This updates the Balance table
' ** Called by:
' **   UpdateBalanceTable1(), Above.

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateBalanceTable2"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim datEndDate As Date

        '*************************************************************************
        ' ** This code will first check to see if the record already exists in
        ' ** the Balance table for this statement ending date.  If it does exist
        ' ** then it will UPDATE the values to what is now current, otherwise it
        ' ** will insert a new record.
        '*************************************************************************

28110   datEndDate = CDate(strEndDate)

28120   Set dbs = CurrentDb

        ' ** Insert a new record for the end balance date into the Balance table.
        ' ** First find out if there is a record for this date already there from printing the Transactions report.

        ' ** Balance, by specified [actno], [baldat].
28130   Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_01")
28140   With qdf.Parameters
28150     ![actno] = strAccountNo
28160     ![baldat] = datEndDate
28170   End With
28180   Set rst = qdf.OpenRecordset
28190   If rst.BOF = True And rst.EOF = True Then
28200     rst.Close
28210     Set rst = Nothing
28220     Set qdf = Nothing
          ' ** Append to Balance, by specified [actno], [baldat], [icsh], [pcsh], [cst], [totmktval], [actval].
28230     Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_02")
28240     With qdf.Parameters
28250       ![actno] = strAccountNo
28260       ![baldat] = datEndDate
28270       ![icsh] = CDbl(strIcash)
28280       ![pcsh] = CDbl(strPcash)
28290       ![CST] = CDbl(strCost)
28300       ![totmktval] = CDbl(strTotalMarketValue)
28310       ![actval] = CDbl(strAccountValue)
28320     End With
          'strSQL2 = "INSERT INTO Balance (accountno, [balance date], icash, pcash, cost, TotalMarketValue, AccountValue) " & _
          '  "VALUES ('" & strAccountNo & "', #" & strEndDate & "#, " & strIcash & ", " & strPcash & ", " & strCost & ", " & _
          '  strTotalMarketValue & ", " & strAccountValue & ");"
28330     qdf.Execute
28340     Set qdf = Nothing
28350   Else
28360     rst.Close
28370     Set rst = Nothing
28380     Set qdf = Nothing
          ' ** Update Balance, by specified [actno], [baldat], [icsh], [pcsh], [cst], [totmktval], [actval].
28390     Set qdf = dbs.QueryDefs("qryStatementParameters_Balance_03")
28400     With qdf.Parameters
28410       ![actno] = strAccountNo
28420       ![baldat] = datEndDate
28430       ![icsh] = CDbl(strIcash)
28440       ![pcsh] = CDbl(strPcash)
28450       ![CST] = CDbl(strCost)
28460       ![totmktval] = CDbl(strTotalMarketValue)
28470       ![actval] = CDbl(strAccountValue)
28480     End With
          'strSQL2 = "UPDATE Balance SET Balance.Icash =  " & strIcash & ", Balance.Pcash =  " & strPcash & ", Balance.Cost = " & _
          '  strCost & ", Balance.TotalMarketValue = " & strTotalMarketValue & ", Balance.AccountValue = " & strAccountValue & " " & _
          '  "WHERE (((Balance.accountno) = '" & strAccountNo & "') AND ((Balance.[balance date])=#" & strEndDate & "#));"
28490     qdf.Execute
28500     Set qdf = Nothing
28510   End If
28520   dbs.Close

EXITP:
28530   Set rst = Nothing
28540   Set qdf = Nothing
28550   Set dbs = Nothing
28560   Exit Sub

ERRH:
28570   Select Case ERR.Number
        Case Else
28580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28590   End Select
28600   Resume EXITP

End Sub

Public Function FirstDate_SP(frm As Access.Form) As Boolean
' ** See if dates entered are earlier than first transaction.
' ** If dates not yet entered, let it go through.

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "FirstDate_SP"

        Dim strFirstDateMsg As String, datFirstDate As Date
        Dim blnFound As Boolean
        Dim strTmp01 As String, datTmp02 As Date, datTmp03 As Date
        Dim lngX As Long
        Dim blnRetVal As Boolean

28710   blnRetVal = True

28720   With frm

28730     Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue

28740       strFirstDateMsg = "There is no data for this report."
28750       .FirstDateMsg_Set strFirstDateMsg  ' ** Form Procedure: frmStatementParameters.

28760       Select Case IsNull(.cmbAccounts.Column(CBX_A_TRXDAT))
            Case True
28770         blnRetVal = False
28780         datFirstDate = DateAdd("y", 1, Date)  ' ** Tomorrow.
28790       Case False
28800         If Trim(.cmbAccounts.Column(CBX_A_TRXDAT)) = vbNullString Then
28810           blnRetVal = False
28820           datFirstDate = DateAdd("y", 1, Date)  ' ** Tomorrow.
28830         Else
28840           datFirstDate = CDate(.cmbAccounts.Column(CBX_A_TRXDAT))
28850         End If
28860       End Select
28870       .FirstDate_Set datFirstDate  ' ** Form Procedure: frmStatementParameters.

28880       If blnRetVal = True Then
28890         If .chkTransactions = True Then
28900           If IsNull(.TransDateStart) = False Then
28910             If CDate(.TransDateStart) < datFirstDate Then  ' ** Starting date is early.
28920               If IsNull(.TransDateEnd) = False Then
28930                 If CDate(.TransDateEnd) < datFirstDate Then  ' ** Ending date is too early.
28940                   blnRetVal = False
28950                 End If
28960               End If
28970             End If
28980           End If
28990         ElseIf .chkAssetList = True Then
29000           If IsNull(.AssetListDate) = False Then
29010             If CDate(.AssetListDate) < datFirstDate Then
29020               blnRetVal = False
29030             End If
29040           End If
29050         ElseIf .chkStatements = True Then
29060           If IsNull(.cmbMonth) = False Then
29070             If IsNull(.StatementsYear) = False Then
29080               strTmp01 = Right("00" & CStr(.cmbMonth.Column(CBX_MON_ID)), 2) & "/01/" & CStr(.StatementsYear)
29090               datTmp02 = CDate(strTmp01)
29100               datTmp02 = DateAdd("m", 1, datTmp02)  ' ** First of next month.
29110               datTmp02 = DateAdd("y", -1, datTmp02)  ' ** Last of this month.
29120               If datTmp02 >= datTmp02 Then
29130                 blnFound = True
29140               End If
29150             End If
29160           End If
29170         End If
29180       End If  ' ** blnRetVal.

29190     Case .opgAccountNumber_optAll.OptionValue

29200       strFirstDateMsg = "There is no data for these reports."
29210       .FirstDateMsg_Set strFirstDateMsg  ' ** Form Procedure: frmStatementParameters.

29220       blnFound = False
29230       For lngX = 0& To (.cmbAccounts.ListCount - 1&)
29240         Select Case IsNull(.cmbAccounts.Column(CBX_A_TRXDAT, lngX))
              Case True
29250           datTmp03 = DateAdd("y", 1, Date)  ' ** Tomorrow.
29260         Case False
29270           If Trim(.cmbAccounts.Column(CBX_A_TRXDAT, lngX)) = vbNullString Then
29280             datTmp03 = DateAdd("y", 1, Date)  ' ** Tomorrow.
29290           Else
29300             datTmp03 = CDate(.cmbAccounts.Column(CBX_A_TRXDAT, lngX))
29310           End If
29320         End Select
29330         If datTmp03 <= Date Then
29340           If .chkTransactions = True Then
29350             If IsNull(.TransDateStart) = False Then
29360               If CDate(.TransDateStart) >= datTmp03 Then
29370                 blnFound = True
29380               Else
29390                 If IsNull(.TransDateEnd) = False Then
29400                   If CDate(.TransDateEnd) >= datTmp03 Then
29410                     blnFound = True
29420                   End If
29430                 End If
29440               End If
29450             End If
29460           ElseIf .chkAssetList = True Then
29470             If IsNull(.AssetListDate) = False Then
29480               If CDate(.AssetListDate) >= datTmp03 Then
29490                 blnFound = True
29500               End If
29510             End If
29520           ElseIf .chkStatements = True Then
29530             If IsNull(.cmbMonth) = False Then
29540               If IsNull(.StatementsYear) = False Then
29550                 strTmp01 = Right("00" & CStr(.cmbMonth.Column(CBX_MON_ID)), 2) & "/01/" & CStr(.StatementsYear)
29560                 datTmp02 = CDate(strTmp01)
29570                 datTmp02 = DateAdd("m", 1, datTmp02)  ' ** First of next month.
29580                 datTmp02 = DateAdd("y", -1, datTmp02)  ' ** Last of this month.
29590                 If datTmp02 >= datTmp03 Then
29600                   blnFound = True
29610                 End If
29620               End If
29630             End If
29640           End If
29650         End If
29660         If blnFound = True Then
                ' ** Just 1 good hit is all we need.
29670           Exit For
29680         End If
29690       Next

29700       If blnFound = False Then
29710         blnRetVal = False
29720       End If

29730     End Select

29740   End With

EXITP:
29750   FirstDate_SP = blnRetVal
29760   Exit Function

ERRH:
29770   blnRetVal = False
29780   Select Case ERR.Number
        Case Else
29790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29800   End Select
29810   Resume EXITP

End Function

Public Function PricingHistory(datAssetListDate As Date) As Boolean
' ** MAKE SURE ALL ASSETS IN MASTER ASSET HAVE THE SAME DATE!

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "PricingHistory"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim datCurrentDate As Date, datPricingDate As Date
        Dim blnContinue As Boolean
        Dim blnRetVal As Boolean

29910   blnRetVal = True
29920   blnContinue = True

        ' ** Because the special queries now have prior value and prior rate,
        ' ** this routine will only establish whether prior numbers are needed.
        ' ** I don't think the dates found here will be necessary.

29930   Set dbs = CurrentDb
29940   With dbs
          ' ** MasterAsset, grouped by currentDate.
29950     Set qdf = .QueryDefs("qryStatementParameters_AssetList_79_01")
29960     Set rst = qdf.OpenRecordset
29970     With rst
29980       .MoveFirst
29990       If IsNull(![currentDate]) = True Then
              ' ** This shouldn't be able to happen!
30000         .MoveLast
30010       End If
30020       datCurrentDate = ![currentDate]
30030       .Close
30040     End With  ' ** rst.
30050     Set rst = Nothing
30060     Set qdf = Nothing
30070     .Close
30080   End With  ' ** dbs.
30090   Set dbs = Nothing

30100   If datAssetListDate < Date Then
30110     If datAssetListDate < datCurrentDate Then
            ' ** Continue to look for earlier pricing.
30120     Else
            ' ** Current pricing is good.
30130       blnContinue = False
30140       Forms("frmStatementParameters").currentDate = datCurrentDate
30150     End If
30160   Else
          ' ** Current pricing is good.
30170     blnContinue = False
30180     Forms("frmStatementParameters").currentDate = datCurrentDate
30190   End If

30200   If blnContinue = True Then
30210     Set dbs = CurrentDb
30220     With dbs
            ' ** qryStatementParameters_AssetList_79_03 (qryStatementParameters_AssetList_79_02
            ' ** (tblPricing_MasterAsset_History, grouped by currentDate, with cnt), all
            ' ** dates <= asset list date, by specified [adat]), grouped, with Max(currentDate).
30230       Set qdf = .QueryDefs("qryStatementParameters_AssetList_79_04")
30240       With qdf.Parameters
30250         ![adat] = datAssetListDate
30260       End With
30270       Set rst = qdf.OpenRecordset
30280       With rst
30290         If .BOF = True And .EOF = True Then
                ' ** No prices prior to the Asset List date, so use current.
30300           blnContinue = False
30310           Forms("frmStatementParameters").currentDate = datCurrentDate
30320         Else
30330           .MoveFirst
30340           datPricingDate = ![currentDate]
30350         End If
30360         .Close
30370       End With  ' ** rst.
30380       Set rst = Nothing
30390       Set qdf = Nothing
30400       .Close
30410     End With  ' ** dbs.
30420     Set dbs = Nothing
30430   End If  ' ** blnContinue.

        'ALL QUERIES USING MASTER ASSET HAVE TO USE PRICING HISTORY!
        'ALL QUERIES USING CURRENCY HAVE TO USE CURRENCY HISTORY!

        ' ** tmpAssetList5 is strictly a foreign exchange table, and not used otherwise.
30440   If blnContinue = True Then
          ' ** datPricingDate is the most recent pricing date prior to or equal to the Asset List date.

30450     Forms("frmStatementParameters").currentDate = datPricingDate
          ' ** Currency rates are not saved in tblCurrency_History by one common date.
          ' ** Every currency may have a different curr_date,
          ' ** so there's no way to find a single date here.

          ' ** Queries and procedures populating tmpAssetList5:
          ' **   qryStatementParameters_AssetList_76_02   ' ** For no assets.
          ' **   FillAListTmp_SP(dbs, rst, "tmpAssetList5")  ' ** These are just Ledger queries, and don't link to ActiveAssets or MasterAsset.
          ' **     qryStatementParameters_AssetList_08a
          ' **     qryStatementParameters_AssetList_08aq
          ' **     qryStatementParameters_AssetList_08b
          ' **     qryStatementParameters_AssetList_08c
          ' **     qryStatementParameters_AssetList_08d
          ' **     qryStatementParameters_AssetList_08e
          ' **     qryStatementParameters_AssetList_08deq
          ' **     qryStatementParameters_AssetList_08f
          ' **     qryStatementParameters_AssetList_08fq
          ' **     qryStatementParameters_AssetList_08h
          ' **     qryStatementParameters_AssetList_08i
          ' **     qryStatementParameters_AssetList_08j
          ' **     qryStatementParameters_AssetList_08k

30460   End If  ' ** blnContinue.

        ' ** blnContinue = False signifies that no prior values are present and/or needed.
        ' ** So I'll return blnRetVal = False for 'Don't use special queries.'
30470   If blnContinue = False Then
30480     blnRetVal = False
30490   End If

EXITP:
30500   Set rst = Nothing
30510   Set qdf = Nothing
30520   Set dbs = Nothing
30530   PricingHistory = blnRetVal
30540   Exit Function

ERRH:
30550   blnRetVal = False
30560   Select Case ERR.Number
        Case Else
30570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30580   End Select
30590   Resume EXITP

End Function

Public Sub SetAccountLastStatement(frm As Access.Form)

30600 On Error GoTo ERRH

        Const THIS_PROC As String = "SetAccountLastStatement"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strCap As String

30610   strCap = vbNullString

30620   With frm
30630     If IsNull(.cmbAccounts) = False Then
30640       If Trim(.cmbAccounts) <> vbNullString Then
30650         Set dbs = CurrentDb
              ' ** Balance table, grouped by accountno, with Max(Balance_Date), by specified [actno].
30660         Set qdf = dbs.QueryDefs("qryStatementParameters_09")
30670         With qdf.Parameters
30680           ![actno] = frm.cmbAccounts
30690         End With
30700         Set rst = qdf.OpenRecordset
30710         With rst
30720           .MoveFirst  ' ** There should be at least one because of CheckDates().
30730           strCap = Format(![balance_date], "mm/dd/yyyy")
30740           .Close
30750         End With
30760         dbs.Close
30770       End If
30780     End If
30790     .cmbAccounts_lbl3.Caption = strCap
30800   End With

EXITP:
30810   Set rst = Nothing
30820   Set qdf = Nothing
30830   Set dbs = Nothing
30840   Exit Sub

ERRH:
30850   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore.
30860     frm.cmbAccounts_lbl3.Caption = vbNullString
30870   Case Else
30880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30890   End Select
30900   Resume EXITP

End Sub

Public Sub Month_AfterUpdate_SP(frm As Access.Form)

31000 On Error GoTo ERRH

        Const THIS_PROC As String = "Month_AfterUpdate_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

31010   With frm

31020     DoCmd.Hourglass True
31030     DoEvents

31040     blnRetVal = True

31050     Set dbs = CurrentDb
31060     With dbs
            ' ** Empty tmpAssetList1.
31070       Set qdf = .QueryDefs("qryStatementParameters_AssetList_09b")
31080       qdf.Execute
31090       Set qdf = Nothing
            ' ** Empty tmpAssetList2.
31100       Set qdf = .QueryDefs("qryStatementParameters_AssetList_09c")
31110       qdf.Execute
31120       Set qdf = Nothing
            ' ** Empty tmpAssetList4.
31130       Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_51")
31140       qdf.Execute
31150       Set qdf = Nothing
            ' ** Empty tmpAssetList5.
31160       Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_52")
31170       qdf.Execute
31180       Set qdf = Nothing
31190       .Close
31200     End With
31210     Set dbs = Nothing

          ' ** Make sure company info variables are set.
31220     CoOptions_Read  ' ** Module Function: modStartupFuncs.
          'Runs Qrys, initializes variables.

31230     If IsNull(.StatementsYear) = True Then
31240       .StatementsYear = year(DateAdd("m", -1, Now()))
31250     End If

31260     Select Case .cmbMonth.Column(CBX_MON_NAME)
          Case "January"
31270       .DateEnd = "01/31/" & .StatementsYear
31280     Case "February"
31290       .DateEnd = Format(CDate("03/01/" & .StatementsYear) - 1, "mm/dd/yyyy")
31300     Case "March"
31310       .DateEnd = "03/31/" & .StatementsYear
31320     Case "April"
31330       .DateEnd = "04/30/" & .StatementsYear
31340     Case "May"
31350       .DateEnd = "05/31/" & .StatementsYear
31360     Case "June"
31370       .DateEnd = "06/30/" & .StatementsYear
31380     Case "July"
31390       .DateEnd = "07/31/" & .StatementsYear
31400     Case "August"
31410       .DateEnd = "08/31/" & .StatementsYear
31420     Case "September"
31430       .DateEnd = "09/30/" & .StatementsYear
31440     Case "October"
31450       .DateEnd = "10/31/" & .StatementsYear
31460     Case "November"
31470       .DateEnd = "11/30/" & .StatementsYear
31480     Case "December"
31490       .StatementsYear = year(Date) - 1
31500       .DateEnd = "12/31/" & .StatementsYear
31510     Case Else
31520       blnRetVal = False
31530       DoCmd.Hourglass False
31540       MsgBox "Please Enter a valid Month.", vbInformation + vbOKOnly, (Left(("Invalid Entry" & Space(40)), 40) & "E01")
31550       .cmbMonth.SetFocus
31560     End Select

31570     If blnRetVal = True Then
31580       DoEvents
31590       .AssetListDate = .DateEnd
            ' ** This is the last date for everyone, and doesn't reflect the individual acct!
31600       varTmp00 = DLookup("[Statement_Date]", "Statement Date")
31610       Select Case IsNull(varTmp00)
            Case True
31620         .DateStart = DateAdd("yyyy", -1, .DateEnd)
31630       Case False
31640         .DateStart = CDate(varTmp00)
31650       End Select
31660     End If

          ' ** Check if any scheduled accounts have foreign currency.
31670     If blnRetVal = True And .HasForeign = True Then
31680       ForEx_ChkScheduled frm  ' ** Procedure: Above.
            'Calls ForExArr_Load()
            '  Runs Qyrs, loads arrays.
            'Calls AcctSched_Load()
            '  Runs Qyrs, loads arrays.
            'Calls chkIncludeCurrency_AfterUpdate()
            '  Sets Bold.
31690       Select Case .chkIncludeCurrency
            Case True
31700         .HasForeign_Sched = "SOME"
31710       Case False
31720         .HasForeign_Sched = "NONE"
31730       End Select
31740     End If

31750     If blnRetVal = True Then
            ' ** If they've changed the month, and there's an account in cmbAccounts,
            ' ** Null it out, otherwise it won't check for scheduled, and will error if not.
31760       If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
31770         If IsNull(.cmbAccounts) = False Then
31780           .cmbAccounts = Null
31790         End If
31800       End If
31810     End If

31820     DoCmd.Hourglass False

31830   End With

EXITP:
31840   Set qdf = Nothing
31850   Set dbs = Nothing
31860   Exit Sub

ERRH:
31870   DoCmd.Hourglass False
31880   Select Case ERR.Number
        Case Else
31890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31900   End Select
31910   Resume EXITP

End Sub

Public Sub SetRelatedOption(frm As Access.Form)
' ** Disable and change caption if there are no related accounts.

32000 On Error GoTo ERRH

        Const THIS_PROC As String = "SetRelatedOption"

        Dim blnHasRels As Boolean

32010   blnHasRels = False

32020   With frm

32030     If IsNull(.cmbAccounts) = False Then
32040       If IsNull(.cmbAccounts.Column(CBX_A_HASREL)) = False Then
32050         blnHasRels = CBool(.cmbAccounts.Column(CBX_A_HASREL))
32060       End If
32070     End If

          ' ** .chkTransactions: DISABLED
          ' ** .chkAssetList: ENABLED
          ' ** .chkStatements: DISABLED

32080     Select Case blnHasRels
          Case True
32090       .chkRelatedAccounts.Enabled = .chkAssetList
32100       .chkRelatedAccounts_lbl.Caption = "Include Related Accounts"
32110     Case False
32120       .chkRelatedAccounts.Enabled = False
32130       .chkRelatedAccounts_lbl.Caption = "No Related Accounts"
32140     End Select

32150   End With

EXITP:
32160   Exit Sub

ERRH:
32170   Select Case ERR.Number
        Case Else
32180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32190   End Select
32200   Resume EXITP

End Sub

Public Function SetRelatedAccts(frm As Access.Form) As Variant

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "SetRelatedAccts"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String, strRelAccts As String
        Dim lngAccts As Long, arr_varAcct() As Variant
        Dim lngLen As Long, lngLastComma As Long
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngE As Long
        Dim blnContinue2 As Boolean
        Dim arr_varRetVal As Variant

        ' ** Array: arr_varAcct().
        Const A_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const A_ACTNO As Integer = 0
        Const A_LAST  As Integer = 1

        ' ** Array: arr_varRetVal().
        Const RV_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const RV_ERR As Integer = 0
        Const RV_REL As Integer = 1
        Const RV_IN  As Integer = 2

32310   blnContinue2 = False  ' ** Simpler this way.

32320   ReDim arr_varRetVal(RV_ELEMS, 0)
32330   arr_varRetVal(RV_ERR, 0) = vbNullString

32340   With frm
32350     If .chkRelatedAccounts = True And IsNull(.cmbAccounts) = False Then

32360       strAccountNo = .cmbAccounts

32370       Set dbs = CurrentDb
32380       With dbs

              ' ** Account, by specified [actno].
32390         Set qdf = .QueryDefs("qryStatementParameters_AssetList_20")
32400         With qdf.Parameters
32410           ![actno] = strAccountNo
32420         End With
32430         Set rst = qdf.OpenRecordset
32440         With rst
32450           If .BOF = True And .EOF = True Then
                  ' ** Shouldn't happend.
32460           Else
32470             .MoveFirst
32480             If IsNull(![related_accountno]) = False Then
32490               If Trim(![related_accountno]) <> vbNullString Then
32500                 blnContinue2 = True
32510                 strRelAccts = ![related_accountno]
32520               End If
32530             End If
32540           End If
32550           .Close
32560         End With  ' ** rst.

32570         If blnContinue2 = True Then

32580           lngAccts = 0&
32590           ReDim arr_varAcct(A_ELEMS, 0)

32600           lngLen = Len(strRelAccts)
32610           lngLastComma = 0&
32620           For lngX = lngLen To 1& Step -1&
32630             If Mid(strRelAccts, lngX, 1) = "," Or Mid(strRelAccts, lngX, 1) = ";" Or _
                      Mid(strRelAccts, lngX, 1) = " " Then  ' ** Should be comma.
32640               lngAccts = lngAccts + 1&
32650               lngE = lngAccts - 1&
32660               ReDim Preserve arr_varAcct(A_ELEMS, lngE)
32670               If lngLastComma = 0& Then
32680                 arr_varAcct(A_ACTNO, lngE) = Trim(Mid(strRelAccts, (lngX + 1&)))
32690                 arr_varAcct(A_LAST, lngE) = 0&
32700                 lngLastComma = lngX
32710               Else
32720                 arr_varAcct(A_ACTNO, lngE) = Trim(Mid(strRelAccts, (lngX + 1&), ((lngLastComma - lngX) - 1&)))
32730                 arr_varAcct(A_LAST, lngE) = lngLastComma
32740                 lngLastComma = lngX
32750               End If
32760             End If
32770             If lngX = 1& Then
32780               lngAccts = lngAccts + 1&
32790               lngE = lngAccts - 1&
32800               ReDim Preserve arr_varAcct(A_ELEMS, lngE)
32810               arr_varAcct(A_ACTNO, lngE) = Trim(Left(strRelAccts, (lngLastComma - 1&)))
32820               arr_varAcct(A_LAST, lngE) = lngLastComma
32830             End If
32840           Next
                ' ** We could also check to make sure there aren't any dupes!

                ' ** Empty tmpRelatedAccount_01.
32850           Set qdf = .QueryDefs("qryStatementParameters_AssetList_21")
32860           qdf.Execute
32870           Set qdf = Nothing
                ' ** Empty tmpRelatedAccount_02.
32880           Set qdf = .QueryDefs("qryStatementParameters_AssetList_22")
32890           qdf.Execute
32900           Set qdf = Nothing
                ' ** Empty tmpRelatedAccount_03.
32910           Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_54")
32920           qdf.Execute
32930           Set qdf = Nothing

32940           strTmp01 = vbNullString: strTmp02 = vbNullString
32950           Set rst = .OpenRecordset("tmpRelatedAccount_01", dbOpenDynaset, dbAppendOnly)
32960           With rst
32970             For lngX = 0& To (lngAccts - 1&)
32980               .AddNew
32990               ![accountno] = arr_varAcct(A_ACTNO, lngX)
33000               ![related_accountno] = strRelAccts
33010               .Update
33020               strTmp01 = strTmp01 & arr_varAcct(A_ACTNO, lngX) & ", "
33030               strTmp02 = strTmp02 & "'" & arr_varAcct(A_ACTNO, lngX) & "',"
33040             Next
33050             .Close
33060           End With

33070           strTmp01 = Trim(strTmp01)
33080           If Right(strTmp01, 1) = "," Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
33090           If Right(strTmp02, 1) = "," Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))

33100           arr_varRetVal(RV_REL, 0) = strTmp01
33110           arr_varRetVal(RV_IN, 0) = strTmp02

33120         End If  ' ** blnContinue2.

33130         .Close
33140       End With  ' ** dbs.

33150     End If  ' ** chkRelatedAccounts.
33160   End With  ' ** frm.

EXITP:
33170   Set rst = Nothing
33180   Set qdf = Nothing
33190   Set dbs = Nothing
33200   SetRelatedAccts = arr_varRetVal
33210   Exit Function

ERRH:
33220   arr_varRetVal(RV_ERR, 0) = RET_ERR
33230   Select Case ERR.Number
        Case Else
33240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33250   End Select
33260   Resume EXITP

End Function

Public Sub SetStatementOptions(frm As Access.Form, blnRunPriorStatement As Boolean)
' ** Called by:
' **   chkTransactions_AfterUpdate()
' **   chkAssetList_AfterUpdate()
' **   chkStatements_AfterUpdate()
' **   opgAccountNumber_AfterUpdate()

33300 On Error GoTo ERRH

        Const THIS_PROC As String = "SetStatementOptions"

        Dim strControlName As String, strStmtFld As String
        Dim datLastStmt_ThisAcct As Date, datLastStmt_AllAccts As Date, datStmtRequested As Date
        Dim varTmp00 As Variant

33310 On Error Resume Next
33320   strControlName = Screen.ActiveControl.Name
33330 On Error GoTo ERRH

33340   With frm
33350     blnRunPriorStatement = False
33360     .cmdPrintStatement_Single.Caption = "Reprint Single Statement"
33370     .cmdPrintStatement_Single.ControlTipText = "Reprint Single" & vbCrLf & "Statement - Ctrl+S"
33380     .cmdPrintStatement_Single.StatusBarText = "Reprint Single Statement - Ctrl+S"
33390     If .chkStatements = True Then
33400       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue  ' ** Specified account.
33410         Select Case strControlName
              Case "cmdPrintStatement_All"
33420           .chkStatements.SetFocus
33430         End Select
33440         .cmdPrintStatement_All.Enabled = False
33450         .cmdPrintStatement_Single.Enabled = True
33460         .cmdPrintStatement_Summary.Visible = True
33470         .cmdPrintStatement_Summary.Enabled = True
33480         If IsNull(.cmbAccounts) = False Then
33490           If Trim(.cmbAccounts) <> vbNullString Then
                  ' ** SetAccountLastStatement() should have already been run.
33500             If Trim(.cmbAccounts_lbl3.Caption) <> vbNullString Then
33510               datLastStmt_ThisAcct = CDate(.cmbAccounts_lbl3.Caption)
33520             End If
33530             varTmp00 = DLookup("[Statement_Date]", "Statement Date")
33540             If IsNull(varTmp00) = False Then
33550               datLastStmt_AllAccts = CDate(varTmp00)
33560             End If
33570             datStmtRequested = DateSerial(CLng(.StatementsYear), CLng(.cmbMonth.Column(CBX_MON_ID)), 31)
33580             strStmtFld = "smt" & .cmbMonth.Column(CBX_MON_SHORT)
33590             varTmp00 = DLookup("[" & strStmtFld & "]", "account", "[accountno] = '" & .cmbAccounts & "'")
33600             If IsNull(varTmp00) = False Then
33610               If CBool(varTmp00) = True Then
33620                 If (datLastStmt_ThisAcct < datLastStmt_AllAccts) And (datStmtRequested = datLastStmt_AllAccts) Then
                        ' ** Missed one!
33630                   blnRunPriorStatement = True
33640                   .cmdPrintStatement_Single.Caption = "Print Single Statement"
33650                   .cmdPrintStatement_Single.ControlTipText = "Print Single" & vbCrLf & "Statement - Ctrl+S"
33660                   .cmdPrintStatement_Single.StatusBarText = "Print Single Statement - Ctrl+S"
33670                   .cmdPrintStatement_Summary.Enabled = False
33680                 End If
33690               End If
33700             End If
33710           End If
33720         End If
33730       Case .opgAccountNumber_optAll.OptionValue        ' ** All accounts.
33740         Select Case strControlName
              Case "cmdPrintStatement_Single", "cmdPrintStatement_Summary"
33750           .chkStatements.SetFocus
33760         End Select
33770         .cmdPrintStatement_All.Enabled = True
33780         .cmdPrintStatement_Single.Enabled = False
33790         .cmdPrintStatement_Summary.Visible = True
33800         .cmdPrintStatement_Summary.Enabled = False
33810       End Select
33820     Else  ' ** Not statements.
33830       Select Case strControlName
            Case "cmdPrintStatement_All", "cmdPrintStatement_Single", "cmdPrintStatement_Summary"
33840         .chkStatements.SetFocus
33850       End Select
33860       .cmdPrintStatement_All.Enabled = False
33870       .cmdPrintStatement_Single.Enabled = False
33880       .cmdPrintStatement_Summary.Visible = True
33890       .cmdPrintStatement_Summary.Enabled = False
33900     End If
33910   End With

EXITP:
33920   Exit Sub

ERRH:
33930   Select Case ERR.Number
        Case Else
33940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33950   End Select
33960   Resume EXITP

End Sub

Public Function ChkFirstBal(dbs As DAO.Database, strAccountNo As String, strEndDate As String) As Integer
' ** There was not a balance date prior to the requested date.
' ** So, one must be created. Let's update the Initial Balance record
' ** to reflect the ending of the previous month;
' ** making like the account had been created at that time.
' ** THIS DOES NOT UPDATE THE ACCOUNT TABLE WITH CURRENT DATA!
' ** SEE ALSO: modUtilities.CheckFirstAcctBal().
' ** Called by:
' **   SetQrys_AList_SP(), modStatementParamFuncs3, once each for All, Specific.
' ** Return codes:
' **    0  Success.
' **    1  Success, with Archive.
' **    2  Success, Archive only.
' **   -2  No data.
' **   -4  Date criteria not met.
' **   -9  Error.

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkFirstBal"

        Dim qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim datStartDate_Local As Date
        Dim strNewDate As String
        Dim blnHasArch As Boolean, blnArchOnly As Boolean
        Dim intRetVal As Integer

34010   intRetVal = 0

34020   blnHasArch = False: blnArchOnly = False

        ' ** Ledger, grouped, with transdate_min, by specified [actno].
34030   Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_04")
34040   With qdf.Parameters
34050     ![actno] = strAccountNo
34060   End With
34070   Set rst = qdf.OpenRecordset
34080   If rst.BOF = True And rst.EOF = True Then
          ' ** The account has no transactions in the Ledger.
34090     rst.Close
34100     Set rst = Nothing
34110     Set qdf = Nothing
          ' ** LedgerArchive, grouped, with transdate_min, by specified [actno].
34120     Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_04a")
34130     With qdf.Parameters
34140       ![actno] = strAccountNo
34150     End With
34160     Set rst = qdf.OpenRecordset
34170     If rst.BOF = True And rst.EOF = True Then
            ' ** The account has no transactions period.
34180       intRetVal = -2
34190       rst.Close
34200       Set rst = Nothing
34210       Set qdf = Nothing
34220     Else
34230       blnHasArch = True: blnArchOnly = True
34240       rst.MoveFirst
34250       datStartDate_Local = CDate(rst![transdate_min])
34260       rst.Close
34270       Set rst = Nothing
34280       Set qdf = Nothing
34290     End If
34300   Else
34310     rst.MoveFirst
34320     datStartDate_Local = CDate(rst![transdate_min])
34330     rst.Close
          ' ** LedgerArchive, grouped, with transdate_min, by specified [actno].
34340     Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_04a")
34350     With qdf.Parameters
34360       ![actno] = strAccountNo
34370     End With
34380     Set rst = qdf.OpenRecordset
34390     If rst.BOF = True And rst.EOF = True Then
            ' ** That's OK.
34400       rst.Close
34410       Set rst = Nothing
34420       Set qdf = Nothing
34430     Else
34440       blnHasArch = True
34450       rst.MoveFirst
34460       datStartDate_Local = CDate(rst![transdate_min])
34470       rst.Close
34480       Set rst = Nothing
34490       Set qdf = Nothing
34500     End If
34510   End If

34520   If intRetVal = 0 Then

34530     If datStartDate_Local > CDate(strEndDate) Then
            ' ** No transcations prior to date submitted.
34540       intRetVal = -4
34550     Else

            ' ** There are transactions that are before the requested date.
            ' ** So, let's figure out what date we need to update the record with.
34560       Select Case Format(datStartDate_Local, "m")
            Case "1"
34570         strNewDate = "12/31/" & CStr(CInt(Format(datStartDate_Local, "yyyy")) - 1)
34580       Case "2"
34590         strNewDate = "01/31/" & Format(datStartDate_Local, "yyyy")
34600       Case "3"
34610         strNewDate = "02/" & Format(CDate("03/01/" & Format(datStartDate_Local, "yyyy")) - 1, "dd") & "/" & Format(datStartDate_Local, "yyyy")
34620       Case "4"
34630         strNewDate = "03/31/" & Format(datStartDate_Local, "yyyy")
34640       Case "5"
34650         strNewDate = "04/30/" & Format(datStartDate_Local, "yyyy")
34660       Case "6"
34670         strNewDate = "05/31/" & Format(datStartDate_Local, "yyyy")
34680       Case "7"
34690         strNewDate = "06/30/" & Format(datStartDate_Local, "yyyy")
34700       Case "8"
34710         strNewDate = "07/31/" & Format(datStartDate_Local, "yyyy")
34720       Case "9"
34730         strNewDate = "08/31/" & Format(datStartDate_Local, "yyyy")
34740       Case "10"
34750         strNewDate = "09/30/" & Format(datStartDate_Local, "yyyy")
34760       Case "11"
34770         strNewDate = "10/31/" & Format(datStartDate_Local, "yyyy")
34780       Case "12"
34790         strNewDate = "11/30/" & Format(datStartDate_Local, "yyyy")
34800       End Select

            ' ** Update Balance, by specified [actno], [datnew].
34810       Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_05")
34820       With qdf.Parameters
34830         ![actno] = strAccountNo
34840         ![datnew] = CDate(strNewDate)
34850       End With
34860       qdf.Execute

34870       Select Case blnHasArch
            Case True
34880         Select Case blnArchOnly
              Case True
34890           intRetVal = 2
34900         Case False
34910           intRetVal = 1
34920         End Select
34930       Case False
              ' ** Let stand intRetVal = 0.
34940       End Select

34950     End If
34960   End If

EXITP:
34970   Set rst = Nothing
34980   Set qdf = Nothing
34990   ChkFirstBal = intRetVal
35000   Exit Function

ERRH:
35010   intRetVal = -9
35020   Select Case ERR.Number
        Case Else
35030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35040   End Select
35050   Resume EXITP

End Function

Public Function AnnualStatement_PrevTrans(strAccountNo As String, datLastYearEnd As Date, blnPrintAll As Boolean) As Boolean

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "AnnualStatement_PrevTrans"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim blnTryArchive As Boolean, blnMsgShown As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim blnRetVal As Boolean

35110   blnRetVal = True

35120   lngRecs = 0&: blnTryArchive = False: blnMsgShown = False
35130   Set dbs = CurrentDb
35140   With dbs
          ' ** qryStatementAnnual_02a (Ledger, by specified [actno], [tdat]), grouped, with cnt, transdate_min, transdate_max.
35150     Set qdf = .QueryDefs("qryStatementAnnual_03a")
35160     With qdf.Parameters
35170       ![actno] = strAccountNo
35180       ![tdat] = datLastYearEnd
35190     End With
35200     Set rst = qdf.OpenRecordset
35210     With rst
35220       If .BOF = True And .EOF = True Then
35230         blnTryArchive = True
35240       Else
35250         .MoveFirst
35260         If IsNull(![cnt]) = True Then
35270           blnTryArchive = True
35280         Else
35290           If ![cnt] < 5 Then  ' ** Arbitrary.
35300             blnTryArchive = True
35310             lngRecs = ![cnt]
35320           Else
35330             If (year(![transdate_max]) = year(datLastYearEnd)) And (year(![transdate_min]) < year(datLastYearEnd)) Then
                    ' ** Prior data covers multiple years.
35340               blnRetVal = False
35350               If blnPrintAll = False Then
35360                 MsgBox "No previous year-end balance exists." & vbCrLf & vbCrLf & _
                        "Annual Statement cannot be run.", vbInformation + vbOKOnly, _
                        (Left("Start-of-Year Balance Not Found" & Space(55), 55) & "Z01")
35370               End If
35380             Else
                    ' ** 5 or more transactions from the beginning of their account, so let's try it!
35390               blnTryArchive = True
35400               lngRecs = ![cnt]
35410             End If
35420           End If
35430         End If
35440       End If
35450       .Close
35460     End With  ' ** Ledger: rst.
35470     If blnTryArchive = True Then
            ' ** qryStatementAnnual_02b (LedgerArchive, by specified [actno], [tdat]), grouped, with cnt, transdate_min, transdate_max.
35480       Set qdf = .QueryDefs("qryStatementAnnual_03b")
35490       With qdf.Parameters
35500         ![actno] = strAccountNo
35510         ![tdat] = datLastYearEnd
35520       End With
35530       Set rst = qdf.OpenRecordset
35540       With rst
35550         If .BOF = True And .EOF = True Then
35560           If lngRecs = 0& Then
35570             blnRetVal = False
35580             If blnPrintAll = False Then
35590               MsgBox "There is no data for the specified year-end.", vbInformation + vbOKOnly, _
                      (Left("Insufficient Data" & Space(55), 55) & "Z02")
35600             End If
35610           ElseIf lngRecs < 5& Then
35620             If blnPrintAll = False Then
35630               msgResponse = MsgBox("There " & IIf(lngRecs > 1&, "are", "is") & " only " & _
                      CStr(lngRecs) & " transaction" & IIf(lngRecs > 1&, "s", vbNullString) & _
                      "in the specified year." & vbCrLf & vbCrLf & "Would you still like to run the statement?", _
                      vbQuestion + vbYesNo, _
                      (Left("Very Little Data" & Space(55), 55) & "Z03"))
35640             Else
35650               msgResponse = vbNo
35660             End If
35670             If msgResponse <> vbYes Then
35680               blnRetVal = False
35690               blnMsgShown = True
35700             End If
35710           Else
                  ' ** Nothing in LedgerArchive, but Ledger has enough.
35720           End If
35730         Else
35740           .MoveFirst
35750           If IsNull(![cnt]) = True Then
35760             If lngRecs = 0& Then
35770               blnRetVal = False
35780               If blnPrintAll = False Then
35790                 MsgBox "There is no data for the specified year-end.", vbInformation + vbOKOnly, _
                        (Left("Insufficient Data" & Space(55), 55) & "Z04")
35800               End If
35810             ElseIf lngRecs < 5& Then
35820               If blnPrintAll = False Then
35830                 msgResponse = MsgBox("There " & IIf(lngRecs > 1&, "are", "is") & " only " & _
                        CStr(lngRecs) & " transaction" & IIf(lngRecs > 1&, "s", vbNullString) & _
                        "in the specified year." & vbCrLf & vbCrLf & "Would you still like to run the statement?", _
                        vbQuestion + vbYesNo, _
                        (Left("Very Little Data" & Space(55), 55) & "Z05"))
35840               Else
35850                 msgResponse = vbNo
35860               End If
35870               If msgResponse <> vbYes Then
35880                 blnRetVal = False
35890                 blnMsgShown = True
35900               End If
35910             Else
                    ' ** Nothing in LedgerArchive, but Ledger has enough.
35920             End If
35930           Else
35940             If ![cnt] < 5 Then  ' ** Arbitrary.
35950               If lngRecs = 0& Or lngRecs < 5& Then  ' ** I know, I know...
35960                 If blnPrintAll = False Then
35970                   msgResponse = MsgBox("There " & IIf((lngRecs + ![cnt]) > 1&, "are", "is") & " only " & _
                          CStr(lngRecs + ![cnt]) & " transaction" & IIf((lngRecs + ![cnt]) > 1&, "s", vbNullString) & _
                          "in the specified year." & vbCrLf & vbCrLf & "Would you still like to run the statement?", _
                          vbQuestion + vbYesNo, _
                          (Left("Very Little Data" & Space(55), 55) & "Z06"))
35980                 Else
35990                   msgResponse = vbNo
36000                 End If
36010                 If msgResponse <> vbYes Then
36020                   blnRetVal = False
36030                   blnMsgShown = True
36040                 End If
36050               Else
                      ' ** Sufficient to proceed.
36060               End If
36070             Else
36080               If (year(![transdate_max]) = year(datLastYearEnd)) And (year(![transdate_min]) < year(datLastYearEnd)) Then
                      ' ** Prior data covers multiple years, regardless of lngRecs.
36090                 blnRetVal = False
36100                 If blnPrintAll = False Then
36110                   MsgBox "No previous year-end balance exists." & vbCrLf & vbCrLf & _
                          "Annual Statement cannot be run.", vbInformation + vbOKOnly, _
                          (Left("Start-of-Year Balance Not Found" & Space(55), 55) & "Z07")
36120                 End If
36130               Else
                      ' ** 5 or more transactions from the beginning of their account, so let's try it!
36140               End If
36150             End If
36160           End If
36170         End If
36180         .Close
36190       End With  ' ** LedgerArchive: rst.
36200     End If
36210     .Close
36220   End With  ' ** dbs.

EXITP:
36230   Set rst = Nothing
36240   Set qdf = Nothing
36250   Set dbs = Nothing
36260   AnnualStatement_PrevTrans = blnRetVal
36270   Exit Function

ERRH:
36280   blnRetVal = False
36290   Select Case ERR.Number
        Case Else
36300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
36310   End Select
36320   Resume EXITP

End Function
