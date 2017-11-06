VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks"

'VGC 09/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   ColumnSet()
' #######################################

' #########################
' ## Use VBA_RenumErrh().  39170
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   All Accts:        A {opgPrint_optAll}
' **   Grouped By Acct:  B {opgShow_optAccount}
' **   Close:            C {cmdClose}
' **   Voided Checks:    D {cmdVoidedChecks}
' **   Selected Acct:    E {opgPrint_optSpecified}
' **   Blank Checks:     F {opgPrint_optBlank}
' **   Short Name:       H {ckgDisplay_opt02_ShortName}
' **   Bank Name:        K {ckgDisplay_opt06_BankName}
' **   Select Acct:      L {lbxShortAccountName}
' **   Set MICR:         M {cmdMICRSet}
' **   Next Check Num:   N {NextCheckNumber}
' **   Account Number:   O {ckgDisplay_opt01_AccountNo}
' **   Print Checks:     P {cmdPrintChecks}
' **   Adjust MICR:      R {cmdMICRAdjust}
' **   Sort Order:       S {cmbSortBy}
' **   Check Amount:     T {ckgDisplay_opt08_CheckAmount}
' **   Bank Acct Number: U {ckgDisplay_opt07_BankAcctNum}
' **   Indiv Checks:     V {opgShow_optAll}
' **   Preview Checks:   W {cmdPreviewChecks}
' **   Exit:             X {cmdClose}
' **   Payee:            Y {ckgDisplay_opt05_Payee}

' ** Shortcut F-keys responsive from this form:
' **   10 Pt:            F10 {opgPayeeFont_opt10Pt}
' **   12 Pt:            F12 {opgPayeeFont_opt12Pt}

' ** All permutations of column options:
' **      1  2  5  6  7  8    Opts  Cols  (Double for w/ and w/o 'Printed' column) (Triple for 'opgPrint' options)  288 permutations!
' **      =  =  =  =  =  =    ====  ====
' ** 1.   X                  | 1     3
' ** 2.      X               |  -     -
' ** 3.   X  X               | 2     4
' ** 4.   X     X            |  -     -
' ** 5.   X        X         |  -     -
' ** 6.   X           X      |  -     -
' ** 7.   X              X   |  -     -
' ** 8.      X  X            |  -     -
' ** 9.      X     X         |  -     -
' ** 10.     X        X      |  -     -
' ** 11.     X           X   |  -     -
' ** 12.  X  X  X            | 3     5
' ** 13.  X  X     X         |  -     -
' ** 14.  X  X        X      |  -     -
' ** 15.  X  X           X   |  -     -
' ** 16.  X     X  X         |  -     -
' ** 17.  X     X     X      |  -     -
' ** 18.  X     X        X   |  -     -
' ** 19.  X        X  X      |  -     -
' ** 20.  X        X     X   |  -     -
' ** 21.  X           X  X   |  -     -
' ** 22.     X  X  X         |  -     -
' ** 23.     X  X     X      |  -     -
' ** 24.     X  X        X   |  -     -
' ** 25.     X     X  X      |  -     -
' ** 26.     X     X     X   |  -     -
' ** 27.     X        X  X   |  -     -
' ** 28.  X  X  X  X         | 4     6
' ** 29.  X  X  X     X      |  -     -
' ** 30.  X  X  X        X   |  -     -
' ** 31.  X  X     X  X      |  -     -
' ** 32.  X  X     X     X   |  -     -
' ** 33.  X  X        X  X   |  -     -
' ** 34.  X     X  X  X      |  -     -
' ** 35.  X     X  X     X   |  -     -
' ** 36.  X     X     X  X   |  -     -
' ** 37.  X        X  X  X   |  -     -
' ** 38.     X  X  X  X      |  -     -
' ** 39.     X  X  X     X   |  -     -
' ** 40.     X  X     X  X   |  -     -
' ** 41.     X     X  X  X   |  -     -
' ** 42.  X  X  X  X  X      | 5     7
' ** 43.  X  X  X  X     X   |  -     -
' ** 44.  X  X  X     X  X   |  -     -
' ** 45.  X  X     X  X  X   |  -     -
' ** 46.  X     X  X  X  X   |  -     -
' ** 47.     X  X  X  X  X   |  -     -
' ** 48.  X  X  X  X  X  X   | 6     8

' *****************************************************************************
' ** CheckNum's are only put into the Journal as checks are being printed.
' ** They will not appear on the Check List until after they've been printed.
' ** Though the order of printing could be specified here,
' ** the check numbering procedure, below, didn't use it!
' ** NOW IT DOES!
' ** And, Rpt_GroupLevel_Set() does set order by both fields.
' **
' ** Clear Check List clears both the CheckNum and PrintCheck.
' ** If there are check numbers in the Journal, Previewing the
' ** checks will give them a new set, then clear them all
' ** when finished.
' *****************************************************************************

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** cmbSortBy combo box constants:
Private Const CBX_SORT_FID  As Integer = 0  ' ** fld_id
'Private Const CBX_SORT_DESC As Integer = 1  ' ** fld_description
Private Const CBX_SORT_FNAM As Integer = 2  ' ** fld_name
'Private Const CBX_SORT_CID  As Integer = 3  ' ** ctl_id_parent
Private Const CBX_SORT_CNAM As Integer = 4  ' ** ctl_name_parent

' ** lbxShortAccountName list box constants:
Private Const LBX_CHK_ID     As Integer = 0  ' ** check_id
Private Const LBX_CHK_ACTNO  As Integer = 1  ' ** Account_Number
'Private Const LBX_CHK_SNAME  As Integer = 2  ' ** Short_Name
Private Const LBX_CHK_CHKCNT As Integer = 3  ' ** Check_Count
'Private Const LBX_CHK_LSTCHK As Integer = 4  ' ** Last_Check_Number
'Private Const LBX_CHK_PAYEE  As Integer = 5  ' ** Payee
Private Const LBX_CHK_BNAME  As Integer = 6  ' ** Bank_Name
Private Const LBX_CHK_BACTNO As Integer = 7  ' ** Bank_Account_Number
'Private Const LBX_CHK_AMT    As Integer = 8  ' ** Check_Amount
Private Const LBX_CHK_STAT   As Integer = 9  ' ** ChkStat

Private blnMICRSet_Focus As Boolean, blnMICRSet_MouseDown As Boolean
Private blnMICRAdjust_Focus As Boolean, blnMICRAdjust_MouseDown As Boolean
Private blnSeparateChecking_R_Focus As Boolean, blnSeparateChecking_R_MouseDown As Boolean
Private blnSeparateChecking_L_Focus As Boolean, blnSeparateChecking_L_MouseDown As Boolean
Private blnVoidedChecks_Focus As Boolean, blnVoidedChecks_MouseDown As Boolean

Private Const strSortOrig As String = "qryPrintChecks_01_01a_Account_Number"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine1 As String = "Sort_line1"
Private Const strSortLine2 As String = "Sort_line2"
Private Const strSortLblAsc As String = "Sort_Up_img"
Private Const strSortLblDesc As String = "Sort_Dn_img"
'Private Const strArwUp As String = "Sort_Up_img"
'Private Const strArwDn As String = "Sort_Dn_img"

' ** Array: arr_varItem().
Private lngItems As Long, arr_varItem() As Variant
Private Const I_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const I_CNAM As Integer = 0
Private Const I_TOP  As Integer = 1
Private Const I_LFT  As Integer = 2

Private lngLbx_Width As Long, strLbx_ColWidths As String, lngThisFrm_Width As Long
Private lngLbx_Col00 As Long, lngLbx_Col01 As Long, lngLbx_Col02 As Long, lngLbx_Col03 As Long, lngLbx_Col04 As Long
Private lngLbx_Col05 As Long, lngLbx_Col06 As Long, lngLbx_Col07 As Long, lngLbx_Col08 As Long, lngLbx_Col09 As Long
Private lngClose_Offset As Long, lngLbxBoxOffset As Long, lngFrmWidthOffset As Long, lngFrmMin_Width As Long

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngLbxLbl_Width As Long, lngLbxHLine01_Width As Long, lngLbxVLine03_Left As Long

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long, lngGTR_Offset As Long
Private lngWarnLbl_Top As Long, lngWarnLbl_Height As Long
Private lngJrnlCnt As Long, lngChkCnt As Long, lngChkVoid_Set As Long
Private blnNeedNext As Boolean, blnPrefCheckNum As Boolean, blnSortAgain As Boolean, blnAdmin As Boolean, blnSepChkApplied As Boolean
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim intPos01 As Integer
        Dim lngTmp01 As Long, arr_varTmp02 As Variant, blnTmp03 As Boolean
        Dim lngX As Long, lngY As Long, lngE As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = Trim(.OpenArgs)
150           intPos01 = InStr(strCallingForm, "~")
160           If intPos01 > 0 Then

170             blnIsOpen = True

180             gblnAdmin = CBool(Mid(strCallingForm, (intPos01 + 1)))
190             blnAdmin = gblnAdmin  ' ** gblnAdmin gets wiped somewhere, so let's try this!
200             strCallingForm = Left(strCallingForm, (intPos01 - 1))

210             CLR_DISABLED_FG = CLR_DKGRY
220             CLR_DISABLED_BG = CLR_LTTEAL

230             blnNeedNext = False: blnPrefCheckNum = False: blnSortAgain = False: blnSepChkApplied = False
240             blnMICRSet_Focus = False: blnMICRSet_MouseDown = False: blnMICRAdjust_Focus = False: blnMICRAdjust_MouseDown = False
250             blnSeparateChecking_R_Focus = False: blnSeparateChecking_R_MouseDown = False
260             blnSeparateChecking_L_Focus = False: blnSeparateChecking_L_MouseDown = False
270             blnVoidedChecks_Focus = False: blnVoidedChecks_MouseDown = False

280             blnGTR_Emblem = False: lngGTR_Stat = 0&

290             .form_width_line.Left = 0&  ' ** Just get it out of the way.

300             lngWarnLbl_Top = .PendingJournal_lbl.Top
310             lngWarnLbl_Height = .PendingJournal_lbl.Height
320             lngGTR_Offset = (.Width - .GoToReport_Emblem_01_img.Left)
330             lngSortLbl_Width = .Sort_Up_img.Width
340             lngLbxLbl_Width = .lbxShortAccountName_lbl.Width
350             lngLbxHLine01_Width = .lbxShortAccountName_hline01.Width
360             lngLbxVLine03_Left = .lbxShortAccountName_vline03.Left

370             lngItems = 0&
380             ReDim arr_varItem(I_ELEMS, 0)

390             lngItems = lngItems + 1&
400             lngE = lngItems - 1&
410             ReDim Preserve arr_varItem(I_ELEMS, lngE)  ' ** 0
420             arr_varItem(I_CNAM, lngE) = .SeparateCheckingOpt_arw_r.Name
430             arr_varItem(I_TOP, lngE) = .SeparateCheckingOpt_arw_r.Top
440             arr_varItem(I_LFT, lngE) = .SeparateCheckingOpt_arw_r.Left

450             lngItems = lngItems + 1&
460             lngE = lngItems - 1&
470             ReDim Preserve arr_varItem(I_ELEMS, lngE)  ' ** 1
480             arr_varItem(I_CNAM, lngE) = .SeparateCheckingOpt_arw_r_dots.Name
490             arr_varItem(I_TOP, lngE) = .SeparateCheckingOpt_arw_r_dots.Top
500             arr_varItem(I_LFT, lngE) = .SeparateCheckingOpt_arw_r_dots.Left

510             lngItems = lngItems + 1&
520             lngE = lngItems - 1&
530             ReDim Preserve arr_varItem(I_ELEMS, lngE)  ' ** 2
540             arr_varItem(I_CNAM, lngE) = .SeparateCheckingOpt_arw_l.Name
550             arr_varItem(I_TOP, lngE) = .SeparateCheckingOpt_arw_l.Top
560             arr_varItem(I_LFT, lngE) = .SeparateCheckingOpt_arw_l.Left

570             lngItems = lngItems + 1&
580             lngE = lngItems - 1&
590             ReDim Preserve arr_varItem(I_ELEMS, lngE)  ' ** 3
600             arr_varItem(I_CNAM, lngE) = .SeparateCheckingOpt_arw_l_dots.Name
610             arr_varItem(I_TOP, lngE) = .SeparateCheckingOpt_arw_l_dots.Top
620             arr_varItem(I_LFT, lngE) = .SeparateCheckingOpt_arw_l_dots.Left

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
630             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

640             If gblnGoToReport = True Then
650               If IsNull(garr_varGoToReport(GTR_FRM2)) = False Then
660                 Select Case garr_varGoToReport(GTR_CTL2)
                    Case "cmdMICRAdjust"
670                   .GoToReport_arw_micr_img.Top = .cmdMICRAdjust.Top
680                   .GoToReport_arw_micr_img.Left = ((.cmdMICRAdjust.Left - .GoToReport_arw_micr_img.Width) + (lngTpp))
690                   .GoToReport_arw_micr_img.Visible = True
700                   DoEvents
710                   .cmdMICRSet.Visible = False
720                   .cmdMICRSet_raised_img.Visible = False
730                   .cmdMICRSet_raised_semifocus_dots_img.Visible = False
740                   .cmdMICRSet_raised_focus_img.Visible = False
750                   .cmdMICRSet_raised_focus_dots_img.Visible = False
760                   .cmdMICRSet_sunken_focus_dots_img.Visible = False
770                   .cmdMICRSet_raised_img_dis.Visible = False
780                   .GoToReport_Emblem_01_img.Visible = True
790                   DoEvents
800                 End Select
810               End If
820             End If

830             ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
                  lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
                  lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
                  lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me  ' ** Module Procedure: modCheckPrintFuncs.
                'ColumnSet THIS_PROC  ' ** Procedure: Below.

                ' ** Variables are fed empty, then populated ByRef.
840             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

850             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

860             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
870             If lngMonitorCnt > 1& Then
880               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
890               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
900               If lngMonitorNum > 1& Then
910                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
920               End If
930             End If

                ' ** SeparateCheckingOpt:
                ' **   1  Message applies, box is open.
                ' **   2  Message applies, box is closed.
                ' **   3  Not Applicable (Separate Checking is off, or last opgPrint choice wasn't 1).

940             Select Case gblnSeparateCheckingAccounts
                Case True
950               .opgPrint.DefaultValue = 2
960             Case False
970               .opgPrint.DefaultValue = 1
980               .SeparateCheckingOpt = 3  ' ** N/A.
990             End Select

1000            Set dbs = CurrentDb
1010            Select Case gblnAdmin
                Case True
                  ' ** Journal, grouped, with cnt.
1020              Set qdf = dbs.QueryDefs("qryPrintChecks_02a")
1030            Case False
                  ' ** Journal, grouped, with cnt, for CurrentUser().
1040              Set qdf = dbs.QueryDefs("qryPrintChecks_02b")
1050            End Select
1060            Set rst = qdf.OpenRecordset
1070            rst.MoveFirst
1080            If rst![cnt] = 0& Then
1090              lngJrnlCnt = 0&
1100              lngChkCnt = 0&
1110              rst.Close
1120            Else
1130              rst.MoveLast
1140              lngJrnlCnt = rst.RecordCount
1150              rst.Close
1160              Select Case gblnAdmin
                  Case True
                    ' ** qryPrintChecks_01 (qryPrintChecks_01a (Journal, just 'Misc.', 'Paid', 'Received'),
                    ' ** linked to qryPrintChecks_01b (Journal, grouped by accountno, with CheckCnt), _
                    ' ** grouped by accountno, for PrintCheck = True, sorted), grouped and summed.
1170                Set qdf = dbs.QueryDefs("qryPrintChecks_03a")
1180              Case False
                    ' ** qryPrintChecks_01j (qryPrintChecks_01a (Journal, just 'Misc.', 'Paid', 'Received'),
                    ' ** linked to qryPrintChecks_01b (Journal, grouped by accountno, with CheckCnt),
                    ' ** grouped by accountno, for PrintCheck = True, sorted, for CurrentUser()),
                    ' ** grouped and summed, for CurrentUser().
1190                Set qdf = dbs.QueryDefs("qryPrintChecks_03b")
1200              End Select
1210              Set rst = qdf.OpenRecordset
1220              If rst.BOF = True And rst.EOF = True Then
1230                rst.Close
1240                lngChkCnt = 0&
1250                .opgShow.Enabled = False
1260                .opgShow_lbl.Visible = False
1270                .opgShow_lbl_dim.Visible = True
1280                .opgShow_lbl_dim_hi.Visible = True
1290                .opgShow_optAccount_lbl.Visible = False
1300                .opgShow_optAccount_lbl_dim.Visible = True
1310                .opgShow_optAccount_lbl_dim_hi.Visible = True
1320                .opgShow_optAll_lbl.Visible = False
1330                .opgShow_optAll_lbl_dim.Visible = True
1340                .opgShow_optAll_lbl_dim_hi.Visible = True
1350              Else
1360                rst.MoveFirst
1370                If IsNull(rst![Checks]) = True Then
1380                  rst.Close
1390                  lngChkCnt = 0&
1400                  .opgShow.Enabled = False
1410                  .opgShow_lbl.Visible = False
1420                  .opgShow_lbl_dim.Visible = True
1430                  .opgShow_lbl_dim_hi.Visible = True
1440                  .opgShow_optAccount_lbl.Visible = False
1450                  .opgShow_optAccount_lbl_dim.Visible = True
1460                  .opgShow_optAccount_lbl_dim_hi.Visible = True
1470                  .opgShow_optAll_lbl.Visible = False
1480                  .opgShow_optAll_lbl_dim.Visible = True
1490                  .opgShow_optAll_lbl_dim_hi.Visible = True
1500                Else
1510                  lngChkCnt = rst![Checks]
1520                  rst.Close
1530                  Select Case gblnAdmin
                      Case True
                        ' ** qryPrintChecks_01e (qryPrintChecks_01d (qryPrintChecks_01c (Journal, just 'Misc.',
                        ' ** 'Paid', 'Received', for PrintCheck = True, with Payee), grouped and summed,
                        ' ** by accountno, with Payeex, cnt), with Payee), grouped, with Max(cnt).
1540                    Set qdf = dbs.QueryDefs("qryPrintChecks_01f")
1550                  Case False
                        ' ** qryPrintChecks_01h (qryPrintChecks_01g (qryPrintChecks_01c (Journal, just 'Misc.',
                        ' ** 'Paid', 'Received', for PrintCheck = True, with Payee), grouped and summed, by
                        ' ** accountno, with Payeex, cnt, for CurrentUser()), with Payee, for CurrentUser()),
                        ' ** grouped, with Max(cnt), for CurrentUser().
1560                    Set qdf = dbs.QueryDefs("qryPrintChecks_01i")
1570                  End Select
1580                  Set rst = qdf.OpenRecordset
1590                  rst.MoveFirst
1600                  If rst![Checks] = 1 Then
1610                    .opgShow.Enabled = False
1620                    .opgShow_lbl.Visible = False
1630                    .opgShow_lbl_dim.Visible = True
1640                    .opgShow_lbl_dim_hi.Visible = True
1650                    .opgShow_optAccount_lbl.Visible = False
1660                    .opgShow_optAccount_lbl_dim.Visible = True
1670                    .opgShow_optAccount_lbl_dim_hi.Visible = True
1680                    .opgShow_optAll_lbl.Visible = False
1690                    .opgShow_optAll_lbl_dim.Visible = True
1700                    .opgShow_optAll_lbl_dim_hi.Visible = True
1710                  Else
1720                    .opgShow.Enabled = True
1730                    .opgShow_lbl.Visible = True
1740                    .opgShow_lbl_dim.Visible = False
1750                    .opgShow_lbl_dim_hi.Visible = False
1760                    .opgShow_optAccount_lbl.Visible = True
1770                    .opgShow_optAccount_lbl_dim.Visible = False
1780                    .opgShow_optAccount_lbl_dim_hi.Visible = False
1790                    .opgShow_optAll_lbl.Visible = True
1800                    .opgShow_optAll_lbl_dim.Visible = False
1810                    .opgShow_optAll_lbl_dim_hi.Visible = False
1820                  End If
1830                End If
1840              End If
1850            End If
1860            dbs.Close

1870            Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1880            If .SeparateCheckingOpt <> 3 Then
                  ' ** The message applied the last time they were here.
1890              blnSepChkApplied = True
1900            End If

1910            opgPrint_AfterUpdate  ' ** Procedure: Below.
1920            ckgDisplay_opt01_AccountNo_AfterUpdate  ' ** Procedure: Below.
1930            ckgDisplay_opt02_ShortName_AfterUpdate  ' ** Procedure: Below.
1940            ckgDisplay_opt05_Payee_AfterUpdate  ' ** Procedure: Below.
1950            ckgDisplay_opt06_BankName_AfterUpdate  ' ** Procedure: Below.
1960            ckgDisplay_opt07_BankAcctNum_AfterUpdate  ' ** Procedure: Below.
1970            ckgDisplay_opt08_CheckAmount_AfterUpdate  ' ** Procedure: Below.
1980            chkSyncListSort_AfterUpdate  ' ** Procedure: Below.

1990            .opgPayeeFont.Enabled = True
2000            .opgPayeeFont.Visible = True
2010            .opgPayeeFont_box.Visible = True
2020            .opgPayeeFont_vline01.Visible = True
2030            .opgPayeeFont_vline02.Visible = True
2040            opgPayeeFont_AfterUpdate  ' ** Procedure: Below.

2050            Select Case .opgPrint
                Case .opgPrint_optSpecified.OptionValue, .opgPrint_optBlank.OptionValue
2060              If .lbxShortAccountName.ListCount > 0 Then
2070                .lbxShortAccountName = .lbxShortAccountName.Column(LBX_CHK_ID, 0)
2080              End If
2090            End Select

2100            .ChecksTot = lngChkCnt

2110            If lngChkCnt = 0& Then
2120              If gblnAdmin = False Then
2130                .PendingJournal_lbl.Caption = "THERE ARE NO PENDING TRANSACTIONS FOR USER: " & _
                      CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2140                .PendingCheck_lbl.Caption = "THERE ARE NO CHECKS MARKED FOR PRINTING FOR USER: " & _
                      CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2150              End If
2160              .PendingCheck_lbl.Visible = True
2170              If lngJrnlCnt = 0& Then
2180                .PendingJournal_lbl.Visible = True
2190              End If
2200              If .PendingJournal_lbl.Visible = True Or .PendingCheck_lbl.Visible = True Then
2210                If .ckgDisplay_opt02_ShortName = False And .ckgDisplay_opt05_Payee = False And _
                        .ckgDisplay_opt06_BankName = False Then
2220                  .PendingJournal_lbl.FontSize = 8
2230                  .PendingJournal_lbl.Top = (lngWarnLbl_Top - (16& * lngTpp))
2240                  .PendingJournal_lbl.Height = (lngWarnLbl_Height + (16& * lngTpp))
2250                  .PendingCheck_lbl.FontSize = 8
2260                  .PendingCheck_lbl.Top = ((.PendingJournal_lbl.Top + .PendingJournal_lbl.Height) + (8& * lngTpp))
2270                  .PendingCheck_lbl.Height = (lngWarnLbl_Height + (16& * lngTpp))
2280                Else
2290                  .PendingJournal_lbl.Top = lngWarnLbl_Top
2300                  .PendingJournal_lbl.Height = lngWarnLbl_Height
2310                  .PendingJournal_lbl.FontSize = 9
2320                  .PendingCheck_lbl.Top = ((.PendingJournal_lbl.Top + .PendingJournal_lbl.Height) + (8& * lngTpp))
2330                  .PendingCheck_lbl.Height = lngWarnLbl_Height
2340                  .PendingCheck_lbl.FontSize = 9
2350                End If
2360              End If
2370              .lbxShortAccountName.Enabled = False
2380              .lbxShortAccountName.Visible = False
2390              .lbxShortAccountName_lbl_dim.Visible = True
2400              .lbxShortAccountName_lbl_dim_hi.Visible = True
2410              SubHeadSet_CP False, Me  ' ** Module Procedure: modCheckPrintFuncs.
2420              MICRSet_CP False, Me  ' ** Module Procedure: modCheckPrintFuncs.
2430              .cmdPrintChecks.Enabled = False
2440              .cmdPreviewChecks.Enabled = False
2450              .NextCheckNumber.Enabled = False
2460              .NextCheckNumber.BorderColor = WIN_CLR_DISR
2470              .NextCheckNumber.BackStyle = acBackStyleTransparent
2480              .ChecksTot.Locked = False
2490              .ChecksTot.BorderColor = WIN_CLR_DISR
2500              .cmbSortBy.Enabled = False
2510              .cmbSortBy.BorderColor = WIN_CLR_DISR
2520              .cmbSortBy.BackStyle = acBackStyleTransparent
2530              .chkSyncListSort.Enabled = False
2540              DisplaySet False, Me  ' ** Module Procedure: modCheckPrintFuncs.
2550            End If

2560            Select Case gblnSeparateCheckingAccounts
                Case True
                  ' ** No message.
2570            Case False
2580              .SeparateCheckingOpt_lbl.Visible = True
2590            End Select

2600            If IsNull(.cmbSortBy) = True Then
2610              .cmbSortBy = .cmbSortBy.Column(0, 0)
2620            End If

2630            blnTmp03 = blnIsOpen
2640            blnIsOpen = True
2650            opgPrint_AfterUpdate  ' ** Procedure: Below.  2ND TIME!
2660            blnIsOpen = blnTmp03
2670            If lngChkCnt > 0& Then
2680              lbxShortAccountName_AfterUpdate  ' ** Procedure: Below.
2690            End If

2700            If lngChkCnt = 0& Then
2710              If .NextCheckNumber.Enabled = True Then
                    ' ** Turned back on by then 2nd opgPring_AfterUpdate(), above.
2720                .NextCheckNumber.Enabled = False
2730                .NextCheckNumber.BorderColor = WIN_CLR_DISR
2740                .NextCheckNumber.BackStyle = acBackStyleTransparent
2750                .ChecksTot.Locked = False
2760                .ChecksTot.BorderColor = WIN_CLR_DISR
2770                .cmbSortBy.Enabled = False
2780                .cmbSortBy.BorderColor = WIN_CLR_DISR
2790                .cmbSortBy.BackStyle = acBackStyleTransparent
2800                .chkSyncListSort.Enabled = False
2810              End If
2820            End If

                ' ** It seems to happen somewhere between here and the timer event.
2830            If gblnAdmin <> blnAdmin Then
2840              gblnAdmin = blnAdmin
2850            End If

2860            blnIsOpen = False

2870          Else
2880            Cancel = -1
2890          End If
2900        Else
2910          Cancel = -1
2920        End If
2930      Else
2940        Cancel = -1
2950      End If
2960    End With

2970    DoCmd.Hourglass False

2980    If Cancel = -1 Then
2990      Beep
3000      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
3010    End If

EXITP:
3020    Set rst = Nothing
3030    Set qdf = Nothing
3040    Set dbs = Nothing
3050    Exit Sub

ERRH:
3060    DoCmd.Hourglass False
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub Form_Load()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim varTmp00 As Variant

3110    With Me

3120      If lngChkCnt > 0& Then
3130        SortNow THIS_PROC  ' ** Procedure: Below.
3140        opgShow_AfterUpdate  ' ** Procedure: Below.
3150        .lbxShortAccountName.Requery
3160      Else
3170        .opgShow.Enabled = False
3180        .opgShow_lbl.Visible = False
3190        .opgShow_lbl_dim.Visible = True
3200        .opgShow_lbl_dim_hi.Visible = True
3210        Select Case .opgShow
            Case .opgShow_optAccount.OptionValue
3220          .opgShow_optAccount_lbl.FontBold = True
3230          .opgShow_optAccount_lbl_dim.FontBold = True
3240          .opgShow_optAccount_lbl_dim_hi.FontBold = True
3250          .opgShow_optAll_lbl.FontBold = False
3260          .opgShow_optAll_lbl_dim.FontBold = False
3270          .opgShow_optAll_lbl_dim_hi.FontBold = False
3280        Case .opgShow_optAll.OptionValue
3290          .opgShow_optAccount_lbl.FontBold = False
3300          .opgShow_optAccount_lbl_dim.FontBold = False
3310          .opgShow_optAccount_lbl_dim_hi.FontBold = False
3320          .opgShow_optAll_lbl.FontBold = True
3330          .opgShow_optAll_lbl_dim.FontBold = True
3340          .opgShow_optAll_lbl_dim_hi.FontBold = True
3350        End Select
3360        .opgShow_optAccount_lbl.Visible = False
3370        .opgShow_optAccount_lbl_dim.Visible = True
3380        .opgShow_optAccount_lbl_dim_hi.Visible = True
3390        .opgShow_optAll_lbl.Visible = False
3400        .opgShow_optAll_lbl_dim.Visible = True
3410        .opgShow_optAll_lbl_dim_hi.Visible = True
3420      End If
3430      DoEvents

3440      LoadAlphaSort  ' ** Module Procedure: modCheckPrintFuncs.
3450      DoEvents

3460      varTmp00 = DCount("*", "tblCheckVoid")
3470      Select Case IsNull(varTmp00)
          Case True
3480        .cmdVoidedChecks.Enabled = False
3490      Case False
3500        If varTmp00 = 0 Then
3510          .cmdVoidedChecks.Enabled = False
3520        Else
3530          .cmdVoidedChecks.Enabled = True
3540        End If
3550      End Select
3560      Select Case .cmdVoidedChecks.Enabled
          Case True
3570        .cmdVoidedChecks_raised_img.Visible = True
3580        .cmdVoidedChecks_raised_img_dis.Visible = False
3590      Case False
3600        .cmdVoidedChecks_raised_img_dis.Visible = True
3610        .cmdVoidedChecks_raised_img.Visible = False
3620      End Select
3630      .cmdVoidedChecks_raised_semifocus_dots_img.Visible = False
3640      .cmdVoidedChecks_raised_focus_img.Visible = False
3650      .cmdVoidedChecks_raised_focus_dots_img.Visible = False
3660      .cmdVoidedChecks_sunken_focus_dots_img.Visible = False

3670      If gblnAdmin <> blnAdmin Then
3680        gblnAdmin = blnAdmin
3690      End If

3700      AlreadyPrinted Me  ' ** Module Function: modCheckPrintFuncs.

3710      DoCmd.Hourglass False

3720    End With

EXITP:
3730    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3820    intShiftDown = (Shift And acShiftMask) > 0
3830    intAltDown = (Shift And acAltMask) > 0
3840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   10 Pt:            F10 {opgPayeeFont_opt10Pt}
        ' **   12 Pt:            F12 {opgPayeeFont_opt12Pt}

        ' ** Plain keys.
3850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3860      Select Case intRetVal
          Case vbKeyEscape
3870        intRetVal = 0
3880        cmdClose_Click  ' ** Procedure: Below.
3890      Case vbKeyF10
3900        With Me
3910          intRetVal = 0
3920          .opgPayeeFont.SetFocus
3930          .opgPayeeFont = .opgPayeeFont_opt10Pt.OptionValue
3940          opgPayeeFont_AfterUpdate  ' ** Procedure: Below.
3950        End With
3960      Case vbKeyF12
3970        With Me
3980          intRetVal = 0
3990          .opgPayeeFont.SetFocus
4000          .opgPayeeFont = .opgPayeeFont_opt12Pt.OptionValue
4010          opgPayeeFont_AfterUpdate  ' ** Procedure: Below.
4020        End With
4030      End Select
4040    End If

        ' ** Alt keys.
4050    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
4060      Select Case intRetVal
          Case vbKeyX
4070        intRetVal = 0
4080        cmdClose_Click  ' ** Procedure: Below.
4090      End Select
4100    End If

        ' ** Ctrl-Shift keys.
4110    If intCtrlDown And (Not intAltDown) And intShiftDown Then
4120      Select Case intRetVal
          Case vbKeyF
4130        With Me
4140          intRetVal = 0
4150          .FocusHolder.SetFocus
4160        End With
4170      End Select
4180    End If

EXITP:
4190    KeyCode = intRetVal
4200    Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub Form_Timer()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngTmp01 As Long, arr_varTmp02 As Variant

        Static lngTmp03 As Long, lngTmp04 As Long

4310    With Me
4320      .TimerInterval = 0&
4330      If gblnAdmin <> blnAdmin Then
4340        gblnAdmin = blnAdmin
4350      End If
4360      If blnGTR_Emblem = True Then
4370        lngGTR_Stat = lngGTR_Stat + 1&
4380        If lngTmp03 = 0& And lngTmp04 = 0& Then .GoToReport_Emblem_01_img.Visible = True
4390        lngTmp03 = lngTmp03 + 1&
4400        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp03), 2) & "_img").Visible = True
4410        If lngTmp03 = 1 Then .GoToReport_Emblem_24_img.Visible = False
4420        If lngTmp03 > 1& Then
4430          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp03 - 1&), 2) & "_img").Visible = False
4440        End If
4450        DoEvents
4460        If lngTmp03 = 24& Then
4470          lngTmp03 = 0&
4480          lngTmp04 = lngTmp04 + 1&
4490          If lngTmp04 < 5& Then
4500            .TimerInterval = 50&
4510          Else
4520            blnGTR_Emblem = False
4530          End If
4540        Else
4550          .TimerInterval = 50&
4560        End If
4570      End If  ' ** blnGTR_Emblem.
4580      If gblnSetFocus = True = True Then
4590        gblnSetFocus = False
4600        strFormName = vbNullString
4610  On Error Resume Next
4620        Do Until strFormName = THIS_NAME
4630          DoCmd.SelectObject acForm, THIS_NAME
4640          strFormName = Screen.ActiveForm.Name
4650        Loop
4660  On Error GoTo ERRH
4670        If IsNull(.NextCheckNumber) = False Then
4680          If Val(.NextCheckNumber) > 0 Then
4690            blnPrefCheckNum = True
4700          End If
4710        End If
4720        If gstrReturningForm <> vbNullString Then
4730          Select Case gstrReturningForm
              Case "frmRpt_Checks_MICR_Set"
4740            .cmdMICRSet.SetFocus
4750          Case "frmRpt_Checks_MICR_Adjust"
4760            .cmdMICRAdjust.SetFocus
4770          Case "frmRpt_Checks_Void_Mgmt"
4780            .cmdVoidedChecks.SetFocus
4790          End Select
4800          DoCmd.Hourglass False
4810        End If
4820      End If
4830      If gblnAdmin <> blnAdmin Then
4840        gblnAdmin = blnAdmin
4850      End If
4860      If gblnGoToReport = True Then
4870        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then  ' And lngGTR_Stat < 96&
              ' ** Just let the emblem go.
4880        Else
4890          If blnGTR_Emblem = False And IsNull(garr_varGoToReport(GTR_FRM2)) = False Then
4900            blnGTR_Emblem = True
4910            lngTmp03 = 0&: lngTmp04 = 0&
4920          End If
4930          DoEvents
4940          Select Case blnGoingToReport
              Case True
4950            Select Case blnGoingToReport2
                Case True
4960              .TimerInterval = 0&
4970              DoCmd.Hourglass False  ' ** It should already be off.
4980              gblnGoToReport = False
4990              blnGoingToReport = False
5000              blnGTR_Emblem = False
5010              .GoToReport_arw_rpt_img.Visible = False
                  'lngTmp01 = lngCtls
                  'arr_varTmp02 = arr_varCtl
                  'FontBoxSpecs "MoveBack", lngTmp01, arr_varTmp02, Me  ' ** Module Procedure: modCheckPrintFuncs.
                  ' ** No changes in arr_varCtl(), so no need to update.
5020              DoEvents
5030            Case False
5040              Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                  Case True
5050                blnGoingToReport2 = True
5060                Select Case garr_varGoToReport(GTR_RNAM)
                    Case "rptChecks_Preprinted"
5070                  .GoToReport_arw_prep_img.Visible = False
5080                  .GoToReport_arw_spec_img.Visible = False
5090                  .GoToReport_lbl_arrow01.Visible = False
5100                  .GoToReport_lbl_arrow02.Visible = False
5110                  .GoToReport_lin_vline01.Visible = False
5120                  .GoToReport_lin_vline02.Visible = False
5130                  .GoToReport_lbl_preprint.Visible = False
5140                Case "rptChecks_Blank"
5150                  .GoToReport_arw_blnk_img.Visible = False
5160                End Select
5170                DoEvents
5180                If .Controls(garr_varGoToReport(GTR_CTL6)).Enabled = True Then
5190                  .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
                      'lngTmp01 = lngCtls
                      'arr_varTmp02 = arr_varCtl
                      'FontBoxSpecs "MoveOut", lngTmp01, arr_varTmp02, Me  ' ** Module Procedure: modCheckPrintFuncs.
                      ' ** No changes in arr_varCtl(), so no need to update.
5200                  .GoToReport_arw_rpt_img.Visible = True
5210                  DoEvents
5220                  Beep
5230                  DoCmd.Hourglass False
5240                  DoEvents
5250                  .TimerInterval = (GTR_WAIT * 2&)
5260                Else
5270                  blnGoingToReport = False
5280                  gblnGoToReport = False
5290                  blnGTR_Emblem = False
5300                  Beep
5310                  DoCmd.Hourglass False
5320                  DoEvents
5330                End If
5340              Case False
5350                Select Case garr_varGoToReport(GTR_INV2)
                    Case True
5360                  Select Case garr_varGoToReport(GTR_CTL2)
                      Case "cmdMICRAdjust"
5370                    .TimerInterval = 50&
5380                    cmdMICRAdjust_Click  ' ** Procedure: Below.
5390                  Case Else
                        ' ** There is no else.
5400                    blnGoingToReport = False
5410                    gblnGoToReport = False
5420                    blnGTR_Emblem = False
5430                    GTREmblem_Off  ' ** Procedure: Below.
5440                    Beep
5450                    DoCmd.Hourglass False
5460                    DoEvents
5470                  End Select
5480                Case False
                      ' ** We're done.
5490                  blnGoingToReport = False
5500                  gblnGoToReport = False
5510                  blnGTR_Emblem = False
5520                  GTREmblem_Off  ' ** Procedure: Below.
5530                  Beep
5540                  DoCmd.Hourglass False
5550                  DoEvents
5560                End Select
5570              End Select
5580            End Select
5590          Case False
5600            DoCmd.Hourglass True  ' ** Make sure it's still running.
5610            DoEvents
5620            blnGoingToReport = True
5630            Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                Case True
5640              Select Case garr_varGoToReport(GTR_RNAM)
                  Case "rptChecks_Preprinted"
5650                If .opgPrint = .opgPrint_optBlank.OptionValue Then
5660                  .opgPrint = .opgPrint_optAll.OptionValue
5670                  opgPrint_AfterUpdate  ' ** Procedure: Below.
5680                End If
5690                .opgPrint.SetFocus
5700                .GoToReport_arw_prep_img.Visible = True
5710                .GoToReport_arw_spec_img.Visible = True
5720                .GoToReport_lbl_arrow01.Visible = True
5730                .GoToReport_lbl_arrow02.Visible = True
5740                .GoToReport_lin_vline01.Visible = True
5750                .GoToReport_lin_vline02.Visible = True
5760                .GoToReport_lbl_preprint.Visible = True
5770              Case "rptChecks_Blank"
5780                If .opgPrint <> .opgPrint_optBlank.OptionValue Then
5790                  .opgPrint = .opgPrint_optBlank.OptionValue
5800                  opgPrint_AfterUpdate  ' ** Procedure: Below.
5810                End If
5820                .opgPrint.SetFocus
5830                .GoToReport_arw_blnk_img.Visible = True
5840              End Select
5850              DoEvents
5860              .TimerInterval = (GTR_WAIT * 2&)
5870            Case False
                  ' ** The arrow should already be visible.
5880              If .Controls(garr_varGoToReport(GTR_CTL2)).Enabled = True Then
5890                .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
5900                DoEvents
5910                .TimerInterval = 50&
5920              Else
5930                If .opgPrint.Enabled = True Then
5940                  .opgPrint.SetFocus
5950                  .opgPrint = .opgPrint_optBlank.OptionValue
5960                  opgPrint_AfterUpdate  ' ** Procedure: Below.
5970                  DoEvents
5980                  If .Controls(garr_varGoToReport(GTR_CTL2)).Enabled = True Then
5990                    .Controls(garr_varGoToReport(GTR_CTL2)).SetFocus
6000                    DoEvents
6010                    .TimerInterval = 50&
6020                  Else
                        ' ** Just stop here, nothing else we can do.
6030                    blnGoingToReport = False
6040                    gblnGoToReport = False
6050                    blnGTR_Emblem = False
6060                    GTREmblem_Off  ' ** Procedure: Below.
6070                    Beep
6080                    DoCmd.Hourglass False
6090                    DoEvents
6100                  End If
6110                Else
                      ' ** Just stop here, nothing else we can do.
6120                  blnGoingToReport = False
6130                  gblnGoToReport = False
6140                  blnGTR_Emblem = False
6150                  GTREmblem_Off  ' ** Procedure: Below.
6160                  Beep
6170                  DoCmd.Hourglass False
6180                  DoEvents
6190                End If
6200              End If
6210            End Select
6220          End Select
6230        End If
6240      End If
6250      If gblnAdmin <> blnAdmin Then
6260        gblnAdmin = blnAdmin
6270      End If
6280    End With

EXITP:
6290    Exit Sub

ERRH:
3060    DoCmd.Hourglass False
3070    blnGoingToReport = False
3080    gblnGoToReport = False
3090    Select Case ERR.Number
        Case Else
3100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3110    End Select
3120    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim lngX As Long
        Dim blnSkip As Boolean

6310    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
6320    DoEvents

        ' ** Assure all reports are closed.
6330    Do While Reports.Count > 0
6340      DoCmd.Close acReport, Reports(0).Name
6350    Loop
6360    DoEvents

        ' ** Assure all auxiliary forms are closed.
6370    For lngX = 1& To 6&
6380      Select Case lngX
          Case 1&
6390        strDocName = "frmRpt_Checks_MICR_Set"
6400      Case 2&
6410        strDocName = "frmRpt_Checks_MICR_Adjust"
6420      Case 3&
6430        strDocName = "frmRpt_Checks_Confirm"
6440      Case 4&
6450        strDocName = "frmRpt_Checks_Bank1"
6460      Case 5&
6470        strDocName = "frmRpt_Checks_Bank2"
6480      Case 6&
6490        strDocName = "frmRpt_Checks_Void"
6500      End Select
6510      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
6520        DoCmd.Close acForm, strDocName
6530      End If
6540      DoEvents
6550    Next

        ' ** Array borrowed for voided check process.
6560    glngPrintRpts = 0&
6570    ReDim garr_varPrintRpt(0, 0)

        'blnSkip = True
        'If blnSkip = False Then
6580    Set dbs = CurrentDb
6590    With dbs
            ' ** Empty tblJournal_Check_Staging1.
6600        Set qdf = .QueryDefs("qryPrintChecks_16a")
6610        qdf.Execute
6620        Set qdf = Nothing
6630        DoEvents
            ' ** Empty tblJournal_Check_Staging2.
6640        Set qdf = .QueryDefs("qryPrintChecks_16b")
6650        qdf.Execute
6660        Set qdf = Nothing
6670        DoEvents
            ' ** Empty tblCheckPrint.
            'Set qdf = .QueryDefs("qryPrintChecks_05_32_01")
            'qdf.Execute
            'Set qdf = Nothing
            'DoEvents
            ' ** Empty tblCheckPrint_Report.
            'Set qdf = .QueryDefs("qryPrintChecks_20_06")
            'qdf.Execute
            'Set qdf = Nothing
            'DoEvents
            ' ** Empty tblCheckVoid_Staging.
6680        Set qdf = .QueryDefs("qryPrintChecks_05_13_05")
6690        qdf.Execute
6700        Set qdf = Nothing
6710        DoEvents
            ' ** Empty tmpAccount.
6720        Set qdf = .QueryDefs("qryPrintChecks_53_01")
6730        qdf.Execute
6740        Set qdf = Nothing
6750        DoEvents
6760      .Close
6770    End With
6780    Set dbs = Nothing
        'End If  ' ** blnSkip.

6790    gstrFormQuerySpec = vbNullString
6800    gblnSetFocus = True
6810    gstrReturningForm = THIS_NAME
6820    Select Case strCallingForm
        Case "frmMenu_Post"
6830      If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
6840        DoEvents
6850        DoCmd.SelectObject acForm, strCallingForm, False
6860        Forms(strCallingForm).UpdateTransCnt  ' ** Form Function: frmMenu_Post.
6870        Forms(strCallingForm).TimerInterval = 250&
6880      Else
6890        DoEvents
6900        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
6910      End If
6920    End Select

EXITP:
6930    Set qdf = Nothing
6940    Set dbs = Nothing
6950    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

7010    Detail_Mouse_CP blnMICRSet_Focus, blnMICRAdjust_Focus, _
          blnVoidedChecks_Focus, lngItems, arr_varItem, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
7020    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

7110    Detail_MouseMove Button, Shift, X, Y

EXITP:
7120    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdClose_Click()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

7210    DoCmd.Close acForm, THIS_NAME

EXITP:
7220    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7320    intShiftDown = (Shift And acShiftMask) > 0
7330    intAltDown = (Shift And acAltMask) > 0
7340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7360      Select Case intRetVal
          Case vbKeyTab
7370        With Me
7380          intRetVal = 0
7390          If .opgPrint.Enabled = True Then
7400            .opgPrint.SetFocus
7410          Else
                ' ** Nowhere else to go!
7420          End If
7430        End With
7440      End Select
7450    End If

        ' ** Shift keys.
7460    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7470      Select Case intRetVal
          Case vbKeyTab
7480        With Me
7490          intRetVal = 0
7500          If .cmdPrintChecks.Enabled = True Then
7510            .cmdPrintChecks.SetFocus
7520          ElseIf .cmdPreviewChecks.Enabled = True Then
7530            .cmdPreviewChecks.SetFocus
7540          Else
7550            .opgPayeeFont.SetFocus
7560          End If
7570        End With
7580      End Select
7590    End If

EXITP:
7600    KeyCode = intRetVal
7610    Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub opgPrint_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_box_MouseMove"

7710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7720    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgPrint_AfterUpdate()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_AfterUpdate"

        Dim varTmp00 As Variant

7810    With Me
7820      Select Case .opgPrint
          Case .opgPrint_optAll.OptionValue        ' ** 1.
            ' ** All checks for all accounts, using a single checking account, on preprinted check stock.
7830        DoCmd.Hourglass True
7840        DoEvents
7850        .opgPrint_optAll_lbl.FontBold = True
7860        .opgPrint_optAll_lbl2.FontBold = True
7870        .opgPrint_optSpecified_lbl.FontBold = False
7880        .opgPrint_optSpecified_lbl2.FontBold = False
7890        .opgPrint_optBlank_lbl.FontBold = False
7900        .opgPrint_optAll_box.Visible = True
7910        .opgPrint_optSpecified_box.Visible = False
7920        .opgPrint_optBlank_box.Visible = False
7930        .opgPrint_PrePrinted_box2.Visible = True
7940        .opgPrint_PrePrinted_box3.Visible = True
7950        .opgPrint_optBlank_box2.Visible = False
7960        .opgPrint_optBlank_box3.Visible = False
7970        If lngChkCnt > 0& Then
7980          .lbxShortAccountName.Enabled = True
7990          .lbxShortAccountName.Locked = True
8000          SubHeadSet_CP True, Me  ' ** Module Procedure: modCheckPrintFuncs.
8010        Else
8020          .lbxShortAccountName.Enabled = False
8030          .lbxShortAccountName.Locked = False
8040          SubHeadSet_CP False, Me  ' ** Module Procedure: modCheckPrintFuncs.
8050        End If
8060        .lbxShortAccountName = Null
8070        .lbxShortAccountName.ForeColor = CLR_DISABLED_FG
8080        .lbxShortAccountName.BackColor = CLR_DISABLED_BG
8090        MICRSet_CP False, Me  ' ** Module Procedure: modCheckPrintFuncs.
            ' ** Both Ledger and Journal:
            ' ** qryPrintChecks_33a (qryPrintChecks_32c (Union of qryPrintChecks_31a
            ' ** (qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
            ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by
            ' ** accountno, with Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b
            ' ** (qryPrintChecks_32a (Journal, just CheckNum <> Null), grouped by accountno,
            ' ** with Max(CheckNum), cnt)), grouped by accountno, with Max(CheckNum), cnt;
            ' ** for AccountNo), grouped, with Max(CheckNum), cnt; for All.
8100        varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33b")
8110        If IsNull(varTmp00) = False Then
8120          If varTmp00 > 0 Then
8130            .NextCheckNumber = CLng(varTmp00 + 1)
8140          End If
8150        End If
8160        DoCmd.Hourglass False
8170      Case .opgPrint_optSpecified.OptionValue  ' ** 2.
            ' ** All checks for a selected account, using a single checking account, on preprinted check stock.
8180        DoCmd.Hourglass True
8190        DoEvents
8200        .opgPrint_optAll_lbl.FontBold = False
8210        .opgPrint_optAll_lbl2.FontBold = False
8220        .opgPrint_optSpecified_lbl.FontBold = True
8230        .opgPrint_optSpecified_lbl2.FontBold = True
8240        .opgPrint_optBlank_lbl.FontBold = False
8250        .opgPrint_optAll_box.Visible = False
8260        .opgPrint_optSpecified_box.Visible = True
8270        .opgPrint_optBlank_box.Visible = False
8280        .opgPrint_PrePrinted_box2.Visible = True
8290        .opgPrint_PrePrinted_box3.Visible = True
8300        .opgPrint_optBlank_box2.Visible = False
8310        .opgPrint_optBlank_box3.Visible = False
8320        If lngChkCnt > 0& Then
8330          .lbxShortAccountName.Enabled = True
8340          .lbxShortAccountName.Locked = False
8350          .lbxShortAccountName.ForeColor = CLR_VDKGRY
8360          .lbxShortAccountName.BackColor = CLR_WHT
8370          SubHeadSet_CP True, Me  ' ** Module Procedure: modCheckPrintFuncs.
8380          .lbxShortAccountName = .lbxShortAccountName.Column(LBX_CHK_ID, 0)
8390        Else
8400          .lbxShortAccountName.Enabled = False
8410          .lbxShortAccountName.Locked = False
8420          .lbxShortAccountName = Null
8430          .lbxShortAccountName.ForeColor = WIN_CLR_DISF
8440          .lbxShortAccountName.BackColor = CLR_WHT  ' ** But disabled changes this.
8450          SubHeadSet_CP False, Me  ' ** Module Procedure: modCheckPrintFuncs.
8460        End If
8470        MICRSet_CP False, Me  ' ** Module Procedure: modCheckPrintFuncs.
            ' ** Both Ledger and Journal:
            ' ** qryPrintChecks_33a (qryPrintChecks_32c (Union of qryPrintChecks_31a
            ' ** (qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
            ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by
            ' ** accountno, with Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b
            ' ** (qryPrintChecks_32a (Journal, just CheckNum <> Null), grouped by accountno,
            ' ** with Max(CheckNum), cnt)), grouped by accountno, with Max(CheckNum), cnt;
            ' ** for AccountNo), grouped, with Max(CheckNum), cnt; for All.
8480        varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33b")
8490        If IsNull(varTmp00) = False Then
8500          If varTmp00 > 0 Then
8510            .NextCheckNumber = CLng(varTmp00 + 1)
8520          End If
8530        End If
8540        DoCmd.Hourglass False
8550      Case .opgPrint_optBlank.OptionValue      ' ** 3.
            ' ** All checks for a selected account, on blank stock; requires MICR font.
8560        DoCmd.Hourglass True
8570        DoEvents
8580        .opgPrint_optAll_lbl.FontBold = False
8590        .opgPrint_optAll_lbl2.FontBold = False
8600        .opgPrint_optSpecified_lbl.FontBold = False
8610        .opgPrint_optSpecified_lbl2.FontBold = False
8620        .opgPrint_optBlank_lbl.FontBold = True
8630        .opgPrint_optAll_box.Visible = False
8640        .opgPrint_optSpecified_box.Visible = False
8650        .opgPrint_optBlank_box.Visible = True
8660        .opgPrint_PrePrinted_box2.Visible = False
8670        .opgPrint_PrePrinted_box3.Visible = False
8680        .opgPrint_optBlank_box2.Visible = True
8690        .opgPrint_optBlank_box3.Visible = True
8700        If lngChkCnt > 0& Then
8710          .lbxShortAccountName.Enabled = True
8720          .lbxShortAccountName.Locked = False
8730          .lbxShortAccountName.ForeColor = CLR_VDKGRY
8740          .lbxShortAccountName.BackColor = CLR_WHT
8750          SubHeadSet_CP True, Me  ' ** Module Procedure: modCheckPrintFuncs.
8760          .lbxShortAccountName = .lbxShortAccountName.Column(LBX_CHK_ID, 0)
8770        Else
8780          .lbxShortAccountName.Enabled = False
8790          .lbxShortAccountName.Locked = False
8800          .lbxShortAccountName = Null
8810          .lbxShortAccountName.ForeColor = WIN_CLR_DISF
8820          .lbxShortAccountName.BackColor = CLR_WHT  ' ** But disabled changes this.
8830          SubHeadSet_CP False, Me  ' ** Module Procedure: modCheckPrintFuncs.
8840        End If
8850        MICRSet_CP True, Me  ' ** Module Procedure: modCheckPrintFuncs.
8860        DoCmd.Hourglass False
8870      Case Else
8880        MsgBox "Please select the Check Print Method.", vbInformation + vbOKOnly, "Select Check Print Method"
8890      End Select

          ' ** SeparateCheckingOpt:
          ' **   1  Message applies, box is open.
          ' **   2  Message applies, box is closed.
          ' **   3  Not Applicable (Separate Checking is off, or last opgPrint choice wasn't 1).

8900      If gblnSeparateCheckingAccounts = True And .opgPrint = .opgPrint_optAll.OptionValue Then
8910        If blnSepChkApplied = True And blnIsOpen = True Then
8920          Select Case .SeparateCheckingOpt
              Case 1
8930            SeparateCheckingOpt_arw_r_cmd_Click  ' ** Procedure: Below
8940          Case 2
8950            SeparateCheckingOpt_arw_l_cmd_Click  ' ** Procedure: Below
8960          End Select
8970        Else
8980          SeparateCheckingOpt_arw_r_cmd_Click  ' ** Procedure: Below
8990          .SeparateCheckingOpt = 1  ' ** Open.
9000        End If
9010      Else
9020        SeparateCheckingOpt_arw_l_cmd_Click  ' ** Procedure: Below
9030        .SeparateCheckingOpt_arw_r_box.Visible = False
9040        .SeparateCheckingOpt_arw_r.Visible = False
9050        .SeparateCheckingOpt_arw_r_cmd.Visible = False
9060        .SeparateCheckingOpt = 3  ' ** N/A.
9070      End If

9080      blnIsOpen = True
9090      lbxShortAccountName_AfterUpdate  ' ** Procedure: Below.
9100      blnIsOpen = False

9110    End With

EXITP:
9120    Exit Sub

ERRH:
3060    DoCmd.Hourglass False
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub opgPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_MouseMove"

9210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9220    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgPrint_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9320    intShiftDown = (Shift And acShiftMask) > 0
9330    intAltDown = (Shift And acAltMask) > 0
9340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9370        With Me
9380          intRetVal = 0
9390          If .cmdMICRSet.Enabled = True Then
9400            .cmdMICRSet.SetFocus
9410          ElseIf .NextCheckNumber.Enabled = True Then
9420            .NextCheckNumber.SetFocus
9430          ElseIf .cmbSortBy.Enabled = True Then
9440            .cmbSortBy.SetFocus
9450          ElseIf .lbxShortAccountName.Enabled = True Then
9460            .lbxShortAccountName.SetFocus
9470          ElseIf .opgShow.Enabled = True Then
9480            .opgShow.SetFocus
9490          ElseIf .ckgDisplay_opt01_AccountNo.Enabled = True Then
9500            .ckgDisplay_opt01_AccountNo.SetFocus
9510          ElseIf .opgPayeeFont.Enabled = True Then
9520            .opgPayeeFont.SetFocus
9530          ElseIf .cmdPreviewChecks.Enabled = True Then
9540            .cmdPreviewChecks.SetFocus
9550          Else
9560            .cmdClose.SetFocus
9570          End If
9580        End With
9590      End Select
9600    End If

        ' ** Shift keys.
9610    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9620      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9630        With Me
9640          intRetVal = 0
9650          .cmdClose.SetFocus
9660        End With
9670      End Select
9680    End If

EXITP:
9690    KeyCode = intRetVal
9700    Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub opgPrint_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_optSpecified_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9820    intShiftDown = (Shift And acShiftMask) > 0
9830    intAltDown = (Shift And acAltMask) > 0
9840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9870        With Me
9880          intRetVal = 0
9890          If .cmdMICRSet.Enabled = True Then
9900            .cmdMICRSet.SetFocus
9910          ElseIf .NextCheckNumber.Enabled = True Then
9920            .NextCheckNumber.SetFocus
9930          ElseIf .cmbSortBy.Enabled = True Then
9940            .cmbSortBy.SetFocus
9950          ElseIf .lbxShortAccountName.Enabled = True Then
9960            .lbxShortAccountName.SetFocus
9970          ElseIf .opgShow.Enabled = True Then
9980            .opgShow.SetFocus
9990          ElseIf .ckgDisplay_opt01_AccountNo.Enabled = True Then
10000           .ckgDisplay_opt01_AccountNo.SetFocus
10010         ElseIf .opgPayeeFont.Enabled = True Then
10020           .opgPayeeFont.SetFocus
10030         ElseIf .cmdPreviewChecks.Enabled = True Then
10040           .cmdPreviewChecks.SetFocus
10050         Else
10060           .cmdClose.SetFocus
10070         End If
10080       End With
10090     End Select
10100   End If

        ' ** Shift keys.
10110   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10130       With Me
10140         intRetVal = 0
10150         .cmdClose.SetFocus
10160       End With
10170     End Select
10180   End If

EXITP:
10190   KeyCode = intRetVal
10200   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub opgPrint_BlankStock_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_BlankStock_box_MouseMove"

10310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10320   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgPrint_optBlank_KeyDown(KeyCode As Integer, Shift As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_optBlank_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10420   intShiftDown = (Shift And acShiftMask) > 0
10430   intAltDown = (Shift And acAltMask) > 0
10440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10470       With Me
10480         intRetVal = 0
10490         If .cmdMICRSet.Enabled = True Then
10500           .cmdMICRSet.SetFocus
10510         ElseIf .NextCheckNumber.Enabled = True Then
10520           .NextCheckNumber.SetFocus
10530         ElseIf .cmbSortBy.Enabled = True Then
10540           .cmbSortBy.SetFocus
10550         ElseIf .lbxShortAccountName.Enabled = True Then
10560           .lbxShortAccountName.SetFocus
10570         ElseIf .opgShow.Enabled = True Then
10580           .opgShow.SetFocus
10590         ElseIf .ckgDisplay_opt01_AccountNo.Enabled = True Then
10600           .ckgDisplay_opt01_AccountNo.SetFocus
10610         ElseIf .opgPayeeFont.Enabled = True Then
10620           .opgPayeeFont.SetFocus
10630         ElseIf .cmdPreviewChecks.Enabled = True Then
10640           .cmdPreviewChecks.SetFocus
10650         Else
10660           .cmdClose.SetFocus
10670         End If
10680       End With
10690     End Select
10700   End If

        ' ** Shift keys.
10710   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10720     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10730       With Me
10740         intRetVal = 0
10750         .cmdClose.SetFocus
10760       End With
10770     End Select
10780   End If

EXITP:
10790   KeyCode = intRetVal
10800   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub opgPrint_optBlank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPrint_optBlank_MouseMove"

10910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10920   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_r_cmd_GotFocus()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_r_cmd_GotFocus"

11010   SeparateCheckingOptSet 1, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
11020   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_r_cmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_r_cmd_MouseDown"

11110   SeparateCheckingOptSet 2, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
11120   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_r_cmd_Click()
' ** Open the box.

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_r_cmd_Click"

11210   SeparateCheckingOptSet 3, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
11220   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_r_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_r_cmd_MouseMove"

11310   SeparateCheckingOptSet 4, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
11320   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_r_cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_r_cmd_MouseUp"

11410   SeparateCheckingOptSet 5, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
11420   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_r_cmd_LostFocus()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_r_cmd_LostFocus"

11510   SeparateCheckingOptSet 6, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
11520   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_l_cmd_GotFocus()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_l_cmd_GotFocus"

11610   SeparateCheckingOptSet 7, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
11620   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_l_cmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_l_cmd_MouseDown"

11710   SeparateCheckingOptSet 8, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
11720   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_l_cmd_Click()
' ** Close the box.

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_l_cmd_Click"

11810   With Me

11820     SeparateCheckingOptSet 9, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
            blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

11830     If lngChkCnt = 0& Then
11840       If .NextCheckNumber.Enabled = True Then
11850         .NextCheckNumber.Enabled = False
11860         .NextCheckNumber.BorderColor = WIN_CLR_DISR
11870         .NextCheckNumber.BackStyle = acBackStyleTransparent
11880         .ChecksTot.Locked = False
11890         .ChecksTot.BorderColor = WIN_CLR_DISR
11900         .cmbSortBy.Enabled = False
11910         .cmbSortBy.BorderColor = WIN_CLR_DISR
11920         .cmbSortBy.BackStyle = acBackStyleTransparent
11930         .chkSyncListSort.Enabled = False
11940       End If
11950     End If

11960   End With

EXITP:
11970   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_l_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_l_cmd_MouseMove"

12010   SeparateCheckingOptSet 10, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
12020   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_l_cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_l_cmd_MouseUp"

12110   SeparateCheckingOptSet 11, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
12120   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_arw_l_cmd_LostFocus()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_arw_l_cmd_LostFocus"

12210   SeparateCheckingOptSet 12, arr_varItem, blnSeparateChecking_R_Focus, blnSeparateChecking_R_MouseDown, _
          blnSeparateChecking_L_Focus, blnSeparateChecking_L_MouseDown, Me  ' ** Module Procedure: modCheckPrintFuncs.

EXITP:
12220   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_lbl1_DblClick(Cancel As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_lbl1_DblClick"

12310   SeparateCheckingOpt_arw_l_cmd_Click  ' ** Procedure: Above.

EXITP:
12320   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_lbl2_DblClick(Cancel As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_lbl2_DblClick"

12410   SeparateCheckingOpt_arw_l_cmd_Click  ' ** Procedure: Above.

EXITP:
12420   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_lbl3_DblClick(Cancel As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_lbl3_DblClick"

12510   SeparateCheckingOpt_arw_l_cmd_Click  ' ** Procedure: Above.

EXITP:
12520   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_lbl4_DblClick(Cancel As Integer)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_lbl4_DblClick"

12610   SeparateCheckingOpt_arw_l_cmd_Click  ' ** Procedure: Above.

EXITP:
12620   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub SeparateCheckingOpt_lbl5_DblClick(Cancel As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "SeparateCheckingOpt_lbl5_DblClick"

12710   SeparateCheckingOpt_arw_l_cmd_Click  ' ** Procedure: Above.

EXITP:
12720   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRSet_GotFocus()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRSet_GotFocus"

12810   With Me
12820     blnMICRSet_Focus = True
12830     .cmdMICRSet_raised_semifocus_dots_img.Visible = True
12840     .cmdMICRSet_raised_img.Visible = False
12850     .cmdMICRSet_raised_focus_img.Visible = False
12860     .cmdMICRSet_raised_focus_dots_img.Visible = False
12870     .cmdMICRSet_sunken_focus_dots_img.Visible = False
12880     .cmdMICRSet_raised_img_dis.Visible = False
12890   End With

EXITP:
12900   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRSet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRSet_MouseDown"

13010   With Me
13020     blnMICRSet_MouseDown = True
13030     .cmdMICRSet_sunken_focus_dots_img.Visible = True
13040     .cmdMICRSet_raised_img.Visible = False
13050     .cmdMICRSet_raised_semifocus_dots_img.Visible = False
13060     .cmdMICRSet_raised_focus_img.Visible = False
13070     .cmdMICRSet_raised_focus_dots_img.Visible = False
13080     .cmdMICRSet_raised_img_dis.Visible = False
13090   End With

EXITP:
13100   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRSet_Click()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRSet_Click"

13210   DoCmd.OpenForm "frmRpt_Checks_MICR_Set", , , , , , THIS_NAME

EXITP:
13220   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRSet_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRSet_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
13370       With Me
13380         intRetVal = 0
13390         .cmdMICRAdjust.SetFocus
13400       End With
13410     End Select
13420   End If

        ' ** Shift keys.
13430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13450       With Me
13460         intRetVal = 0
13470         .opgPrint.SetFocus
13480       End With
13490     End Select
13500   End If

EXITP:
13510   KeyCode = intRetVal
13520   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmdMICRSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRSet_MouseMove"

13610   With Me
13620     If blnMICRSet_MouseDown = False Then
13630       Select Case blnMICRSet_Focus
            Case True
13640         .cmdMICRSet_raised_focus_dots_img.Visible = True
13650         .cmdMICRSet_raised_focus_img.Visible = False
13660       Case False
13670         .cmdMICRSet_raised_focus_img.Visible = True
13680         .cmdMICRSet_raised_focus_dots_img.Visible = False
13690       End Select
13700       .cmdMICRSet_raised_img.Visible = False
13710       .cmdMICRSet_raised_semifocus_dots_img.Visible = False
13720       .cmdMICRSet_sunken_focus_dots_img.Visible = False
13730       .cmdMICRSet_raised_img_dis.Visible = False
13740     End If
13750     If .cmdMICRAdjust_raised_focus_dots_img.Visible = True Or .cmdMICRAdjust_raised_focus_img.Visible = True Then
13760       Select Case blnMICRAdjust_Focus
            Case True
13770         .cmdMICRAdjust_raised_semifocus_dots_img.Visible = True
13780         .cmdMICRAdjust_raised_img.Visible = False
13790       Case False
13800         .cmdMICRAdjust_raised_img.Visible = True
13810         .cmdMICRAdjust_raised_semifocus_dots_img.Visible = False
13820       End Select
13830       .cmdMICRAdjust_raised_focus_img.Visible = False
13840       .cmdMICRAdjust_raised_focus_dots_img.Visible = False
13850       .cmdMICRAdjust_sunken_focus_dots_img.Visible = False
13860       .cmdMICRAdjust_raised_img_dis.Visible = False
13870     End If
13880   End With

EXITP:
13890   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRSet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRSet_MouseUp"

13910   With Me
13920     .cmdMICRSet_raised_focus_dots_img.Visible = True
13930     .cmdMICRSet_raised_img.Visible = False
13940     .cmdMICRSet_raised_semifocus_dots_img.Visible = False
13950     .cmdMICRSet_raised_focus_img.Visible = False
13960     .cmdMICRSet_sunken_focus_dots_img.Visible = False
13970     .cmdMICRSet_raised_img_dis.Visible = False
13980     blnMICRSet_MouseDown = False
13990   End With

EXITP:
14000   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRSet_LostFocus()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRSet_LostFocus"

14110   With Me
14120     .cmdMICRSet_raised_img.Visible = True
14130     .cmdMICRSet_raised_semifocus_dots_img.Visible = False
14140     .cmdMICRSet_raised_focus_img.Visible = False
14150     .cmdMICRSet_raised_focus_dots_img.Visible = False
14160     .cmdMICRSet_sunken_focus_dots_img.Visible = False
14170     .cmdMICRSet_raised_img_dis.Visible = False
14180     blnMICRSet_Focus = False
14190   End With

EXITP:
14200   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRAdjust_GotFocus()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRAdjust_GotFocus"

14310   With Me
14320     blnMICRAdjust_Focus = True
14330     .cmdMICRAdjust_raised_semifocus_dots_img.Visible = True
14340     .cmdMICRAdjust_raised_img.Visible = False
14350     .cmdMICRAdjust_raised_focus_img.Visible = False
14360     .cmdMICRAdjust_raised_focus_dots_img.Visible = False
14370     .cmdMICRAdjust_sunken_focus_dots_img.Visible = False
14380     .cmdMICRAdjust_raised_img_dis.Visible = False
14390   End With

EXITP:
14400   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRAdjust_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRAdjust_MouseDown"

14510   With Me
14520     blnMICRAdjust_MouseDown = True
14530     .cmdMICRAdjust_sunken_focus_dots_img.Visible = True
14540     .cmdMICRAdjust_raised_img.Visible = False
14550     .cmdMICRAdjust_raised_semifocus_dots_img.Visible = False
14560     .cmdMICRAdjust_raised_focus_img.Visible = False
14570     .cmdMICRAdjust_raised_focus_dots_img.Visible = False
14580     .cmdMICRAdjust_raised_img_dis.Visible = False
14590   End With

EXITP:
14600   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRAdjust_Click()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRAdjust_Click"

        Dim strDocName As String

14710   With Me

14720     DoCmd.Hourglass True
14730     DoEvents

14740     gblnSetFocus = True
14750     strDocName = "frmRpt_Checks_MICR_Adjust"
14760     DoCmd.OpenForm strDocName, , , , , , THIS_NAME
14770     If gblnGoToReport = True Then
14780       .TimerInterval = 0&
14790       Forms(strDocName).TimerInterval = 100&
14800       blnGoingToReport = False
14810       blnGTR_Emblem = False
14820       .GoToReport_arw_micr_img.Visible = False
14830       .cmdMICRSet.Visible = True
14840       .cmdMICRSet_raised_img.Visible = True
14850       GTREmblem_Off  ' ** Procedure: Below.
14860     End If

14870   End With

EXITP:
14880   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRAdjust_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRAdjust_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14920   intShiftDown = (Shift And acShiftMask) > 0
14930   intAltDown = (Shift And acAltMask) > 0
14940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14970       With Me
14980         intRetVal = 0
14990         If .NextCheckNumber.Enabled = True Then
15000           .NextCheckNumber.SetFocus
15010         ElseIf .cmbSortBy.Enabled = True Then
15020           .cmbSortBy.SetFocus
15030         ElseIf .lbxShortAccountName.Enabled = True Then
15040           .lbxShortAccountName.SetFocus
15050         ElseIf .opgShow.Enabled = True Then
15060           .opgShow.SetFocus
15070         ElseIf .ckgDisplay_opt01_AccountNo.Enabled = True Then
15080           .ckgDisplay_opt01_AccountNo.SetFocus
15090         ElseIf .opgPayeeFont.Enabled = True Then
15100           .opgPayeeFont.SetFocus
15110         ElseIf .cmdPreviewChecks.Enabled = True Then
15120           .cmdPreviewChecks.SetFocus
15130         Else
15140           .cmdClose.SetFocus
15150         End If
15160       End With
15170     End Select
15180   End If

        ' ** Shift keys.
15190   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15200     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15210       With Me
15220         intRetVal = 0
15230         .cmdMICRSet.SetFocus
15240       End With
15250     End Select
15260   End If

EXITP:
15270   KeyCode = intRetVal
15280   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmdMICRAdjust_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRAdjust_MouseMove"

15310   With Me
15320     If blnMICRAdjust_MouseDown = False Then
15330       Select Case blnMICRAdjust_Focus
            Case True
15340         .cmdMICRAdjust_raised_focus_dots_img.Visible = True
15350         .cmdMICRAdjust_raised_focus_img.Visible = False
15360       Case False
15370         .cmdMICRAdjust_raised_focus_img.Visible = True
15380         .cmdMICRAdjust_raised_focus_dots_img.Visible = False
15390       End Select
15400       .cmdMICRAdjust_raised_img.Visible = False
15410       .cmdMICRAdjust_raised_semifocus_dots_img.Visible = False
15420       .cmdMICRAdjust_sunken_focus_dots_img.Visible = False
15430       .cmdMICRAdjust_raised_img_dis.Visible = False
15440     End If
15450     If .cmdMICRSet_raised_focus_dots_img.Visible = True Or .cmdMICRSet_raised_focus_img.Visible = True Then
15460       Select Case blnMICRSet_Focus
            Case True
15470         .cmdMICRSet_raised_semifocus_dots_img.Visible = True
15480         .cmdMICRSet_raised_img.Visible = False
15490       Case False
15500         .cmdMICRSet_raised_img.Visible = True
15510         .cmdMICRSet_raised_semifocus_dots_img.Visible = False
15520       End Select
15530       .cmdMICRSet_raised_focus_img.Visible = False
15540       .cmdMICRSet_raised_focus_dots_img.Visible = False
15550       .cmdMICRSet_sunken_focus_dots_img.Visible = False
15560       .cmdMICRSet_raised_img_dis.Visible = False
15570     End If
15580   End With

EXITP:
15590   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRAdjust_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRAdjust_MouseUp"

15610   With Me
15620     .cmdMICRAdjust_raised_focus_dots_img.Visible = True
15630     .cmdMICRAdjust_raised_img.Visible = False
15640     .cmdMICRAdjust_raised_semifocus_dots_img.Visible = False
15650     .cmdMICRAdjust_raised_focus_img.Visible = False
15660     .cmdMICRAdjust_sunken_focus_dots_img.Visible = False
15670     .cmdMICRAdjust_raised_img_dis.Visible = False
15680     blnMICRAdjust_MouseDown = False
15690   End With

EXITP:
15700   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdMICRAdjust_LostFocus()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMICRAdjust_LostFocus"

15810   With Me
15820     .cmdMICRAdjust_raised_img.Visible = True
15830     .cmdMICRAdjust_raised_semifocus_dots_img.Visible = False
15840     .cmdMICRAdjust_raised_focus_img.Visible = False
15850     .cmdMICRAdjust_raised_focus_dots_img.Visible = False
15860     .cmdMICRAdjust_sunken_focus_dots_img.Visible = False
15870     .cmdMICRAdjust_raised_img_dis.Visible = False
15880     blnMICRAdjust_Focus = False
15890   End With

EXITP:
15900   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub NextCheckNumber_KeyDown(KeyCode As Integer, Shift As Integer)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "NextCheckNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16020   intShiftDown = (Shift And acShiftMask) > 0
16030   intAltDown = (Shift And acAltMask) > 0
16040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16070       With Me
16080         intRetVal = 0
16090         If .cmbSortBy.Enabled = True Then
16100           .cmbSortBy.SetFocus
16110         ElseIf .lbxShortAccountName.Enabled = True Then
16120           .lbxShortAccountName.SetFocus
16130         ElseIf .opgShow.Enabled = True Then
16140           .opgShow.SetFocus
16150         ElseIf .ckgDisplay_opt01_AccountNo.Enabled = True Then
16160           .ckgDisplay_opt01_AccountNo.SetFocus
16170         ElseIf .opgPayeeFont.Enabled = True Then
16180           .opgPayeeFont.SetFocus
16190         ElseIf .cmdPreviewChecks.Enabled = True Then
16200           .cmdPreviewChecks.SetFocus
16210         Else
16220           .cmdClose.SetFocus
16230         End If
16240       End With
16250     End Select
16260   End If

        ' ** Shift keys.
16270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16290       With Me
16300         intRetVal = 0
16310         If .cmdMICRAdjust.Enabled = True Then
16320           .cmdMICRAdjust.SetFocus
16330         Else
16340           .opgPrint.SetFocus
16350         End If
16360       End With
16370     End Select
16380   End If

EXITP:
16390   KeyCode = intRetVal
16400   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmbSortBy_AfterUpdate()
' ** Starts with:
' **   qryPrintChecks_01_01a_Account_Number

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbSortBy_AfterUpdate"

        Dim strThatProc As String, strCurSort As String
        Dim strTmp01 As String
        Dim intPos01 As Integer

16510   With Me
          ' ** Can't sort checks by:
          ' **   Check_Count
          ' **   Last_Check_Number
16520     If .chkSyncListSort = True Then
            ' ** If the column isn't currently showing, show it.
16530       If .Controls(.cmbSortBy.Column(CBX_SORT_CNAM)) = False Then
16540         .Controls(.cmbSortBy.Column(CBX_SORT_CNAM)) = True
16550         Eval ("Fnc_" & .cmbSortBy.Column(CBX_SORT_CNAM) & "_AfterUpdate()")  ' ** Module Functions: modReportFunctions.
16560         DoEvents
16570       End If
16580       strCurSort = .lbxShortAccountName.RowSource  ' ** qryPrintChecks_01_07d_Bank_Account_Number.
16590       intPos01 = CharPos(strCurSort, 3, "_")  ' ** Module Function: modStringFuncs.
16600       strTmp01 = Mid(strCurSort, (intPos01 + 1))
16610       If .cmbSortBy.Column(CBX_SORT_FNAM) <> strTmp01 Then
              ' ** Only change the sort if it's not already there.
16620         strThatProc = (.cmbSortBy.Column(CBX_SORT_FNAM) & "_lbl_DblClick")
16630         SortNow strThatProc  ' ** Procedure: Below.
16640       End If
16650       If .cmbSortBy.Column(CBX_SORT_FNAM) = "Check_Count" Or .cmbSortBy.Column(CBX_SORT_FNAM) = "Last_Check_Number" Then
16660         Beep
16670         .cmbSortBy_lbl2.Visible = True
16680         .cmdPrintChecks.Enabled = False
16690         .cmdPreviewChecks.Enabled = False
16700         .cmbSortBy.ForeColor = CLR_DISABLED_FG
16710         .cmbSortBy.BackColor = CLR_DISABLED_BG
16720       ElseIf .cmdPrintChecks.Enabled = False Then
16730         .cmbSortBy_lbl2.Visible = False
16740         .cmdPrintChecks.Enabled = True
16750         .cmdPreviewChecks.Enabled = True
16760         .cmbSortBy.ForeColor = CLR_VDKGRY
16770         .cmbSortBy.BackColor = CLR_WHT
16780       End If
16790     End If
16800   End With

EXITP:
16810   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmbSortBy_KeyDown(KeyCode As Integer, Shift As Integer)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbSortBy_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16920   intShiftDown = (Shift And acShiftMask) > 0
16930   intAltDown = (Shift And acAltMask) > 0
16940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16970       With Me
16980         intRetVal = 0
16990         .chkSyncListSort.SetFocus
17000       End With
17010     End Select
17020   End If

        ' ** Shift keys.
17030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17050       With Me
17060         intRetVal = 0
17070         If .NextCheckNumber.Enabled = True Then
17080           .NextCheckNumber.SetFocus
17090         ElseIf .cmdMICRAdjust.Enabled = True Then
17100           .cmdMICRAdjust.SetFocus
17110         Else
17120           .opgPrint.SetFocus
17130         End If
17140       End With
17150     End Select
17160   End If

EXITP:
17170   KeyCode = intRetVal
17180   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub chkSyncListSort_AfterUpdate()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSyncListSort_AfterUpdate"

        Dim strThatProc As String, strCurSort As String
        Dim strTmp01 As String
        Dim intPos01 As Integer

17210   With Me
17220     Select Case .chkSyncListSort
          Case True
17230       .chkSyncListSort_lbl.FontBold = True
17240       .cmbSortBy.RowSource = "qryPrintChecks_23x"
17250       If IsNull(.cmbSortBy) = False Then
17260         If Trim(.cmbSortBy) <> vbNullString Then
                ' ** If the column isn't currently showing, show it.
17270           If .Controls(.cmbSortBy.Column(CBX_SORT_CNAM)) = False Then
17280             .Controls(.cmbSortBy.Column(CBX_SORT_CNAM)) = True
17290             Eval ("Fnc_" & .cmbSortBy.Column(CBX_SORT_CNAM) & "_AfterUpdate()")  ' ** Module Functions: modReportFunctions.
17300             DoEvents
17310           End If
17320           strCurSort = .lbxShortAccountName.RowSource  ' ** qryPrintChecks_01_07d_Bank_Account_Number.
17330           intPos01 = CharPos(strCurSort, 3, "_")  ' ** Module Function: modStringFuncs.
17340           strTmp01 = Mid(strCurSort, (intPos01 + 1))
17350           If .cmbSortBy.Column(CBX_SORT_FNAM) <> strTmp01 Then
                  ' ** Only change the sort if it's not already there.
17360             strThatProc = (.cmbSortBy.Column(CBX_SORT_FNAM) & "_lbl_DblClick")
17370             SortNow strThatProc  ' ** Procedure: Below.
17380           End If
17390         End If
17400       End If
17410     Case False
17420       .chkSyncListSort_lbl.FontBold = False
17430       .cmbSortBy.RowSource = "qryPrintChecks_24x"
17440     End Select
17450   End With

EXITP:
17460   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub chkSyncListSort_KeyDown(KeyCode As Integer, Shift As Integer)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSyncListSort_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17520   intShiftDown = (Shift And acShiftMask) > 0
17530   intAltDown = (Shift And acAltMask) > 0
17540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17570       With Me
17580         intRetVal = 0
17590         If .lbxShortAccountName.Enabled = True Then
17600           .lbxShortAccountName.SetFocus
17610         ElseIf .opgShow.Enabled = True Then
17620           .opgShow.SetFocus
17630         ElseIf .ckgDisplay_opt01_AccountNo.Enabled = True Then
17640           .ckgDisplay_opt01_AccountNo.SetFocus
17650         ElseIf .opgPayeeFont.Enabled = True Then
17660           .opgPayeeFont.SetFocus
17670         ElseIf .cmdPreviewChecks.Enabled = True Then
17680           .cmdPreviewChecks.SetFocus
17690         Else
17700           .cmdClose.SetFocus
17710         End If
17720       End With
17730     End Select
17740   End If

        ' ** Shift keys.
17750   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17770       With Me
17780         intRetVal = 0
17790         .cmbSortBy.SetFocus
17800       End With
17810     End Select
17820   End If

EXITP:
17830   KeyCode = intRetVal
17840   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub lbxShortAccountName_AfterUpdate()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "lbxShortAccountName_AfterUpdate"

        Dim blnUseDefault As Boolean
        Dim varTmp00 As Variant

17910   blnUseDefault = True

17920   With Me
17930     If .lbxShortAccountName.Column(LBX_CHK_STAT) = "Yes" Then
17940       If blnIsOpen = False Then
17950         MsgBox "All checks for this account have a check number listed," & vbCrLf & "indicating they have already been printed." & vbCrLf & vbCrLf & _
                "Checks cannot be reprinted unless marked as void.", vbInformation + vbOKOnly, "Checks Already Printed"
17960       End If
17970       .lbxShortAccountName = 0
            'SO, HOW COULD THEY VOID THESE CHECKS AT THIS POINT?
            'OUR NEW FEATURE ONLY ALLOWS THAT AT THE TIME OF PRINTING!
17980     Else
17990       If .lbxShortAccountName.Column(LBX_CHK_STAT) = "Mixed" Then
18000         If blnIsOpen = False Then
18010           MsgBox "Some of the checks for this account have a check number listed," & vbCrLf & "indicating they have already been printed." & vbCrLf & vbCrLf & _
                  "Only those checks without a check number will be printed.", vbInformation + vbOKOnly, "Some Checks Already Printed"
18020         End If
18030       End If

18040       If lngChkCnt > 0& Then
18050         Select Case .opgPrint
              Case .opgPrint_optAll.OptionValue
                ' ** Just Ledger:
                ' ** qryPrintChecks_31a (qryPrintChecks_30d (Union of qryPrintChecks_30a
                ' ** (Ledger, CheckNum <> Null), qryPrintChecks_30b (LedgerArchive, just
                ' ** CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno, with Max(CheckNum), cnt;
                ' ** by AccountNo), grouped, with Max(CheckNum), cnt, for All.
                'varTmp00 = DLookup("[LastCheckNum]", "qryPrintChecks_31b")
                ' **
                ' ** Both Ledger and Journal:
                ' ** qryPrintChecks_33a (qryPrintChecks_32c (Union of qryPrintChecks_31a
                ' ** (qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
                ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by
                ' ** accountno, with Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b
                ' ** (qryPrintChecks_32a (Journal, just CheckNum <> Null), grouped by accountno,
                ' ** with Max(CheckNum), cnt)), grouped by accountno, with Max(CheckNum), cnt;
                ' ** for AccountNo), grouped, with Max(CheckNum), cnt; for All.
18060           varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33b")
18070         Case Else
                ' ** Just Ledger:
                ' ** qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
                ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno,
                ' ** with Max(CheckNum), cnt; by AccountNo.
                'varTmp00 = DLookup("[LastCheckNum]", "qryPrintChecks_31a", "[accountno] = '" & .lbxShortAccountName.Column(LBX_CHK_ACTNO) & "'")
                ' **
                ' ** Both Ledger and Journal:
                ' ** qryPrintChecks_32c (Union of qryPrintChecks_31a (qryPrintChecks_30d
                ' ** (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null), qryPrintChecks_30b
                ' ** (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno, with
                ' ** Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b (qryPrintChecks_32a
                ' ** (Journal, just CheckNum <> Null), grouped by accountno, with Max(CheckNum),
                ' ** cnt)), grouped by accountno, with Max(CheckNum), cnt; for AccountNo.
18080           varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33a", "[accountno] = '" & .lbxShortAccountName.Column(LBX_CHK_ACTNO) & "'")
18090         End Select
18100         If IsNull(varTmp00) = False Then
18110           If varTmp00 > 0 Then
18120             .NextCheckNumber = CLng(varTmp00 + 1)
18130             blnUseDefault = False
18140           Else
18150             Select Case blnPrefCheckNum
                  Case True
18160               If IsNull(.NextCheckNumber) = False Then
18170                 If Val(.NextCheckNumber) > 0 Then
                        ' ** Let it stand.
18180                   blnUseDefault = False
18190                 End If
18200               End If
18210             Case False
                    ' ** Use the default.
18220             End Select
18230           End If
18240         End If
18250       Else
18260         Select Case blnPrefCheckNum
              Case True
18270           If IsNull(.NextCheckNumber) = False Then
18280             If Val(.NextCheckNumber) > 0 Then
                    ' ** Let it stand.
18290               blnUseDefault = False
18300             End If
18310           End If
18320         Case False
18330           .NextCheckNumber = Null
18340           blnUseDefault = False
18350         End Select
18360       End If
18370       If blnUseDefault = True Then
18380         .NextCheckNumber = 1000&
18390       End If
18400     End If  ' ** ChkStat.
18410   End With

EXITP:
18420   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub lbxShortAccountName_KeyDown(KeyCode As Integer, Shift As Integer)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "lbxShortAccountName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18520   intShiftDown = (Shift And acShiftMask) > 0
18530   intAltDown = (Shift And acAltMask) > 0
18540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18570       With Me
18580         intRetVal = 0
18590         If .opgShow.Enabled = True Then
18600           .opgShow.SetFocus
18610         ElseIf .ckgDisplay_opt01_AccountNo.Enabled = True Then
18620           .ckgDisplay_opt01_AccountNo.SetFocus
18630         ElseIf .opgPayeeFont.Enabled = True Then
18640           .opgPayeeFont.SetFocus
18650         ElseIf .cmdPreviewChecks.Enabled = True Then
18660           .cmdPreviewChecks.SetFocus
18670         Else
18680           .cmdClose.SetFocus
18690         End If
18700       End With
18710     End Select
18720   End If

        ' ** Shift keys.
18730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18750       With Me
18760         intRetVal = 0
18770         If .chkSyncListSort.Enabled = True Then
18780           .chkSyncListSort.SetFocus
18790         ElseIf .cmbSortBy.Enabled = True Then
18800           .cmbSortBy.SetFocus
18810         ElseIf .NextCheckNumber.Enabled = True Then
18820           .NextCheckNumber.SetFocus
18830         ElseIf .cmdMICRAdjust.Enabled = True Then
18840           .cmdMICRAdjust.SetFocus
18850         ElseIf .opgPrint.Enabled = True Then
18860           .opgPrint.SetFocus
18870         Else
18880           .cmdClose.SetFocus
18890         End If
18900       End With
18910     End Select
18920   End If

EXITP:
18930   KeyCode = intRetVal
18940   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub lbxShortAccountName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "lbxShortAccountName_MouseMove"

19010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19020   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Account_Number_lbl_DblClick(Cancel As Integer)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "Account_Number_lbl_DblClick"

19110   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19120   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Short_Name_lbl_DblClick(Cancel As Integer)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "Short_Name_lbl_DblClick"

19210   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19220   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Check_Count_lbl_DblClick(Cancel As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "Check_Count_lbl_DblClick"

19310   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19320   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Last_Check_Number_lbl_DblClick(Cancel As Integer)

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "Last_Check_Number_lbl_DblClick"

19410   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19420   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Payee_lbl_DblClick(Cancel As Integer)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "Payee_lbl_DblClick"

19510   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19520   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Bank_Name_lbl_DblClick(Cancel As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Name_lbl_DblClick"

19610   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19620   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Bank_Account_Number_lbl_DblClick(Cancel As Integer)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Account_Number_lbl_DblClick"

19710   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19720   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub Check_Amount_lbl_DblClick(Cancel As Integer)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "Check_Amount_lbl_DblClick"

19810   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19820   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub ChkStat_lbl_DblClick(Cancel As Integer)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkStat_lbl_DblClick"

19910   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19920   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgShow_AfterUpdate()
' ** strSortNow is a query name.
' ** Sorting is accomplished by a pre-written group of queries
' ** that cover every combination of fields and sorts, both
' ** ascending and descending.
' ** 07/02/2017: I just discovered that when it's on individual checks,
' **             you can preview or print the checks one-at-a-time!
' **             I didn't plan that! And I don't think it's a good idea.

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_AfterUpdate"

        Dim strCap As String
        Dim blnDesc As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long

20010   With Me

20020     intPos01 = InStr((InStr((InStr(strSortNow, "_") + 1), strSortNow, "_") + 1), strSortNow, "_")

20030     strTmp01 = Right(Left(strSortNow, intPos01), 2)
20040     Select Case strTmp01
          Case "b_", "d_"
20050       blnDesc = True
20060     Case "a_", "c_"
20070       blnDesc = False
20080     End Select

20090     strTmp01 = Left(strSortNow, (intPos01 - 2))
20100     strTmp02 = Mid(strSortNow, (intPos01 + 1))

20110     Select Case .opgShow
          Case .opgShow_optAccount.OptionValue
20120       .opgShow_optAccount_lbl.FontBold = True
20130       .opgShow_optAccount_lbl_dim.FontBold = True
20140       .opgShow_optAccount_lbl_dim_hi.FontBold = True
20150       .opgShow_optAll_lbl.FontBold = False
20160       .opgShow_optAll_lbl_dim.FontBold = False
20170       .opgShow_optAll_lbl_dim_hi.FontBold = False
20180       .Check_Amount_lbl.Caption = "Check Total"
20190       .Check_Amount_lbl_dim_hi.Caption = "Check Total"
20200       strCap = "Se&lect Account to Print"
20210       .lbxShortAccountName_lbl.Caption = strCap
20220       .lbxShortAccountName_lbl_dim.Caption = strCap
20230       .lbxShortAccountName_lbl_dim_hi.Caption = strCap
20240       .lbxShortAccountName_lbl.Width = lngLbxLbl_Width
20250       .lbxShortAccountName_lbl_dim.Width = lngLbxLbl_Width
20260       .lbxShortAccountName_lbl_dim_hi.Width = lngLbxLbl_Width
20270       .lbxShortAccountName_box2.Width = (lngLbxLbl_Width + lngTpp)
20280       .lbxShortAccountName_hline01.Width = lngLbxHLine01_Width
20290       .lbxShortAccountName_hline02.Width = (lngLbxHLine01_Width - (2& * lngTpp))
20300       .lbxShortAccountName_hline03.Width = (lngLbxHLine01_Width - (2& * lngTpp))
20310       .lbxShortAccountName_vline03.Left = lngLbxVLine03_Left
20320       .lbxShortAccountName_vline04.Left = (lngLbxVLine03_Left + lngTpp)
20330       Select Case blnDesc
            Case True
20340         strTmp01 = strTmp01 & "b_" & strTmp02
20350       Case False
20360         strTmp01 = strTmp01 & "a_" & strTmp02
20370       End Select
20380       .lbxShortAccountName.RowSource = strTmp01
20390     Case .opgShow_optAll.OptionValue
20400       .opgShow_optAccount_lbl.FontBold = False
20410       .opgShow_optAccount_lbl_dim.FontBold = False
20420       .opgShow_optAccount_lbl_dim_hi.FontBold = False
20430       .opgShow_optAll_lbl.FontBold = True
20440       .opgShow_optAll_lbl_dim.FontBold = True
20450       .opgShow_optAll_lbl_dim_hi.FontBold = True
20460       .Check_Amount_lbl.Caption = "Check Amount"
20470       .Check_Amount_lbl_dim_hi.Caption = "Check Amount"
20480       strCap = "&List of Individual Checks"
20490       .lbxShortAccountName_lbl.Caption = strCap
20500       .lbxShortAccountName_lbl_dim.Caption = strCap
20510       .lbxShortAccountName_lbl_dim_hi.Caption = strCap
20520       lngTmp03 = (6& * lngTpp)
20530       .lbxShortAccountName_lbl.Width = (lngLbxLbl_Width + lngTmp03)
20540       .lbxShortAccountName_lbl_dim.Width = (lngLbxLbl_Width + lngTmp03)
20550       .lbxShortAccountName_lbl_dim_hi.Width = (lngLbxLbl_Width + lngTmp03)
20560       .lbxShortAccountName_box2.Width = ((lngLbxLbl_Width + lngTpp) + lngTmp03)
20570       .lbxShortAccountName_hline01.Width = (lngLbxHLine01_Width + lngTmp03)
20580       .lbxShortAccountName_hline02.Width = ((lngLbxHLine01_Width - (2& * lngTpp)) + lngTmp03)
20590       .lbxShortAccountName_hline03.Width = ((lngLbxHLine01_Width - (2& * lngTpp)) + lngTmp03)
20600       .lbxShortAccountName_vline03.Left = (lngLbxVLine03_Left + lngTmp03)
20610       .lbxShortAccountName_vline04.Left = ((lngLbxVLine03_Left + lngTpp) + lngTmp03)
20620       Select Case blnDesc
            Case True
20630         strTmp01 = strTmp01 & "d_" & strTmp02
20640       Case False
20650         strTmp01 = strTmp01 & "c_" & strTmp02
20660       End Select
20670       .lbxShortAccountName.RowSource = strTmp01
20680     End Select

20690     .lbxShortAccountName.SetFocus

          'THIS COVERS WHAT TYPE OF ROWS ARE SHOWN, BUT DOES
          'NOT DEAL AT ALL WITH WHICH COLUMNS ARE SHOWN.

20700   End With

EXITP:
20710   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgShow_optAccount_KeyDown(KeyCode As Integer, Shift As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_optAccount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20820   intShiftDown = (Shift And acShiftMask) > 0
20830   intAltDown = (Shift And acAltMask) > 0
20840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20870       With Me
20880         intRetVal = 0
20890         If .ckgDisplay_opt01_AccountNo.Enabled = True Then
20900           .ckgDisplay_opt01_AccountNo.SetFocus
20910         ElseIf .opgPayeeFont.Enabled = True Then
20920           .opgPayeeFont.SetFocus
20930         ElseIf .cmdPreviewChecks.Enabled = True Then
20940           .cmdPreviewChecks.SetFocus
20950         Else
20960           .cmdClose.SetFocus
20970         End If
20980       End With
20990     End Select
21000   End If

        ' ** Shift keys.
21010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21020     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21030       With Me
21040         intRetVal = 0
21050         If .lbxShortAccountName.Enabled = True Then
21060           .lbxShortAccountName.SetFocus
21070         ElseIf .chkSyncListSort.Enabled = True Then
21080           .chkSyncListSort.SetFocus
21090         ElseIf .cmbSortBy.Enabled = True Then
21100           .cmbSortBy.SetFocus
21110         ElseIf .NextCheckNumber.Enabled = True Then
21120           .NextCheckNumber.SetFocus
21130         ElseIf .cmdMICRAdjust.Enabled = True Then
21140           .cmdMICRAdjust.SetFocus
21150         Else
21160           .opgPrint.SetFocus
21170         End If
21180       End With
21190     End Select
21200   End If

EXITP:
21210   KeyCode = intRetVal
21220   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub opgShow_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_optAll_KeyDown"

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
21390         If .ckgDisplay_opt01_AccountNo.Enabled = True Then
21400           .ckgDisplay_opt01_AccountNo.SetFocus
21410         ElseIf .opgPayeeFont.Enabled = True Then
21420           .opgPayeeFont.SetFocus
21430         ElseIf .cmdPreviewChecks.Enabled = True Then
21440           .cmdPreviewChecks.SetFocus
21450         Else
21460           .cmdClose.SetFocus
21470         End If
21480       End With
21490     End Select
21500   End If

        ' ** Shift keys.
21510   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21530       With Me
21540         intRetVal = 0
21550         If .lbxShortAccountName.Enabled = True Then
21560           .lbxShortAccountName.SetFocus
21570         ElseIf .chkSyncListSort.Enabled = True Then
21580           .chkSyncListSort.SetFocus
21590         ElseIf .cmbSortBy.Enabled = True Then
21600           .cmbSortBy.SetFocus
21610         ElseIf .NextCheckNumber.Enabled = True Then
21620           .NextCheckNumber.SetFocus
21630         ElseIf .cmdMICRAdjust.Enabled = True Then
21640           .cmdMICRAdjust.SetFocus
21650         Else
21660           .opgPrint.SetFocus
21670         End If
21680       End With
21690     End Select
21700   End If

EXITP:
21710   KeyCode = intRetVal
21720   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub ckgDisplay_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_box_MouseMove"

21810   Detail_MouseMove Button, Shift, X, Y

EXITP:
21820   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Public Sub ckgDisplay_opt01_AccountNo_AfterUpdate()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt01_AccountNo_AfterUpdate"

        Dim Cancel As Integer

21910   With Me
21920     Cancel = 0

21930     Select Case .ckgDisplay_opt01_AccountNo
          Case True
21940       .ckgDisplay_opt01_AccountNo_lbl.FontBold = True
21950       .ckgDisplay_opt01_AccountNo_lbl_dim.FontBold = True
21960       .ckgDisplay_opt01_AccountNo_lbl_dim_hi.FontBold = True
21970     Case False
21980       .ckgDisplay_opt01_AccountNo_lbl.FontBold = False
21990       .ckgDisplay_opt01_AccountNo_lbl_dim.FontBold = False
22000       .ckgDisplay_opt01_AccountNo_lbl_dim_hi.FontBold = False
22010     End Select

22020     ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
            lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
            lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
            lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me, .ckgDisplay_opt01_AccountNo  ' ** Module Procedure: modCheckPrintFuncs.
          'ColumnSet THIS_PROC, .ckgDisplay_opt01_AccountNo  ' ** Procedure: Below
22030     blnSortAgain = True  ' ** So it won't change the sort direction.

22040     Select Case .ckgDisplay_opt01_AccountNo
          Case True
22050       If InStr(strSortNow, "Account_Number") > 0 Then
              ' ** If this is now on, and the sort already included 'Account_Number', something's screwy!
22060         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22070       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
              ' ** If this is now on, and the sort was on 'Short_Name', check whether 'Short_Name' is still on.
22080         Select Case .ckgDisplay_opt02_ShortName
              Case True
22090           Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
22100         Case False
22110           Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22120         End Select
22130       End If
22140     Case False
22150       If InStr(strSortNow, "Account_Number") > 0 Then
              ' ** If this is now off, and the sort was on it, move to 'Short_Name'.
22160         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
22170       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
              ' ** If this is now off, but the sort was on 'Short_Name', leave it there.
22180         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
22190       End If
22200     End Select

22210     If InStr(strSortNow, "ChkStat") > 0 Then
22220       ChkStat_lbl_DblClick Cancel ' ** Procedure: Above.
22230     ElseIf InStr(strSortNow, "Check_Count") > 0 Then
22240       Check_Count_lbl_DblClick Cancel ' ** Procedure: Above.
22250     ElseIf InStr(strSortNow, "Last_Check_Number") > 0 Then
22260       Last_Check_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22270     ElseIf InStr(strSortNow, "Payee") > 0 Then
22280       Payee_lbl_DblClick Cancel ' ** Procedure: Above.
22290     ElseIf InStr(strSortNow, "Bank_Name") > 0 Then
22300       Bank_Name_lbl_DblClick Cancel ' ** Procedure: Above.
22310     ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
22320       Bank_Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22330     ElseIf InStr(strSortNow, "Check_Amount") > 0 Then
22340       Check_Amount_lbl_DblClick Cancel ' ** Procedure: Above.
22350     End If

          'Account_Number
          'Short_Name
          'Check_Count
          'Last_Check_Number
          'Payee
          'Bank_Name
          'Bank_Account_Number
          'Check_Amount
          'ChkStat
22360   End With

EXITP:
22370   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Public Sub ckgDisplay_opt02_ShortName_AfterUpdate()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt02_ShortName_AfterUpdate"

        Dim Cancel As Integer

22410   With Me
22420     Cancel = 0

22430     Select Case .ckgDisplay_opt02_ShortName
          Case True
22440       .ckgDisplay_opt02_ShortName_lbl.FontBold = True
22450       .ckgDisplay_opt02_ShortName_lbl_dim.FontBold = True
22460       .ckgDisplay_opt02_ShortName_lbl_dim_hi.FontBold = True
22470     Case False
22480       .ckgDisplay_opt02_ShortName_lbl.FontBold = False
22490       .ckgDisplay_opt02_ShortName_lbl_dim.FontBold = False
22500       .ckgDisplay_opt02_ShortName_lbl_dim_hi.FontBold = False
22510     End Select

22520     ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
            lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
            lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
            lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me, .ckgDisplay_opt02_ShortName  ' ** Module Procedure: modCheckPrintFuncs.
          'ColumnSet THIS_PROC, .ckgDisplay_opt02_ShortName  ' ** Procedure: Below
22530     blnSortAgain = True  ' ** So it won't change the sort direction.

22540     Select Case .ckgDisplay_opt02_ShortName
          Case True
22550       If InStr(strSortNow, "Short_Name") > 0 Then
              ' ** If this is now on, and the sort already included 'Short_Name', something's screwy!
22560         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
22570       ElseIf InStr(strSortNow, "Account_Number") > 0 Then
              ' ** If this is now on, and the sort was on 'Account_Number', check whether 'Account_Number' is still on.
22580         Select Case .ckgDisplay_opt01_AccountNo
              Case True
22590           Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22600         Case False
22610           Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
22620         End Select
22630       End If
22640     Case False
22650       If InStr(strSortNow, "Short_Name") > 0 Then
              ' ** If this is now off, and the sort was on it, move to 'Account_Number'.
22660         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22670       ElseIf InStr(strSortNow, "Account_Number") > 0 Then
              ' ** If this is now off, but the sort was on 'Account_Number', leave it there.
22680         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22690       End If
22700     End Select

22710     If InStr(strSortNow, "ChkStat") > 0 Then
22720       ChkStat_lbl_DblClick Cancel ' ** Procedure: Above.
22730     ElseIf InStr(strSortNow, "Check_Count") > 0 Then
22740       Check_Count_lbl_DblClick Cancel ' ** Procedure: Above.
22750     ElseIf InStr(strSortNow, "Last_Check_Number") > 0 Then
22760       Last_Check_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22770     ElseIf InStr(strSortNow, "Payee") > 0 Then
22780       Payee_lbl_DblClick Cancel ' ** Procedure: Above.
22790     ElseIf InStr(strSortNow, "Bank_Name") > 0 Then
22800       Bank_Name_lbl_DblClick Cancel ' ** Procedure: Above.
22810     ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
22820       Bank_Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
22830     ElseIf InStr(strSortNow, "Check_Amount") > 0 Then
22840       Check_Amount_lbl_DblClick Cancel ' ** Procedure: Above.
22850     End If

22860   End With

EXITP:
22870   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub ckgDisplay_opt02_ShortName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt02_ShortName_MouseMove"

22910   Detail_MouseMove Button, Shift, X, Y

EXITP:
22920   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Public Sub ckgDisplay_opt05_Payee_AfterUpdate()

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt05_Payee_AfterUpdate"

        Dim Cancel As Integer

23010   With Me
23020     Cancel = 0

23030     Select Case .ckgDisplay_opt05_Payee
          Case True
23040       .ckgDisplay_opt05_Payee_lbl.FontBold = True
23050       .ckgDisplay_opt05_Payee_lbl_dim.FontBold = True
23060       .ckgDisplay_opt05_Payee_lbl_dim_hi.FontBold = True
23070     Case False
23080       .ckgDisplay_opt05_Payee_lbl.FontBold = False
23090       .ckgDisplay_opt05_Payee_lbl_dim.FontBold = False
23100       .ckgDisplay_opt05_Payee_lbl_dim_hi.FontBold = False
23110     End Select

23120     ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
            lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
            lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
            lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me, .ckgDisplay_opt05_Payee  ' ** Module Procedure: modCheckPrintFuncs.
          'ColumnSet THIS_PROC, .ckgDisplay_opt05_Payee  ' ** Procedure: Below
23130     blnSortAgain = True  ' ** So it won't change the sort direction.

23140     Select Case .ckgDisplay_opt05_Payee
          Case True
23150       If InStr(strSortNow, "Account_Number") > 0 Then
23160         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
23170       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
23180         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
23190       ElseIf InStr(strSortNow, "Payee") > 0 Then
23200         Payee_lbl_DblClick Cancel ' ** Procedure: Above.
23210       End If
23220     Case False
23230       If InStr(strSortNow, "Account_Number") > 0 Then
23240         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
23250       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
23260         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
23270       ElseIf InStr(strSortNow, "Payee") > 0 Then
23280         If .Account_Number_lbl.Visible = True Then
23290           Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
23300         Else
23310           Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
23320         End If
23330       End If
23340     End Select

23350     If InStr(strSortNow, "ChkStat") > 0 Then
23360       ChkStat_lbl_DblClick Cancel ' ** Procedure: Above.
23370     ElseIf InStr(strSortNow, "Check_Count") > 0 Then
23380       Check_Count_lbl_DblClick Cancel ' ** Procedure: Above.
23390     ElseIf InStr(strSortNow, "Last_Check_Number") > 0 Then
23400       Last_Check_Number_lbl_DblClick Cancel ' ** Procedure: Above.
23410     ElseIf InStr(strSortNow, "Bank_Name") > 0 Then
23420       Bank_Name_lbl_DblClick Cancel ' ** Procedure: Above.
23430     ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
23440       Bank_Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
23450     ElseIf InStr(strSortNow, "Check_Amount") > 0 Then
23460       Check_Amount_lbl_DblClick Cancel ' ** Procedure: Above.
23470     End If

23480   End With

EXITP:
23490   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub ckgDisplay_opt05_Payee_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt05_Payee_MouseMove"

23510   Detail_MouseMove Button, Shift, X, Y

EXITP:
23520   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Public Sub ckgDisplay_opt06_BankName_AfterUpdate()

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt06_BankName_AfterUpdate"

        Dim Cancel As Integer

23610   With Me
23620     Cancel = 0

23630     Select Case .ckgDisplay_opt06_BankName
          Case True
23640       .ckgDisplay_opt06_BankName_lbl.FontBold = True
23650       .ckgDisplay_opt06_BankName_lbl_dim.FontBold = True
23660       .ckgDisplay_opt06_BankName_lbl_dim_hi.FontBold = True
23670     Case False
23680       .ckgDisplay_opt06_BankName_lbl.FontBold = False
23690       .ckgDisplay_opt06_BankName_lbl_dim.FontBold = False
23700       .ckgDisplay_opt06_BankName_lbl_dim_hi.FontBold = False
23710     End Select

23720     ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
            lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
            lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
            lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me, .ckgDisplay_opt06_BankName  ' ** Module Procedure: modCheckPrintFuncs.
          'ColumnSet THIS_PROC, .ckgDisplay_opt06_BankName  ' ** Procedure: Below
23730     blnSortAgain = True  ' ** So it won't change the sort direction.

23740     Select Case .ckgDisplay_opt05_Payee
          Case True
23750       If InStr(strSortNow, "Account_Number") > 0 Then
23760         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
23770       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
23780         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
23790       ElseIf InStr(strSortNow, "Bank_Name") > 0 Then
23800         Bank_Name_lbl_DblClick Cancel ' ** Procedure: Above.
23810       End If
23820     Case False
23830       If InStr(strSortNow, "Account_Number") > 0 Then
23840         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
23850       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
23860         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
23870       ElseIf InStr(strSortNow, "Bank_Name") > 0 Then
23880         If .Account_Number_lbl.Visible = True Then
23890           Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
23900         Else
23910           Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
23920         End If
23930       End If
23940     End Select

23950     If InStr(strSortNow, "ChkStat") > 0 Then
23960       ChkStat_lbl_DblClick Cancel ' ** Procedure: Above.
23970     ElseIf InStr(strSortNow, "Check_Count") > 0 Then
23980       Check_Count_lbl_DblClick Cancel ' ** Procedure: Above.
23990     ElseIf InStr(strSortNow, "Last_Check_Number") > 0 Then
24000       Last_Check_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24010     ElseIf InStr(strSortNow, "Payee") > 0 Then
24020       Payee_lbl_DblClick Cancel ' ** Procedure: Above.
24030     ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
24040       Bank_Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24050     ElseIf InStr(strSortNow, "Check_Amount") > 0 Then
24060       Check_Amount_lbl_DblClick Cancel ' ** Procedure: Above.
24070     End If

24080   End With

EXITP:
24090   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub ckgDisplay_opt06_BankName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt06_BankName_MouseMove"

24110   Detail_MouseMove Button, Shift, X, Y

EXITP:
24120   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Public Sub ckgDisplay_opt07_BankAcctNum_AfterUpdate()

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt07_BankAcctNum_AfterUpdate"

        Dim Cancel As Integer

24210   With Me
24220     Cancel = 0

24230     Select Case .ckgDisplay_opt07_BankAcctNum
          Case True
24240       .ckgDisplay_opt07_BankAcctNum_lbl.FontBold = True
24250       .ckgDisplay_opt07_BankAcctNum_lbl_dim.FontBold = True
24260       .ckgDisplay_opt07_BankAcctNum_lbl_dim_hi.FontBold = True
24270     Case False
24280       .ckgDisplay_opt07_BankAcctNum_lbl.FontBold = False
24290       .ckgDisplay_opt07_BankAcctNum_lbl_dim.FontBold = False
24300       .ckgDisplay_opt07_BankAcctNum_lbl_dim_hi.FontBold = False
24310     End Select

24320     ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
            lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
            lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
            lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me, .ckgDisplay_opt07_BankAcctNum  ' ** Module Procedure: modCheckPrintFuncs.
          'ColumnSet THIS_PROC, .ckgDisplay_opt07_BankAcctNum  ' ** Procedure: Below
24330     blnSortAgain = True  ' ** So it won't change the sort direction.

24340     Select Case .ckgDisplay_opt05_Payee
          Case True
24350       If InStr(strSortNow, "Account_Number") > 0 Then
24360         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24370       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
24380         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
24390       ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
24400         Bank_Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24410       End If
24420     Case False
24430       If InStr(strSortNow, "Account_Number") > 0 Then
24440         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24450       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
24460         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
24470       ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
24480         If .Account_Number_lbl.Visible = True Then
24490           Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24500         Else
24510           Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
24520         End If
24530       End If
24540     End Select

24550     If InStr(strSortNow, "ChkStat") > 0 Then
24560       ChkStat_lbl_DblClick Cancel ' ** Procedure: Above.
24570     ElseIf InStr(strSortNow, "Check_Count") > 0 Then
24580       Check_Count_lbl_DblClick Cancel ' ** Procedure: Above.
24590     ElseIf InStr(strSortNow, "Last_Check_Number") > 0 Then
24600       Last_Check_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24610     ElseIf InStr(strSortNow, "Payee") > 0 Then
24620       Payee_lbl_DblClick Cancel ' ** Procedure: Above.
24630     ElseIf InStr(strSortNow, "Bank_Name") > 0 Then
24640       Bank_Name_lbl_DblClick Cancel ' ** Procedure: Above.
24650     ElseIf InStr(strSortNow, "Check_Amount") > 0 Then
24660       Check_Amount_lbl_DblClick Cancel ' ** Procedure: Above.
24670     End If

24680   End With

EXITP:
24690   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Public Sub ckgDisplay_opt08_CheckAmount_AfterUpdate()

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDisplay_opt08_CheckAmount_AfterUpdate"

        Dim Cancel As Integer

24710   With Me
24720     Cancel = 0

24730     Select Case .ckgDisplay_opt08_CheckAmount
          Case True
24740       .ckgDisplay_opt08_CheckAmount_lbl.FontBold = True
24750       .ckgDisplay_opt08_CheckAmount_lbl_dim.FontBold = True
24760       .ckgDisplay_opt08_CheckAmount_lbl_dim_hi.FontBold = True
24770     Case False
24780       .ckgDisplay_opt08_CheckAmount_lbl.FontBold = False
24790       .ckgDisplay_opt08_CheckAmount_lbl_dim.FontBold = False
24800       .ckgDisplay_opt08_CheckAmount_lbl_dim_hi.FontBold = False
24810     End Select

24820     ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
            lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
            lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
            lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me, .ckgDisplay_opt08_CheckAmount  ' ** Module Procedure: modCheckPrintFuncs.
          'ColumnSet THIS_PROC, .ckgDisplay_opt08_CheckAmount  ' ** Procedure: Below
24830     blnSortAgain = True  ' ** So it won't change the sort direction.

          'I'M NOT SURE WHAT THIS IS FOR?
24840     Select Case .ckgDisplay_opt05_Payee
          Case True
24850       If InStr(strSortNow, "Account_Number") > 0 Then
24860         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24870       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
24880         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
24890       ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
24900         Bank_Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24910       End If
24920     Case False
24930       If InStr(strSortNow, "Account_Number") > 0 Then
24940         Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
24950       ElseIf InStr(strSortNow, "Short_Name") > 0 Then
24960         Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
24970       ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
24980         If .Account_Number_lbl.Visible = True Then
24990           Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
25000         Else
25010           Short_Name_lbl_DblClick Cancel ' ** Procedure: Above.
25020         End If
25030       End If
25040     End Select

25050     If InStr(strSortNow, "ChkStat") > 0 Then
25060       ChkStat_lbl_DblClick Cancel ' ** Procedure: Above.
25070     ElseIf InStr(strSortNow, "Check_Count") > 0 Then
25080       Check_Count_lbl_DblClick Cancel ' ** Procedure: Above.
25090     ElseIf InStr(strSortNow, "Last_Check_Number") > 0 Then
25100       Last_Check_Number_lbl_DblClick Cancel ' ** Procedure: Above.
25110     ElseIf InStr(strSortNow, "Payee") > 0 Then
25120       Payee_lbl_DblClick Cancel ' ** Procedure: Above.
25130     ElseIf InStr(strSortNow, "Bank_Account_Number") > 0 Then
25140       Bank_Account_Number_lbl_DblClick Cancel ' ** Procedure: Above.
25150     ElseIf InStr(strSortNow, "Bank_Name") > 0 Then
25160       Bank_Name_lbl_DblClick Cancel ' ** Procedure: Above.
25170     End If

25180   End With

EXITP:
25190   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgPayeeFont_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPayeeFont_box_MouseMove"

25210   Detail_MouseMove Button, Shift, X, Y

EXITP:
25220   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgPayeeFont_AfterUpdate()

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPayeeFont_AfterUpdate"

25310   With Me
25320     Select Case .opgPayeeFont
          Case .opgPayeeFont_opt10Pt.OptionValue
25330       .opgPayeeFont_opt10Pt_lbl.FontBold = True
25340       .opgPayeeFont_opt12Pt_lbl.FontBold = False
25350     Case .opgPayeeFont_opt12Pt.OptionValue
25360       .opgPayeeFont_opt10Pt_lbl.FontBold = False
25370       .opgPayeeFont_opt12Pt_lbl.FontBold = True
25380     End Select
25390   End With

EXITP:
25400   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgPayeeFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPayeeFont_MouseMove"

25510   Detail_MouseMove Button, Shift, X, Y

EXITP:
25520   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub opgPayeeFont_opt10Pt_KeyDown(KeyCode As Integer, Shift As Integer)

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPayeeFont_opt10Pt_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25620   intShiftDown = (Shift And acShiftMask) > 0
25630   intAltDown = (Shift And acAltMask) > 0
25640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25670       With Me
25680         intRetVal = 0
25690         If .cmdPreviewChecks.Enabled = True Then
25700           .cmdPreviewChecks.SetFocus
25710         Else
25720           .cmdClose.SetFocus
25730         End If
25740       End With
25750     End Select
25760   End If

        ' ** Shift keys.
25770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25790       With Me
25800         intRetVal = 0
25810         If .ckgDisplay_opt08_CheckAmount.Enabled = True Then
25820           .ckgDisplay_opt08_CheckAmount.SetFocus
25830         ElseIf .opgShow.Enabled = True Then
25840           .opgShow.SetFocus
25850         ElseIf .lbxShortAccountName.Enabled = True Then
25860           .lbxShortAccountName.SetFocus
25870         ElseIf .chkSyncListSort.Enabled = True Then
25880           .chkSyncListSort.SetFocus
25890         ElseIf .cmbSortBy.Enabled = True Then
25900           .cmbSortBy.SetFocus
25910         ElseIf .NextCheckNumber.Enabled = True Then
25920           .NextCheckNumber.SetFocus
25930         ElseIf .cmdMICRAdjust.Enabled = True Then
25940           .cmdMICRAdjust.SetFocus
25950         Else
25960           .opgPrint.SetFocus
25970         End If
25980       End With
25990     End Select
26000   End If

EXITP:
26010   KeyCode = intRetVal
26020   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub opgPayeeFont_opt12Pt_KeyDown(KeyCode As Integer, Shift As Integer)

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPayeeFont_opt12Pt_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26120   intShiftDown = (Shift And acShiftMask) > 0
26130   intAltDown = (Shift And acAltMask) > 0
26140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26170       With Me
26180         intRetVal = 0
26190         If .cmdPreviewChecks.Enabled = True Then
26200           .cmdPreviewChecks.SetFocus
26210         Else
26220           .cmdClose.SetFocus
26230         End If
26240       End With
26250     End Select
26260   End If

        ' ** Shift keys.
26270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26290       With Me
26300         intRetVal = 0
26310         If .ckgDisplay_opt08_CheckAmount.Enabled = True Then
26320           .ckgDisplay_opt08_CheckAmount.SetFocus
26330         ElseIf .opgShow.Enabled = True Then
26340           .opgShow.SetFocus
26350         ElseIf .lbxShortAccountName.Enabled = True Then
26360           .lbxShortAccountName.SetFocus
26370         ElseIf .chkSyncListSort.Enabled = True Then
26380           .chkSyncListSort.SetFocus
26390         ElseIf .cmbSortBy.Enabled = True Then
26400           .cmbSortBy.SetFocus
26410         ElseIf .NextCheckNumber.Enabled = True Then
26420           .NextCheckNumber.SetFocus
26430         ElseIf .cmdMICRAdjust.Enabled = True Then
26440           .cmdMICRAdjust.SetFocus
26450         Else
26460           .opgPrint.SetFocus
26470         End If
26480       End With
26490     End Select
26500   End If

EXITP:
26510   KeyCode = intRetVal
26520   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub opgPayeeFont_opt12Pt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgPayeeFont_opt12Pt_MouseMove"

26610   Detail_MouseMove Button, Shift, X, Y

EXITP:
26620   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdVoidedChecks_GotFocus()

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdVoidedChecks_GotFocus"

26710   With Me
26720     blnVoidedChecks_Focus = True
26730     .cmdVoidedChecks_raised_semifocus_dots_img.Visible = True
26740     .cmdVoidedChecks_raised_img.Visible = False
26750     .cmdVoidedChecks_raised_focus_img.Visible = False
26760     .cmdVoidedChecks_raised_focus_dots_img.Visible = False
26770     .cmdVoidedChecks_sunken_focus_dots_img.Visible = False
26780     .cmdVoidedChecks_raised_img_dis.Visible = False
26790   End With

EXITP:
26800   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdVoidedChecks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdVoidedChecks_MouseDown"

26910   With Me
26920     blnVoidedChecks_MouseDown = True
26930     .cmdVoidedChecks_sunken_focus_dots_img.Visible = True
26940     .cmdVoidedChecks_raised_img.Visible = False
26950     .cmdVoidedChecks_raised_semifocus_dots_img.Visible = False
26960     .cmdVoidedChecks_raised_focus_img.Visible = False
26970     .cmdVoidedChecks_raised_focus_dots_img.Visible = False
26980     .cmdVoidedChecks_raised_img_dis.Visible = False
26990   End With

EXITP:
27000   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdVoidedChecks_Click()

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdVoidedChecks_Click"

        Dim strDocName

27110   strDocName = "frmRpt_Checks_Void_Mgmt"
27120   DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME
27130   DoCmd.Hourglass False

EXITP:
27140   Exit Sub

ERRH:
3060    DoCmd.Hourglass False
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmdVoidedChecks_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdVoidedChecks_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27220   intShiftDown = (Shift And acShiftMask) > 0
27230   intAltDown = (Shift And acAltMask) > 0
27240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27260     Select Case intRetVal
          Case vbKeyTab
27270       With Me
27280         intRetVal = 0
27290         .cmdPreviewChecks.SetFocus
27300       End With
27310     End Select
27320   End If

        ' ** Shift keys.
27330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27340     Select Case intRetVal
          Case vbKeyTab
27350       With Me
27360         intRetVal = 0
27370         .opgPayeeFont.SetFocus
27380       End With
27390     End Select
27400   End If

EXITP:
27410   KeyCode = intRetVal
27420   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmdVoidedChecks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdVoidedChecks_MouseMove"

27510   With Me
27520     If blnVoidedChecks_MouseDown = False Then
27530       Select Case blnVoidedChecks_Focus
            Case True
27540         .cmdVoidedChecks_raised_focus_dots_img.Visible = True
27550         .cmdVoidedChecks_raised_focus_img.Visible = False
27560       Case False
27570         .cmdVoidedChecks_raised_focus_img.Visible = True
27580         .cmdVoidedChecks_raised_focus_dots_img.Visible = False
27590       End Select
27600       .cmdVoidedChecks_raised_img.Visible = False
27610       .cmdVoidedChecks_raised_semifocus_dots_img.Visible = False
27620       .cmdVoidedChecks_sunken_focus_dots_img.Visible = False
27630       .cmdVoidedChecks_raised_img_dis.Visible = False
27640     End If
27650   End With

EXITP:
27660   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdVoidedChecks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdVoidedChecks_MouseUp"

27710   With Me
27720     .cmdVoidedChecks_raised_focus_dots_img.Visible = True
27730     .cmdVoidedChecks_raised_img.Visible = False
27740     .cmdVoidedChecks_raised_semifocus_dots_img.Visible = False
27750     .cmdVoidedChecks_raised_focus_img.Visible = False
27760     .cmdVoidedChecks_sunken_focus_dots_img.Visible = False
27770     .cmdVoidedChecks_raised_img_dis.Visible = False
27780     blnVoidedChecks_MouseDown = False
27790   End With

EXITP:
27800   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdVoidedChecks_LostFocus()

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdVoidedChecks_LostFocus"

27910   With Me
27920     .cmdVoidedChecks_raised_img.Visible = True
27930     .cmdVoidedChecks_raised_semifocus_dots_img.Visible = False
27940     .cmdVoidedChecks_raised_focus_img.Visible = False
27950     .cmdVoidedChecks_raised_focus_dots_img.Visible = False
27960     .cmdVoidedChecks_sunken_focus_dots_img.Visible = False
27970     .cmdVoidedChecks_raised_img_dis.Visible = False
27980     blnVoidedChecks_Focus = False
27990   End With

EXITP:
28000   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub cmdPreviewChecks_Click()

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewChecks_Click"

        Dim strFont As String

28110   With Me

          ' ** If they've chosen the separate checking option, and then use the first option,
          ' ** one part of the code is expecting only 1 bank and account, and another part
          ' ** is expecting multiple banks and numbers!

28120     DoCmd.Hourglass True
28130     DoEvents

28140     Select Case .opgPrint

          Case .opgPrint_optAll.OptionValue        ' ** 1.
            ' ** All checks for all accounts, using a single checking account, on preprinted check stock.
            ' ** This prints checks for all Journal entries with JournalType = 'Paid', where PrintCheck = True.
            ' ** This could be multiple accounts, but because it's on preprinted check stock, it can only be one checking account.
28150       If .lbxShortAccountName.ListCount = 0 Then
28160         DoCmd.Hourglass False
28170         MsgBox "There are no checks to print!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
28180       Else
28190         PrintSingleAccount vbNullString, acViewPreview  ' ** Procedure: Below.
28200       End If

28210     Case .opgPrint_optSpecified.OptionValue  ' ** 2.
            ' ** All checks for a selected account, using a single checking account, on preprinted check stock.
            ' ** This prints checks like above, except that an account needs to be specified.
28220       If .lbxShortAccountName.ListCount = 0 Then
28230         DoCmd.Hourglass False
28240         MsgBox "There are no checks to print!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
28250       ElseIf IsNull(.lbxShortAccountName) = True Or IsNull(.lbxShortAccountName.Column(LBX_CHK_ACTNO)) = True Then
28260         DoCmd.Hourglass False
28270         MsgBox "Please select an account from which to print checks.", vbInformation + vbOKOnly, "Entry Required"
28280       Else
28290         If Len(Trim(.lbxShortAccountName.Column(LBX_CHK_ACTNO))) = 0 Then
28300           DoCmd.Hourglass False
28310           MsgBox "Please select an account from which to print checks.", vbInformation + vbOKOnly, "Entry Required"
28320         Else
28330           PrintSingleAccount .lbxShortAccountName.Column(LBX_CHK_ACTNO), acViewPreview  ' ** Procedure: Below.
28340         End If
28350       End If

28360     Case .opgPrint_optBlank.OptionValue      ' ** 3.
            ' ** All checks for a selected account, on blank stock; requires MICR font.
            ' ** Because it's on blank stock, a MICR font is needed in order to print the routing and account numbers.
28370       If .lbxShortAccountName.ListCount = 0 Then
28380         DoCmd.Hourglass False
28390         MsgBox "There are no checks to print!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
28400       ElseIf IsNull(.lbxShortAccountName) = True Or IsNull(.lbxShortAccountName.Column(LBX_CHK_ACTNO)) = True Then
28410         DoCmd.Hourglass False
28420         MsgBox "Please select an account from which to print checks.", vbInformation + vbOKOnly, "Entry Required"
28430       Else
28440         If Len(Trim(.lbxShortAccountName.Column(LBX_CHK_ACTNO))) = 0 Then
28450           DoCmd.Hourglass False
28460           MsgBox "Please select an account from which to print checks.", vbInformation + vbOKOnly, "Entry Required"
28470         Else
28480           strFont = IniFile_Get("General", "MICRFont", "#ERROR", CurrentAppPath & LNK_SEP & gstrFile_INI)  ' ** Module Procedure: modStartupFuncs, modFileUtilities.
28490           If strFont <> "#ERROR" Then
28500             .MICRFontName = strFont
28510             PrintBlankChecks .lbxShortAccountName.Column(LBX_CHK_ACTNO), acViewPreview  ' ** Procedure: Below.
28520           Else
28530             DoCmd.Hourglass False
28540             MsgBox "You must select the proper MICR font for your printer.", vbInformation + vbOKOnly, "Select MICR Font"
28550             .cmdMICRSet.SetFocus
28560             DoCmd.OpenForm "frmRpt_Checks_MICR_Set", , , , , , THIS_NAME
28570           End If
28580         End If
28590       End If

28600     Case Else
28610       DoCmd.Hourglass False
28620       MsgBox "Please select the Check Print Method.", vbInformation + vbOKOnly, "Select Check Print Method"

28630     End Select

28640     If blnNeedNext = True Then
28650       .NextCheckNumber.SetFocus
28660       blnNeedNext = False
28670     End If

28680     DoCmd.Hourglass False

28690   End With

EXITP:
28700   Exit Sub

ERRH:
3060    DoCmd.Hourglass False
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmdPreviewChecks_KeyDown(KeyCode As Integer, Shift As Integer)

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreviewChecks_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28820   intShiftDown = (Shift And acShiftMask) > 0
28830   intAltDown = (Shift And acAltMask) > 0
28840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28860     Select Case intRetVal
          Case vbKeyTab
28870       With Me
28880         intRetVal = 0
28890         .cmdPrintChecks.SetFocus
28900       End With
28910     End Select
28920   End If

        ' ** Shift keys.
28930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28940     Select Case intRetVal
          Case vbKeyTab
28950       With Me
28960         intRetVal = 0
28970         .opgPayeeFont.SetFocus
28980       End With
28990     End Select
29000   End If

EXITP:
29010   KeyCode = intRetVal
29020   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmdPrintChecks_Click()

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintChecks_Click"

        Dim strFont As String

29110   With Me

29120     Select Case .opgShow
          Case .opgShow_optAccount.OptionValue

29130       DoCmd.Hourglass True
29140       DoEvents

29150       Select Case .opgPrint

            Case .opgPrint_optAll.OptionValue        ' ** 1.
              ' ** All checks for all accounts, using a single checking account, on preprinted check stock.
              ' ** This prints checks for all Journal entries with JournalType = 'Paid', where PrintCheck = True.
              ' ** This could be multiple accounts, but because it's on preprinted check stock, it can only be one checking account.
29160         If .lbxShortAccountName.ListCount = 0 Then
29170           DoCmd.Hourglass False
29180           MsgBox "There are no checks to print!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
29190         Else
29200           PrintSingleAccount vbNullString, acViewNormal  ' ** Procedure: Below.
29210         End If

29220       Case .opgPrint_optSpecified.OptionValue  ' ** 2.
              ' ** All checks for a selected account, using a single checking account, on preprinted check stock.
              ' ** This prints checks like above, except that an account needs to be specified.
29230         If .lbxShortAccountName.ListCount = 0 Then
29240           DoCmd.Hourglass False
29250           MsgBox "There are no checks to print!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
29260         ElseIf IsNull(.lbxShortAccountName) = True Or IsNull(.lbxShortAccountName.Column(LBX_CHK_ACTNO)) = True Then
29270           DoCmd.Hourglass False
29280           MsgBox "Please select an account from which to print checks.", vbInformation + vbOKOnly, "Entry Required"
29290         Else
29300           If Len(Trim(.lbxShortAccountName.Column(LBX_CHK_ACTNO))) = 0 Then
29310             DoCmd.Hourglass False
29320             MsgBox "Please select an account from which to print checks.", vbInformation + vbOKOnly, "Entry Required"
29330           Else
29340             PrintSingleAccount .lbxShortAccountName.Column(LBX_CHK_ACTNO), acViewNormal  ' ** Procedure: Below.
29350           End If
29360         End If

29370       Case .opgPrint_optBlank.OptionValue      ' ** 3.
              ' ** All checks for a selected account, on blank stock; requires MICR font.
              ' ** Because it's on blank stock, a MICR font is needed in order to print the routing and account numbers.
29380         If .lbxShortAccountName.ListCount = 0 Then
29390           DoCmd.Hourglass False
29400           MsgBox "There are no checks to print!", vbExclamation + vbOKOnly, ("Nothing To Do" & Space(40))
29410         ElseIf IsNull(.lbxShortAccountName) = True Or IsNull(.lbxShortAccountName.Column(LBX_CHK_ACTNO)) = True Then
29420           DoCmd.Hourglass False
29430           MsgBox "Please select an account from which to print checks.", vbInformation + vbOKOnly, "Entry Required"
29440         Else
29450           If Len(Trim(.lbxShortAccountName.Column(LBX_CHK_ACTNO))) = 0 Then
29460             DoCmd.Hourglass False
29470             MsgBox "Please select an account from which to print checks.", vbInformation + vbOKOnly, "Entry Required"
29480           Else
29490             strFont = IniFile_Get("General", "MICRFont", "#ERROR", CurrentAppPath & LNK_SEP & gstrFile_INI)  ' ** Module Procedure: modStartupFuncs, modFileUtilities.
29500             If strFont <> "#ERROR" Then
29510               .MICRFontName = strFont
29520               PrintBlankChecks .lbxShortAccountName.Column(LBX_CHK_ACTNO), acViewNormal  ' ** Procedure: Below.
29530             Else
29540               DoCmd.Hourglass False
29550               MsgBox "You must select the proper MICR font for your printer.", vbInformation + vbOKOnly, "Select MICR Font"
29560               .cmdMICRSet.SetFocus
29570               DoCmd.OpenForm "frmRpt_Checks_MICR_Set", , , , , , THIS_NAME
29580             End If
29590           End If
29600         End If

29610       Case Else
29620         DoCmd.Hourglass False
29630         MsgBox "Please select the Check Print Method.", vbInformation + vbOKOnly, "Select Check Print Method"

29640       End Select

29650       If blnNeedNext = True Then
29660         .NextCheckNumber.SetFocus
29670         blnNeedNext = False
29680       End If

29690       DoCmd.Hourglass False

29700     Case .opgShow_optAll.OptionValue
29710       MsgBox "Though you may preview checks one-at-a-time," & vbCrLf & _
              "choose Checks By Account to print.", vbInformation + vbOKOnly, "Print Only With Checks By Account"
29720       .opgShow.SetFocus
29730     End Select

29740   End With

EXITP:
29750   Exit Sub

ERRH:
3060    DoCmd.Hourglass False
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub cmdPrintChecks_KeyDown(KeyCode As Integer, Shift As Integer)

29800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintChecks_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29820   intShiftDown = (Shift And acShiftMask) > 0
29830   intAltDown = (Shift And acAltMask) > 0
29840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29860     Select Case intRetVal
          Case vbKeyTab
29870       With Me
29880         intRetVal = 0
29890         .cmdClose.SetFocus
29900       End With
29910     End Select
29920   End If

        ' ** Shift keys.
29930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29940     Select Case intRetVal
          Case vbKeyTab
29950       With Me
29960         .cmdPreviewChecks.SetFocus
29970       End With
29980     End Select
29990   End If

EXITP:
30000   KeyCode = intRetVal
30010   Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub PrintSingleAccount(strAccountNo As String, intMode As Integer)
' ** These are pre-printed checks with a check number already present.
' ** opgPrint = opgPrint_optAll: 1. strAccountNo = vbNullString.
' ** opgPrint = opgPrint_optSpecified: 2.

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintSingleAccount"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngChecks As Long, lngConfirm As Long
        Dim lngPreviousCheckNum As Long, lngStartCheckNum As Long, lngLastCheckNumPrinted As Long
        Dim strDocName As String, strReportCriteria As String, strPlural As String
        Dim blnUseDefault As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant
        Dim blnContinue As Boolean

30110   blnContinue = True
30120   blnUseDefault = True

30130   With Me

30140     DoCmd.Hourglass True  ' ** Make sure it's still running.
30150     DoEvents

30160     lngPreviousCheckNum = 0&: lngStartCheckNum = 0&: lngLastCheckNumPrinted = 0&: lngConfirm = 0&

          ' ** Check for existing.
30170     If strAccountNo <> vbNullString Then
30180       If IsNull(.lbxShortAccountName.Column(LBX_CHK_CHKCNT)) = False Then
30190         lngChecks = CLng(.lbxShortAccountName.Column(LBX_CHK_CHKCNT))
30200         If .ChkStat <> "No" Then  ' ** AlreadyPrinted(), in modCheckPrintFuncs, sets overall status of group.
30210           Select Case .lbxShortAccountName.Column(LBX_CHK_STAT)
                Case "No"
                  ' ** All's well. Proceed.
30220           Case "Yes"
                  ' ** Shouldn't have been able to get here.
30230             lngChecks = 0&
30240           Case "Mixed"
                  ' ** Get the adjusted count.
30250             Set dbs = CurrentDb
                  ' ** Journal, just PrintCheck = True, CheckNum = Null, by specified [actno].
30260             Set qdf = dbs.QueryDefs("qryPrintChecks_05_31_01")
30270             With qdf.Parameters
30280               ![actno] = strAccountNo
30290             End With
30300             Set rst = qdf.OpenRecordset
30310             With rst
30320               If .BOF = True And .EOF = True Then
                      ' ** Shouldn't have been able to get here.
30330                 lngChecks = 0&
30340               Else
30350                 .MoveLast
30360                 lngChecks = .RecordCount
30370               End If
30380               .Close
30390             End With
30400             dbs.Close
30410             Set rst = Nothing
30420             Set qdf = Nothing
30430             Set dbs = Nothing
30440             DoEvents
30450           End Select
30460         End If  ' ** ChkStat.
30470       Else
30480         blnContinue = False
30490         lngChecks = 0&
30500         DoCmd.Hourglass False
30510         MsgBox "There are no checks to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
30520       End If
30530     Else
30540       varTmp00 = DLookup("[Checks]", "qryPrintChecks_03a")
30550       If IsNull(varTmp00) = False Then
30560         If varTmp00 > 0 Then
30570           lngChecks = CLng(varTmp00)
30580           Select Case .ChkStat  ' ** AlreadyPrinted(), in modCheckPrintFuncs, sets overall status of group.
                Case "No"
                  ' ** All's well. Proceed.
30590           Case "Yes"
                  ' ** Shouldn't have been able to get here.
30600             lngChecks = 0&
30610           Case "Mixed"
                  ' ** Get the adjusted count.
30620             Set dbs = CurrentDb
                  ' ** Journal, just PrintCheck = True, CheckNum = Null.
30630             Set qdf = dbs.QueryDefs("qryPrintChecks_05_30_01")
30640             Set rst = qdf.OpenRecordset
30650             With rst
30660               If .BOF = True And .EOF = True Then
                      ' ** Shouldn't have been able to get here.
30670                 lngChecks = 0&
30680               Else
30690                 .MoveLast
30700                 lngChecks = .RecordCount
30710               End If
30720               .Close
30730             End With
30740             dbs.Close
30750             Set rst = Nothing
30760             Set qdf = Nothing
30770             Set dbs = Nothing
30780             DoEvents
30790           End Select  ' ** ChkStat.
30800         Else
30810           lngChecks = 0&
30820         End If
30830       Else
30840         lngChecks = 0&
30850       End If
30860       If lngChecks = 0& Then
30870         blnContinue = False
30880         DoCmd.Hourglass False
30890         MsgBox "There are no checks to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
30900       End If
30910     End If  ' ** vbNullString.

          ' ** Check for next check number.
30920     If blnContinue = True Then
30930       If IsNull(.NextCheckNumber) = False Then
30940         If Val(.NextCheckNumber) > 0 Then
                ' ** OK to proceed.
30950           lngPreviousCheckNum = .NextCheckNumber - 1&
30960           lngStartCheckNum = .NextCheckNumber
30970           blnNeedNext = False
30980         Else
30990           blnContinue = False
31000           blnNeedNext = True
31010           Beep
31020           DoCmd.Hourglass False
31030           MsgBox "Please enter a Next Check Number.", vbInformation + vbOKOnly, "Entry Required"
31040         End If
31050       Else
31060         blnContinue = False
31070         blnNeedNext = True
31080         Beep
31090         DoCmd.Hourglass False
31100         MsgBox "Please enter a Next Check Number.", vbInformation + vbOKOnly, "Entry Required"
31110       End If
31120     End If  ' ** blnContinue.

          ' ** Advise user of checks to be printed.
31130     If blnContinue = True Then
31140       If lngChecks > 1& Then
31150         strPlural = "s"
31160       Else
31170         strPlural = vbNullString
31180       End If
31190       DoCmd.Hourglass False
31200       Select Case intMode
            Case acViewPreview
31210         msgResponse = MsgBox(CStr(lngChecks) & " check" & strPlural & " will be peviewed." & vbCrLf & vbCrLf & _
                "Continue?", vbInformation + vbOKCancel, "Preview Checks")
31220       Case acViewNormal
31230         msgResponse = MsgBox(CStr(lngChecks) & " check" & strPlural & " will be printed." & vbCrLf & vbCrLf & _
                "Check printer. Check forms should be in printer at this time." & vbCrLf & vbCrLf & _
                "Continue?", vbInformation + vbOKCancel, "Print Checks")
31240       End Select
31250       If msgResponse <> vbOK Then
31260         blnContinue = False
31270       End If
31280     End If  ' ** blnContinue.

          ' ** Write check numbers to Journal.
31290     If blnContinue = True Then
31300       DoCmd.Hourglass True  ' ** Make sure it's still running.
31310       DoEvents
            ' ** Only Journal entries with CheckNum = Null.
31320       lngLastCheckNumPrinted = FillJournalWithCheckNums_CP(strAccountNo, lngStartCheckNum, Me)  ' ** Module Function: modCheckPrintFuncs.
31330       If lngLastCheckNumPrinted = -999 Then
31340         blnContinue = False
31350         DoCmd.Hourglass False
31360         MsgBox "There was a problem assigning check numbers." & vbCrLf & vbCrLf & "Check printing failed.", _
                vbExclamation + vbOKOnly, "Journal Update Failed"
31370       End If
31380     End If  ' ** blnContinue.

31390     If blnContinue = True Then

31400       DoCmd.Hourglass True  ' ** Make sure it's still running.
31410       DoEvents

            ' ** Check for discrepancies in the Journal entry.
31420       Set dbs = CurrentDb
            ' ** Journal, just discrepancies in Paid's, with Ax, icash_new, pcash_new.
31430       Set qdf = dbs.QueryDefs("qryPrintChecks_25")
31440       Set rst = qdf.OpenRecordset
31450       If rst.BOF = True And rst.EOF = True Then
              ' ** All's well.
31460         rst.Close
31470       Else
31480         rst.Close
              ' ** Update qryPrintChecks_26 (Journal, with DLookups() to qryPrintChecks_25
              ' ** (Journal, just discrepancies in Paid's, with Ax, icash_new, pcash_new)).
31490         Set qdf = dbs.QueryDefs("qryPrintChecks_27")
31500         qdf.Execute
31510       End If
31520       Set rst = Nothing
31530       Set qdf = Nothing
31540       dbs.Close
31550       Set dbs = Nothing

            ' ** The report itself calls GetDollarString() from modUtilities.

            ' ** Now print or preview.
31560       strDocName = "rptChecks_Preprinted"
31570       Select Case gblnAdmin
            Case True
31580         strReportCriteria = vbNullString
31590       Case False
31600         strReportCriteria = "[journal_USER] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
31610       End Select
31620       If strAccountNo <> vbNullString Then
31630         If strReportCriteria <> vbNullString Then strReportCriteria = strReportCriteria & " And "
31640         strReportCriteria = strReportCriteria & "[accountno] = '" & strAccountNo & "'"
              '##GTR_Ref: rptChecks_Preprinted
31650         DoCmd.OpenReport strDocName, acViewPreview, , strReportCriteria, , _
                (.cmbSortBy.Column(CBX_SORT_FNAM) & "~" & "False" & "~" & CStr(intMode)) & "~" & IIf(strAccountNo = vbNullString, "All", strAccountNo)
31660       Else
31670         If strReportCriteria = vbNullString Then
31680           DoCmd.OpenReport strDocName, acViewPreview, , , , _
                  (.cmbSortBy.Column(CBX_SORT_FNAM) & "~" & "False" & "~" & CStr(intMode)) & "~" & IIf(strAccountNo = vbNullString, "All", strAccountNo)
31690         Else
31700           DoCmd.OpenReport strDocName, acViewPreview, , strReportCriteria, , _
                  (.cmbSortBy.Column(CBX_SORT_FNAM) & "~" & "False" & "~" & CStr(intMode)) & "~" & IIf(strAccountNo = vbNullString, "All", strAccountNo)
31710         End If
31720       End If

31730       If intMode = acViewPreview Then
31740         DoCmd.Maximize
31750         DoCmd.RunCommand acCmdFitToWindow
31760       End If

31770       DoEvents

31780     End If  ' ** blnContinue.

31790     Select Case intMode
          Case acViewPreview

            ' ** Clear out the check numbers put in above, using tblCheckPrint.
31800       If blnContinue = True Then
31810         Set dbs = CurrentDb
31820         With dbs
31830           If strAccountNo <> vbNullString Then
                  ' ** Update Journal, linked to tblCheckPrint, for CheckNum = Null, by specified [actno].
31840             Set qdf = .QueryDefs("qryPrintChecks_28a")
31850             With qdf.Parameters
31860               ![actno] = strAccountNo
31870             End With
31880             qdf.Execute
31890           Else
                  ' ** Update Journal, linked to tblCheckPrint, for CheckNum = Null
31900             Set qdf = .QueryDefs("qryPrintChecks_28b")
31910             qdf.Execute
31920           End If  ' ** vbNullString.
31930           .Close
31940         End With  ' ** dbs.
31950         Set qdf = Nothing
31960         Set dbs = Nothing
31970         DoEvents
31980       End If  ' ** blnContinue.

31990     Case acViewNormal

32000       If blnContinue = True Then

32010         DoCmd.Hourglass True  ' ** Make sure it's still running.
32020         DoEvents

32030 On Error Resume Next
32040         DoCmd.RunCommand acCmdPrint  ' ** Overlay the print options dialog.

32050         If ERR = 2501 Then  ' ** The '|' action was Canceled.
                ' ** User canceled printing.
32060 On Error GoTo ERRH
32070           blnContinue = False
32080           If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
32090             DoCmd.Close acReport, strDocName
32100           End If
32110         ElseIf ERR <> 0 Then  ' ** In case it's something else.
32120           zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32130 On Error GoTo ERRH
32140           blnContinue = False
32150           If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
32160             DoCmd.Close acReport, strDocName
32170           End If
32180         Else
32190 On Error GoTo ERRH
32200         End If

32210       End If  ' ** blnContinue.

            ' ** Check with user as to printing success.
32220       If blnContinue = True Then

32230         If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
32240           DoCmd.Close acReport, strDocName
32250         End If

              ' ** Did all print correctly?
              ' ** If the answer is 'No', these check numbers must be voided.
32260         strDocName = "frmRpt_Checks_Confirm"
32270         lngConfirm = 0&
32280         DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME
32290         lngConfirm = .opgConfirm  ' ** Populated by frmRpt_Checks_Confirm.

32300         DoCmd.Hourglass True  ' ** Make sure it's still running.
32310         DoEvents

              ' ** 1. ALL PRINTED CORRECTLY.
              ' ** 2. ALL PRINTED BADLY.
              ' ** 3. ONE OR MORE PRINTED BADLY.

32320         Set dbs = CurrentDb

32330         Select Case .opgConfirm
              Case 1  ' ** opgConfirm_optYes.

                ' **********************************
                ' ** All checks printed correctly.
                ' **********************************

                ' ** Write last check num to account table.
                ' ** Check numbers in the Journal will be copied to the Ledger during Posting.
32340           If strAccountNo <> vbNullString And strAccountNo <> "All" Then

                  ' ************************************
                  ' ** Update for a specified account.
                  ' ************************************

                  ' ** Update Account, by specified [actno], [chknum].
32350             Set qdf = dbs.QueryDefs("qryPrintChecks_06")
32360             With qdf.Parameters
32370               ![actno] = strAccountNo
32380               ![chknum] = lngLastCheckNumPrinted
32390             End With
32400             qdf.Execute
32410             Set qdf = Nothing
32420             DoEvents

32430           Else

                  ' *******************************
                  ' ** Update for a all accounts.
                  ' *******************************

                  ' ** qryPrintChecks_08 (Account, via DLookups to qryPrintChecks_07 (Journal, linked to tblCheckPrint,
                  ' ** grouped by accountno, with Max(CheckNum)), with LastCheckNum_new), just LastCheckNum_new <> Null.
32440             Set qdf = dbs.QueryDefs("qryPrintChecks_09")
32450             Set rst = qdf.OpenRecordset
32460             If rst.BOF = True And rst.EOF = True Then
32470               blnContinue = False
32480               rst.Close
32490               Set rst = Nothing
32500               Set qdf = Nothing
32510               DoEvents
32520             Else
32530               rst.Close
32540               Set rst = Nothing
32550               Set qdf = Nothing
32560               DoEvents
                    ' ** Update qryPrintChecks_09 (qryPrintChecks_08 (Account, via DLookups to qryPrintChecks_07 (Journal, linked to
                    ' ** tblCheckPrint, grouped by accountno, with Max(CheckNum)), with LastCheckNum_new), just LastCheckNum_new <> Null).
32570               Set qdf = dbs.QueryDefs("qryPrintChecks_10")
32580               qdf.Execute
32590               Set qdf = Nothing
32600               DoEvents
32610             End If

32620           End If

                '#################################################
                'DO WE WANT TO GRAB THE BANK AND ACCOUNT NUMBER?
                '#################################################

32630           If blnContinue = True Then

32640             DoCmd.Hourglass False
32650             MsgBox "Last check number saved.", vbInformation + vbOKOnly, "Last Check Number Saved"
32660             .lbxShortAccountName.Requery
32670             DoEvents

32680             DoCmd.Hourglass True
32690             DoEvents

                  ' ****************************************
                  ' ** Update next check number on screen.
                  ' ****************************************

32700             Select Case .opgPrint
                  Case .opgPrint_optAll.OptionValue
                    ' ** Both Ledger and Journal (not linked to tblCheckPrint):
                    ' ** qryPrintChecks_33a (qryPrintChecks_32c (Union of qryPrintChecks_31a
                    ' ** (qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
                    ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by
                    ' ** accountno, with Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b
                    ' ** (qryPrintChecks_32a (Journal, just CheckNum <> Null), grouped by accountno,
                    ' ** with Max(CheckNum), cnt)), grouped by accountno, with Max(CheckNum), cnt;
                    ' ** for AccountNo), grouped, with Max(CheckNum), cnt; for All.
32710               varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33b")
32720             Case Else
                    ' ** Both Ledger and Journal (not linked to tblCheckPrint):
                    ' ** qryPrintChecks_32c (Union of qryPrintChecks_31a (qryPrintChecks_30d
                    ' ** (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null), qryPrintChecks_30b
                    ' ** (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno, with
                    ' ** Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b (qryPrintChecks_32a
                    ' ** (Journal, just CheckNum <> Null), grouped by accountno, with Max(CheckNum),
                    ' ** cnt)), grouped by accountno, with Max(CheckNum), cnt; for AccountNo.
32730               varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33a", "[accountno] = '" & .lbxShortAccountName.Column(LBX_CHK_ACTNO) & "'")
32740             End Select  ' ** opgPrint.
32750             DoEvents

32760             If IsNull(varTmp00) = False Then
32770               If varTmp00 > 0 Then
32780                 .NextCheckNumber = CLng(varTmp00 + 1)
32790                 blnUseDefault = False
32800               Else
32810                 Select Case blnPrefCheckNum
                      Case True
32820                   If IsNull(.NextCheckNumber) = False Then
32830                     If Val(.NextCheckNumber) > 0 Then
                            ' ** Let it stand.
32840                       blnUseDefault = False
32850                     End If
32860                   End If
32870                 Case False
                        ' ** Use the default.
32880                 End Select
32890               End If
32900             End If

32910             DoCmd.Hourglass True  ' ** Make sure it's still running.
32920             DoEvents

                  ' ************
                  ' ** Finish.
                  ' ************

32930             If blnUseDefault = True Then
32940               .NextCheckNumber = 1000&
32950             End If
32960             DoEvents

32970           End If  ' ** blnContinue.

32980         Case 2  ' ** opgConfirm_optNo.

                ' *************************************
                ' ** All checks need to be reprinted.
                ' *************************************

                ' ** Save voided checks to tblCheckVoid.
32990           gblnMessage = True  ' ** For emergency closing of the forms.
33000           CheckVoid strAccountNo, True, lngChecks, lngChkVoid_Set, Me  ' ** Module Procedure: modCheckPrintFuncs.
33010           DoEvents

33020           If gblnMessage = True Then

33030             DoCmd.Hourglass True  ' ** Make sure it's still running.
33040             DoEvents

                  ' ******************************************
                  ' ** Clean out check numbers from Journal.
                  ' ******************************************

33050             blnContinue = False

33060             If strAccountNo <> vbNullString And strAccountNo <> "All" Then

                    ' **********************************
                    ' ** Delete for specified account.
                    ' **********************************

                    ' ** Update Journal, linked to tblCheckPrint, for CheckNum = Null, by specified [actno].
33070               Set qdf = dbs.QueryDefs("qryPrintChecks_05_01")
33080               With qdf.Parameters
33090                 ![actno] = strAccountNo
33100               End With
33110               qdf.Execute
33120               Set qdf = Nothing

33130             Else

                    ' **********************************
                    ' ** Delete for all accounts.
                    ' **********************************

                    ' ** Update Journal, linked to tblCheckPrint, for CheckNum = Null.
33140               Set qdf = dbs.QueryDefs("qryPrintChecks_05_02")
33150               qdf.Execute
33160               Set qdf = Nothing

33170             End If
33180             DoEvents

                  ' ****************************************
                  ' ** Update next check number on screen.
                  ' ****************************************

33190             Select Case .opgPrint
                  Case .opgPrint_optAll.OptionValue
                    ' ** Both Ledger and Journal (not linked to tblCheckPrint):
                    ' ** qryPrintChecks_33a (qryPrintChecks_32c (Union of qryPrintChecks_31a
                    ' ** (qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
                    ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by
                    ' ** accountno, with Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b
                    ' ** (qryPrintChecks_32a (Journal, just CheckNum <> Null), grouped by accountno,
                    ' ** with Max(CheckNum), cnt)), grouped by accountno, with Max(CheckNum), cnt;
                    ' ** for AccountNo), grouped, with Max(CheckNum), cnt; for All.
33200               varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33b")
33210             Case Else
                    ' ** Both Ledger and Journal (not linked to tblCheckPrint):
                    ' ** qryPrintChecks_32c (Union of qryPrintChecks_31a (qryPrintChecks_30d
                    ' ** (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null), qryPrintChecks_30b
                    ' ** (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno, with
                    ' ** Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b (qryPrintChecks_32a
                    ' ** (Journal, just CheckNum <> Null), grouped by accountno, with Max(CheckNum),
                    ' ** cnt)), grouped by accountno, with Max(CheckNum), cnt; for AccountNo.
33220               varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33a", "[accountno] = '" & .lbxShortAccountName.Column(LBX_CHK_ACTNO) & "'")
33230             End Select  ' ** opgPrint.
33240             DoEvents

33250             If IsNull(varTmp00) = False Then
33260               If varTmp00 > 0 Then
33270                 .NextCheckNumber = CLng(varTmp00 + 1)
33280                 blnUseDefault = False
33290               Else
33300                 Select Case blnPrefCheckNum
                      Case True
33310                   If IsNull(.NextCheckNumber) = False Then
33320                     If Val(.NextCheckNumber) > 0 Then
                            ' ** Let it stand.
33330                       blnUseDefault = False
33340                     End If
33350                   End If
33360                 Case False
                        ' ** Use the default.
33370                 End Select
33380               End If
33390             End If

                  ' ************
                  ' ** Finish.
                  ' ************

33400             DoCmd.Hourglass False
33410             MsgBox "Voided check numbers saved." & vbCrLf & _
                    "You may now reprint the checks.", vbInformation + vbOKOnly, "Ready To Reprint"

33420             DoCmd.Hourglass True  ' ** Make sure it's still running.
33430             DoEvents

33440           Else
33450             blnContinue = False
33460           End If  ' ** gblnMessage.

33470         Case 3  ' ** opgConfirm_optPartial.

                ' ***************************
                ' ** Choose checks to void.
                ' ***************************

33480           glngTaxCode_Distribution = 0&  ' ** Borrowing this variable.
33490           If strAccountNo = vbNullString Then
33500             strAccountNo = "All"
33510           End If
33520           gblnMessage = True  ' ** This will signify whether user changes mind and all checks are good.
33530           CheckVoid strAccountNo, False, lngChecks, lngChkVoid_Set, Me  ' ** Module Procedure: modCheckPrintFuncs.
33540           DoEvents

33550           If gblnMessage = True Then

33560             DoCmd.Hourglass True  ' ** Make sure it's still running.
33570             DoEvents

                  ' ********************************************
                  ' ** Bank info already saved and/or updated.
                  ' ********************************************

                  ' ************************************
                  ' ** Voids already in tblChecksVoid.
                  ' ************************************

                  ' *********************************************
                  ' ** Clear voided check numbers from Journal.
                  ' *********************************************

                  ' ** Update qryPrintChecks_05_33_02 (Journal, via subquery to
                  ' ** qryPrintChecks_05_33_01 (tblCheckVoid, linked to tblCheckVoid_Staging,
                  ' ** Journal entries to clear CheckNum)), for CheckNum = Null.
33580             Set qdf = dbs.QueryDefs("qryPrintChecks_05_33_03")
33590             qdf.Execute
33600             Set qdf = Nothing
33610             DoEvents

                  ' ****************************************
                  ' ** Update next check number on screen.
                  ' ****************************************

33620             Select Case .opgPrint
                  Case .opgPrint_optAll.OptionValue
                    ' ** Both Ledger and Journal (not linked to tblCheckPrint):
                    ' ** qryPrintChecks_33a (qryPrintChecks_32c (Union of qryPrintChecks_31a
                    ' ** (qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
                    ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by
                    ' ** accountno, with Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b
                    ' ** (qryPrintChecks_32a (Journal, just CheckNum <> Null), grouped by accountno,
                    ' ** with Max(CheckNum), cnt)), grouped by accountno, with Max(CheckNum), cnt;
                    ' ** for AccountNo), grouped, with Max(CheckNum), cnt; for All.
33630               varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33b")
33640             Case Else
                    ' ** Both Ledger and Journal (not linked to tblCheckPrint):
                    ' ** qryPrintChecks_32c (Union of qryPrintChecks_31a (qryPrintChecks_30d
                    ' ** (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null), qryPrintChecks_30b
                    ' ** (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno, with
                    ' ** Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b (qryPrintChecks_32a
                    ' ** (Journal, just CheckNum <> Null), grouped by accountno, with Max(CheckNum),
                    ' ** cnt)), grouped by accountno, with Max(CheckNum), cnt; for AccountNo.
33650               varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33a", "[accountno] = '" & .lbxShortAccountName.Column(LBX_CHK_ACTNO) & "'")
33660             End Select  ' ** opgPrint.
33670             DoEvents

33680             If IsNull(varTmp00) = False Then
33690               If varTmp00 > 0 Then
33700                 .NextCheckNumber = CLng(varTmp00 + 1)
33710                 blnUseDefault = False
33720               Else
33730                 Select Case blnPrefCheckNum
                      Case True
33740                   If IsNull(.NextCheckNumber) = False Then
33750                     If Val(.NextCheckNumber) > 0 Then
                            ' ** Let it stand.
33760                       blnUseDefault = False
33770                     End If
33780                   End If
33790                 Case False
                        ' ** Use the default.
33800                 End Select
33810               End If
33820             End If

                  ' ************
                  ' ** Finish.
                  ' ************

33830             DoCmd.Hourglass False
33840             MsgBox "The chosen voided check numbers have been saved." & vbCrLf & _
                    "You may now reprint the checks.", vbInformation + vbOKOnly, "Ready To Reprint"

33850             DoCmd.Hourglass True  ' ** Make sure it's still running.
33860             DoEvents

33870           Else
                  ' ** Don't show reprint message, do show successful message.
33880           End If

33890         End Select  ' ** opgConfirm.
33900         dbs.Close

33910         If blnContinue = True Then

33920           DoCmd.Hourglass True  ' ** Make sure it's still running.
33930           DoEvents

                ' **************************
                ' ** Refresh the list box.
                ' **************************

33940           Set dbs = CurrentDb

                ' ** Empty tblJournal_Check_Staging1.
33950           Set qdf = dbs.QueryDefs("qryPrintChecks_16a")
33960           qdf.Execute
33970           Set qdf = Nothing
33980           DoEvents
                ' ** Empty tblJournal_Check_Staging2.
33990           Set qdf = dbs.QueryDefs("qryPrintChecks_16b")
34000           qdf.Execute
34010           Set qdf = Nothing
34020           DoEvents
                ' ** Append qryPrintChecks_17a (qryPrintChecks_01c (Journal, just 'Misc.', 'Paid', 'Received',
                ' ** for PrintCheck = True, with Payee), with add'l fields) to tblJournal_Check_Staging1.
34030           Set qdf = dbs.QueryDefs("qryPrintChecks_18a")
34040           qdf.Execute
34050           Set qdf = Nothing
34060           DoEvents
                ' ** Append qryPrintChecks_17c (qryPrintChecks_17b (qryPrintChecks_17a (qryPrintChecks_01c
                ' ** (Journal, just 'Misc.', 'Paid', 'Received', for PrintCheck = True, with Payee), with
                ' ** add'l fields), grouped and summed, by Account Number, with Payeex), with Payee)
                ' ** to tblJournal_Check_Staging2.
34070           Set qdf = dbs.QueryDefs("qryPrintChecks_18b")
34080           qdf.Execute
34090           Set qdf = Nothing
34100           DoEvents

34110           dbs.Close

34120           .lbxShortAccountName.Requery

34130           If lngChecks > 1 Then
34140             strPlural = "s have"
34150           Else
34160             strPlural = " has"
34170           End If
34180           DoCmd.Hourglass False
34190           If glngTaxCode_Distribution <> 0& Then
34200             MsgBox CStr(lngChecks) & " check" & strPlural & " been printed," & vbCrLf & _
                    "of which " & CStr(Abs(glngTaxCode_Distribution)) & " " & _
                    IIf(glngTaxCode_Distribution = 1, "has", "have") & " been voided." & vbCrLf & vbCrLf & _
                    "Please review for accuracy, and print the Check List for your records.", vbInformation + vbOKOnly, ("Finished" & Space(40))
34210           Else
34220             MsgBox CStr(lngChecks) & " check" & strPlural & " been printed." & vbCrLf & vbCrLf & _
                    "Please review for accuracy, and print the Check List for your records.", vbInformation + vbOKOnly, ("Finished" & Space(40))
34230           End If

34240           AlreadyPrinted Me  ' ** Module Function: modCheckPrintFuncs.
34250           DoEvents

                ' ********************************************
                ' ** Show 'Printed' column when appropriate.
                ' ********************************************

34260           If .ChkStat = "Yes" Or .ChkStat = "Mixed" Then
                  ' ** I think the only way to add the 'Printed' column is by re-calling one of the options.
34270             If .ckgDisplay_opt01_AccountNo = True Then
34280               ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
                      lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
                      lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
                      lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me, .ckgDisplay_opt01_AccountNo  ' ** Module Procedure: modCheckPrintFuncs.
34290             ElseIf .ckgDisplay_opt02_ShortName = True Then
34300               ColumnSet THIS_PROC, lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, lngThisFrm_Width, _
                      lngFrmMin_Width, lngFrmWidthOffset, lngClose_Offset, lngLbx_Width, lngLbxBoxOffset, strLbx_ColWidths, _
                      lngLbx_Col00, lngLbx_Col01, lngLbx_Col02, lngLbx_Col03, lngLbx_Col04, lngLbx_Col05, lngLbx_Col06, _
                      lngLbx_Col07, lngLbx_Col08, lngLbx_Col09, lngGTR_Offset, lngMonitorCnt, lngMonitorNum, Me, .ckgDisplay_opt02_ShortName  ' ** Module Procedure: modCheckPrintFuncs.
34310             End If
34320           End If
34330           DoEvents

34340         End If  ' ** blnContinue.

34350       End If  ' ** blnContinue.

34360     End Select  ' ** intMode.

34370     DoCmd.Hourglass False

34380   End With

EXITP:
34390   Set rst = Nothing
34400   Set qdf = Nothing
34410   Set dbs = Nothing
34420   Exit Sub

ERRH:
3060    DoCmd.Hourglass False
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub PrintBlankChecks(strAccountNo As String, intMode As Integer)
' ** These are blank checks without a pre-existing check number.
' ** opgPrint = opgPrint_optBlank: 3.

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintBlankChecks"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngChecks As Long
        Dim lngPreviousCheckNum As Long, lngStartCheckNum As Long, lngLastCheckNumPrinted As Long
        Dim strDocName As String, strReportCriteria As String, strPlural As String
        Dim blnUseDefault As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, lngTmp01 As Long
        Dim blnContinue As Boolean

34510   blnContinue = True
34520   blnUseDefault = True

34530   With Me

34540     lngPreviousCheckNum = 0&: lngStartCheckNum = 0&: lngLastCheckNumPrinted = 0&

34550     If IsNull(.lbxShortAccountName.Column(LBX_CHK_CHKCNT)) = False Then

            ' ** Check for existing.
34560       If IsNull(.lbxShortAccountName.Column(LBX_CHK_CHKCNT)) = False Then
34570         Select Case .opgShow
              Case .opgShow_optAccount.OptionValue
34580           lngChecks = CLng(.lbxShortAccountName.Column(LBX_CHK_CHKCNT))
34590         Case .opgShow_optAll.OptionValue
34600           gstrAccountNo = .lbxShortAccountName.Column(LBX_CHK_ACTNO)
34610           lngChecks = CLng(DLookup("[Check_Count]", "qryPrintChecks_01_01a_Account_Number", "[Account_Number] = '" & gstrAccountNo & "'"))
34620         End Select
34630       Else
34640         blnContinue = False
34650         DoCmd.Hourglass False
34660         MsgBox "There are no checks to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
34670       End If

34680       If blnContinue = True Then
34690         If lngChecks > 1& Then
34700           strPlural = "s"
34710         Else
34720           strPlural = vbNullString
34730         End If
34740         Select Case intMode
              Case acViewPreview
34750           DoCmd.Hourglass False
34760           msgResponse = MsgBox(CStr(lngChecks) & " check" & strPlural & " will be previewed." & vbCrLf & vbCrLf & _
                  "Any previous check numbers will be cleared." & vbCrLf & vbCrLf & "Continue?", vbInformation + vbOKCancel, "Print Checks")
34770         Case acViewNormal
34780           msgResponse = MsgBox(CStr(lngChecks) & " check" & strPlural & " will be printed." & vbCrLf & vbCrLf & _
                  "Check printer. Blank check stock should be in printer at this time." & vbCrLf & vbCrLf & _
                  "Continue?", vbInformation + vbOKCancel, "Print Checks")
34790         End Select
34800         If msgResponse <> vbOK Then
34810           blnContinue = False
34820         End If
34830       End If  ' ** blnContinue.

34840       If blnContinue = True Then
34850         DoCmd.Hourglass True
34860         DoEvents
              ' ** Check for next check number.
34870         If IsNull(.NextCheckNumber) = False Then
34880           If Val(.NextCheckNumber) > 0 Then
                  ' ** OK to proceed.
34890             lngPreviousCheckNum = .NextCheckNumber - 1&
34900             lngStartCheckNum = .NextCheckNumber
34910             blnNeedNext = False
34920           Else
34930             blnContinue = False
34940             DoCmd.Hourglass False
34950             blnNeedNext = True
34960             Beep
34970             MsgBox "Please enter a Next Check Number.", vbInformation + vbOKOnly, "Entry Required"
34980           End If
34990         Else
35000           blnContinue = False
35010           DoCmd.Hourglass False
35020           blnNeedNext = True
35030           Beep
35040           MsgBox "Please enter a Next Check Number.", vbInformation + vbOKOnly, "Entry Required"
35050         End If
35060       End If  ' ** blnContinue.

            ' ** Write check numbers to Journal.
35070       If blnContinue = True Then
35080         lngLastCheckNumPrinted = FillJournalWithCheckNums_CP(strAccountNo, lngStartCheckNum, Me)  ' ** Module Function: modCheckPrintFuncs.
35090         If lngLastCheckNumPrinted = -999 Then
35100           blnContinue = False
35110           DoCmd.Hourglass False
35120           MsgBox "There was a problem assigning check numbers." & vbCrLf & vbCrLf & "Check printing failed.", _
                  vbExclamation + vbOKOnly, "Journal Update Failed"
35130         End If
35140       End If  ' ** blnContinue.

35150       If blnContinue = True Then

35160         Set dbs = CurrentDb
              ' ** Journal, just discrepancies in Paid's, with Ax, icash_new, pcash_new.
35170         Set qdf = dbs.QueryDefs("qryPrintChecks_25")
35180         Set rst = qdf.OpenRecordset
35190         If rst.BOF = True And rst.EOF = True Then
                ' ** All's well.
35200           rst.Close
35210         Else
35220           rst.Close
                ' ** Update qryPrintChecks_26 (Journal, with DLookups() to qryPrintChecks_25
                ' ** (Journal, just discrepancies in Paid's, with Ax, icash_new, pcash_new)).
35230           Set qdf = dbs.QueryDefs("qryPrintChecks_27")
35240           qdf.Execute
35250         End If
35260         Set rst = Nothing
35270         Set qdf = Nothing
35280         dbs.Close
35290         Set dbs = Nothing

              ' ** The report itself calls GetDollarString() from modUtilities.

              ' ** Now print.
35300         Select Case gblnAdmin
              Case True
35310           lngTmp01 = 0&
35320           strDocName = "frmRpt_Checks_Multi"
35330           Set dbs = CurrentDb
35340           With dbs
                  ' ** qryRpt_Checks_Multi_03_01 (Journal, grouped by journal_USER, just PrintCheck = True,
                  ' ** with cnt_chks, by specified [actno]), grouped, with cnt_users.
35350             Set qdf = .QueryDefs("qryRpt_Checks_Multi_03_02")
35360             With qdf.Parameters
35370               ![actno] = strAccountNo
35380             End With
35390             Set rst = qdf.OpenRecordset
35400             With rst
35410               If .BOF = True And .EOF = True Then
                      ' ** Something's messed up!
35420                 blnContinue = False
35430                 strDocName = vbNullString
35440               Else
35450                 .MoveFirst
35460                 lngTmp01 = ![cnt_users]
35470                 Select Case lngTmp01
                      Case 0&
                        ' ** Same sentiments!
35480                   blnContinue = False
35490                   strDocName = vbNullString
35500                 Case 1&
35510                   If ![journal_USER] = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
                          ' ** No need for choice.
35520                     strDocName = vbNullString
35530                   Else
                          ' ** Bring up multi form.
35540                   End If
35550                 Case Else
                        ' ** Bring up multi form.
35560                 End Select
35570               End If
35580               .Close
35590             End With
35600             Set rst = Nothing
35610             Set qdf = Nothing
35620             .Close
35630           End With
35640           Set dbs = Nothing
35650           If strDocName <> vbNullString Then
                  ' ** Only bring up if not or more than current user.
35660             gstrJournalUser = vbNullString
35670             gstrFormQuerySpec = strDocName
35680             DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME
35690             Select Case gstrJournalUser
                  Case vbNullString
35700               blnContinue = False
35710               strReportCriteria = vbNullString
35720             Case "All"
35730               strReportCriteria = vbNullString
35740             Case Else
35750               strReportCriteria = "[journal_USER] = '" & gstrJournalUser & "' And "
35760             End Select
35770             DoCmd.Hourglass False
35780           Else
35790             strReportCriteria = vbNullString
35800           End If
35810         Case False
35820           strReportCriteria = "[journal_USER] = '" & CurrentUser & "' And "  ' ** Internal Access Function: Trust Accountant login.
35830         End Select

35840         If blnContinue = True Then

35850           DoCmd.Hourglass False

35860           strDocName = "rptChecks_Blank"
35870           strReportCriteria = strReportCriteria & "[accountno] = '" & strAccountNo & "'"
                '##GTR_Ref: rptChecks_Blank
35880           DoCmd.OpenReport strDocName, acViewPreview, , strReportCriteria, , _
                  (.cmbSortBy.Column(CBX_SORT_FNAM) & "~" & "False" & "~" & CStr(intMode)) & "~" & IIf(strAccountNo = vbNullString, "All", strAccountNo)

35890           If intMode = acViewPreview Then
35900             DoCmd.Maximize
35910             DoCmd.RunCommand acCmdFitToWindow
35920           End If

35930           DoEvents
35940         End If  ' ** blnContinue.

35950       End If  ' ** blnContinue.

35960       If blnContinue = True Then

35970         Select Case intMode
              Case acViewPreview

35980           If blnContinue = True Then
                  ' ** Clear out the check numbers put in above.
35990             Set dbs = CurrentDb
36000             With dbs
36010               If strAccountNo <> vbNullString Then
                      ' ** Update Journal, for CheckNum = Null, by specified [actno].
36020                 Set qdf = .QueryDefs("qryPrintChecks_28a")
36030                 With qdf.Parameters
36040                   ![actno] = strAccountNo
36050                 End With
36060                 qdf.Execute
36070               Else
                      ' ** Update Journal, for CheckNum = Null
36080                 Set qdf = .QueryDefs("qryPrintChecks_28b")
36090                 qdf.Execute
36100               End If  ' ** vbNullString.
36110               .Close
36120             End With  ' ** dbs.
36130             Set qdf = Nothing
36140             Set dbs = Nothing
36150           End If  ' ** blnContinue.

36160         Case acViewNormal

36170           If blnContinue = True Then

36180 On Error Resume Next
36190             DoCmd.RunCommand acCmdPrint  ' ** Overlay the print options dialog.

36200             If ERR = 2501 Then  ' ** The '|' action was Canceled.
                    ' ** User canceled printing.
36210 On Error GoTo ERRH
36220               blnContinue = False
36230               If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
36240                 DoCmd.Close acReport, strDocName
36250               End If
36260             ElseIf ERR <> 0 Then  ' ** In case it's something else.
36270               zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
36280 On Error GoTo ERRH
36290               blnContinue = False
36300               If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
36310                 DoCmd.Close acReport, strDocName
36320               End If
36330             Else
36340 On Error GoTo ERRH
36350             End If

36360           End If  ' ** blnContinue.

36370           If blnContinue = True Then

36380             If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
36390               DoCmd.Close acReport, strDocName
36400             End If

                  ' ** Did all print correctly?
                  ' ** No voided check issue since Trust Accountant assigns the check number.
36410             msgResponse = MsgBox("Did all checks print correctly?" & vbCrLf & vbCrLf & _
                    "Leave this question box on the screen until your checks have finished printing." & vbCrLf & vbCrLf & _
                    "Click Yes if the checks look correct." & vbCrLf & _
                    "Click No to reset the saved check numbers.", vbQuestion + vbYesNo, "Confirm After Printing Finished")

36420             If msgResponse <> vbYes Then
                    ' ** Delete all check numbers for that account from Journal table.
36430               blnContinue = False
                    'strSQL = "UPDATE Journal SET [CheckNum] = NULL WHERE [accountno] = '" & strAccountNo & "';"
                    ' ** Update Journal, for CheckNum = Null, by specified [actno].
36440               Set dbs = CurrentDb
36450               With dbs
36460                 Set qdf = .QueryDefs("qryPrintChecks_05_01")
36470                 With qdf.Parameters
36480                   ![actno] = strAccountNo
36490                 End With
36500                 qdf.Execute
36510                 .Close
36520               End With
36530               MsgBox "Check numbers reset.", vbInformation + vbOKOnly, "Check Numbers Rolled Back"
36540             Else
                    ' ** Write last check num to account table.
                    'strSQL = "UPDATE Account SET [LastCheckNum] = " & lngLastCheckNumPrinted & " WHERE [accountno] = '" & strAccountNo & "';"
36550               Set dbs = CurrentDb
36560               With dbs
36570                 Set qdf = dbs.QueryDefs("qryPrintChecks_06")
36580                 With qdf.Parameters
36590                   ![actno] = strAccountNo
36600                   ![chknum] = lngLastCheckNumPrinted
36610                 End With
36620                 qdf.Execute
36630                 .Close
36640               End With
36650               MsgBox "Last check number saved.", vbInformation + vbOKOnly, "Last Check Number Saved"
36660               .lbxShortAccountName.Requery
36670               DoEvents
36680               Select Case .opgPrint
                    Case .opgPrint_optAll.OptionValue
                      ' ** Just Ledger:
                      ' ** qryPrintChecks_31a (qryPrintChecks_30d (Union of qryPrintChecks_30a
                      ' ** (Ledger, CheckNum <> Null), qryPrintChecks_30b (LedgerArchive, just
                      ' ** CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno, with Max(CheckNum), cnt;
                      ' ** by AccountNo), grouped, with Max(CheckNum), cnt, for All.
                      'varTmp00 = DLookup("[LastCheckNum]", "qryPrintChecks_31b")
                      ' **
                      ' ** Both Ledger and Journal:
                      ' ** qryPrintChecks_33a (qryPrintChecks_32c (Union of qryPrintChecks_31a
                      ' ** (qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
                      ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by
                      ' ** accountno, with Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b
                      ' ** (qryPrintChecks_32a (Journal, just CheckNum <> Null), grouped by accountno,
                      ' ** with Max(CheckNum), cnt)), grouped by accountno, with Max(CheckNum), cnt;
                      ' ** for AccountNo), grouped, with Max(CheckNum), cnt; for All.
36690                 varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33b")
36700               Case Else
                      ' ** Just Ledger:
                      ' ** qryPrintChecks_30d (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null),
                      ' ** qryPrintChecks_30b (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno,
                      ' ** with Max(CheckNum), cnt; by AccountNo.
                      'varTmp00 = DLookup("[LastCheckNum]", "qryPrintChecks_31a", "[accountno] = '" & .lbxShortAccountName.Column(LBX_CHK_ACTNO) & "'")
                      ' **
                      ' ** Both Ledger and Journal:
                      ' ** qryPrintChecks_32c (Union of qryPrintChecks_31a (qryPrintChecks_30d
                      ' ** (Union of qryPrintChecks_30a (Ledger, CheckNum <> Null), qryPrintChecks_30b
                      ' ** (LedgerArchive, just CheckNum <> Null), qryPrintChecks_30c (tblCheckVoid, just CheckNum)), grouped by accountno, with
                      ' ** Max(CheckNum), cnt; by AccountNo), qryPrintChecks_32b (qryPrintChecks_32a
                      ' ** (Journal, just CheckNum <> Null), grouped by accountno, with Max(CheckNum),
                      ' ** cnt)), grouped by accountno, with Max(CheckNum), cnt; for AccountNo.
36710                 varTmp00 = DLookup("[CheckNum]", "qryPrintChecks_33a", "[accountno] = '" & .lbxShortAccountName.Column(LBX_CHK_ACTNO) & "'")
36720               End Select
36730               If IsNull(varTmp00) = False Then
36740                 If varTmp00 > 0 Then
36750                   .NextCheckNumber = CLng(varTmp00 + 1)
36760                   blnUseDefault = False
36770                 Else
36780                   Select Case blnPrefCheckNum
                        Case True
36790                     If IsNull(.NextCheckNumber) = False Then
36800                       If Val(.NextCheckNumber) > 0 Then
                              ' ** Let it stand.
36810                         blnUseDefault = False
36820                       End If
36830                     End If
36840                   Case False
                          ' ** Use the default.
36850                   End Select
36860                 End If
36870               End If
36880               DoEvents
36890               If blnUseDefault = True Then
36900                 .NextCheckNumber = 1000&
36910               End If
36920             End If

36930             If msgResponse = vbYes Then
36940               If lngChecks > 1 Then
36950                 strPlural = "s have"
36960               Else
36970                 strPlural = " has"
36980               End If
36990               MsgBox CStr(lngChecks) & " check" & strPlural & " been printed." & vbCrLf & vbCrLf & _
                      "Please review for accuracy, and print the Check List for your records.", vbInformation + vbOKOnly, ("Finished" & Space(40))
37000             End If

37010           End If  ' ** blnContinue.

37020         End Select  ' ** intMode.

37030       End If  ' ** blnContinue.

37040     Else
37050       DoCmd.Hourglass False
37060       MsgBox "A problem with the check list has been detected." & vbCrLf & _
              "Please close this window and reopen it.", vbInformation + vbOKOnly, "Error Detected"
37070       .cmdClose.SetFocus
37080     End If
37090     DoCmd.Hourglass False
37100   End With

EXITP:
37110   Set rst = Nothing
37120   Set qdf = Nothing
37130   Set dbs = Nothing
37140   Exit Sub

ERRH:
3060    DoCmd.Hourglass False
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub SortNow(strProc As String)
' ** strSortNow is a query name.
' ** Sorting is accomplished by a pre-written group of queries
' ** that cover every combination of fields and sorts, both
' ** ascending and descending.
' ** If blnSortAgain = True, don't change sort direction.

37200 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim blnAll As Boolean
        Dim intPos01 As Integer
        Dim intX As Integer

37210   With Me

37220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
37230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
37240     End If

37250     Select Case .opgShow
          Case .opgShow_optAccount.OptionValue
37260       blnAll = False
37270     Case .opgShow_optAll.OptionValue
37280       blnAll = True
37290     End Select

37300     If .lbxShortAccountName.Enabled = True Then
37310       .Controls(strSortLblAsc).Visible = False
37320       .Controls(strSortLblDesc).Visible = False
37330       .Controls(strSortLine1).Visible = False
37340       .Controls(strSortLine2).Visible = False
37350       .Controls(strSortLine1).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
37360       .Controls(strSortLine2).Width = lngTpp
37370       If strProc = "Form_Load" Then
37380         strCalled = "Account_Number"
37390         Select Case blnAll
              Case True
37400           strSortNow = StringReplace(strSortOrig, "_01a_", "_01c_")  ' ** Module Function: modStringFuncs.
37410         Case False
                ' ** qryPrintChecks_01_01a_Account_Number
37420           strSortNow = strSortOrig
37430         End Select
37440         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
37450         lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (2& * lngTpp))
37460         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
37470         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
37480         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
37490         .Controls(strSortLblAsc).Top = lngSortLbl_Top
37500         .Controls(strSortLblDesc).Top = lngSortLbl_Top
37510         .Controls(strSortLblAsc).Left = lngSortLbl_Left
37520         .Controls(strSortLblDesc).Left = lngSortLbl_Left
37530         .Controls(strSortLine1).Top = lngSortLine_Top
37540         .Controls(strSortLine2).Top = (lngSortLine_Top + lngTpp)
37550         .Controls(strSortLine1).Left = lngSortLine_Left
37560         .Controls(strSortLine2).Left = lngSortLine_Left
37570         .Controls(strSortLine1).Width = lngSortLine_Width
37580         .Controls(strSortLine2).Width = lngSortLine_Width
37590         .Controls(strSortLblAsc).Visible = True
37600         .Controls(strSortLblDesc).Visible = False
37610       Else
37620         strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
37630         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
37640         lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (2& * lngTpp))
37650         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - (2& * lngTpp))
37660         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
37670         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
37680         .Controls(strSortLblAsc).Top = lngSortLbl_Top
37690         .Controls(strSortLblDesc).Top = lngSortLbl_Top
37700         .Controls(strSortLblAsc).Left = lngSortLbl_Left
37710         .Controls(strSortLblDesc).Left = lngSortLbl_Left
37720         .Controls(strSortLine1).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
37730         .Controls(strSortLine2).Width = lngTpp
37740         .Controls(strSortLine1).Top = lngSortLine_Top
37750         .Controls(strSortLine2).Top = (lngSortLine_Top + lngTpp)
37760         .Controls(strSortLine1).Left = lngSortLine_Left
37770         .Controls(strSortLine2).Left = lngSortLine_Left
37780         .Controls(strSortLine1).Width = lngSortLine_Width
37790         .Controls(strSortLine2).Width = lngSortLine_Width
37800         If strCalled = "Account_Number" Then
37810           Select Case blnAll
                Case True
37820             strSortAsc = StringReplace(strSortOrig, "_01a_", "_01c_")  ' ** Module Function: modStringFuncs.
37830           Case False
37840             strSortAsc = strSortOrig
37850           End Select
37860           If strSortNow = strSortAsc And blnSortAgain = False Then
37870             Select Case blnAll
                  Case True
37880               strSortNow = StringReplace(strSortOrig, "_01a_", "_01d_")  ' ** Module Function: modStringFuncs.
37890             Case False
37900               strSortNow = "qryPrintChecks_01_01b_Account_Number"
37910             End Select
37920             .Controls(strSortLblAsc).Visible = False
37930             .Controls(strSortLblDesc).Visible = True
37940           Else
37950             strSortNow = strSortAsc
37960             .Controls(strSortLblAsc).Visible = True
37970             .Controls(strSortLblDesc).Visible = False
37980           End If
37990         Else
38000           Select Case strCalled
                Case "Short_Name"
38010             Select Case blnAll
                  Case True
38020               strSortAsc = "qryPrintChecks_01_02c_Short_Name"
38030             Case False
38040               strSortAsc = "qryPrintChecks_01_02a_Short_Name"
38050             End Select
38060           Case "Check_Count"
38070             lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
38080             .Controls(strSortLblAsc).Left = lngSortLbl_Left
38090             .Controls(strSortLblDesc).Left = lngSortLbl_Left
38100             Select Case blnAll
                  Case True
38110               strSortAsc = "qryPrintChecks_01_03c_Check_Count"
38120             Case False
38130               strSortAsc = "qryPrintChecks_01_03a_Check_Count"
38140             End Select
38150           Case "Last_Check_Number"
38160             lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
38170             .Controls(strSortLblAsc).Left = lngSortLbl_Left
38180             .Controls(strSortLblDesc).Left = lngSortLbl_Left
38190             Select Case blnAll
                  Case True
38200               strSortAsc = "qryPrintChecks_01_04c_Last_Check_Number"
38210             Case False
38220               strSortAsc = "qryPrintChecks_01_04a_Last_Check_Number"
38230             End Select
38240           Case "Payee"
38250             Select Case blnAll
                  Case True
38260               strSortAsc = "qryPrintChecks_01_05c_Payee"
38270             Case False
38280               strSortAsc = "qryPrintChecks_01_05a_Payee"
38290             End Select
38300           Case "Bank_Name"
38310             Select Case blnAll
                  Case True
38320               strSortAsc = "qryPrintChecks_01_06c_Bank_Name"
38330             Case False
38340               strSortAsc = "qryPrintChecks_01_06a_Bank_Name"
38350             End Select
38360           Case "Bank_Account_Number"
38370             Select Case blnAll
                  Case True
38380               strSortAsc = "qryPrintChecks_01_07c_Bank_Account_Number"
38390             Case False
38400               strSortAsc = "qryPrintChecks_01_07a_Bank_Account_Number"
38410             End Select
38420           Case "Check_Amount"
38430             Select Case blnAll
                  Case True
38440               strSortAsc = "qryPrintChecks_01_08c_Check_Amount"
38450             Case False
38460               strSortAsc = "qryPrintChecks_01_08a_Check_Amount"
38470             End Select
38480           Case "ChkStat"
38490             Select Case blnAll
                  Case True
38500               strSortAsc = "qryPrintChecks_01_09c_ChkStat"
38510             Case False
38520               strSortAsc = "qryPrintChecks_01_09a_ChkStat"
38530             End Select
38540           End Select
38550           intPos01 = CharPos(strSortAsc, 3, "_")  ' ** Module Function: modStringFuncs.
38560           If strSortNow = strSortAsc And blnSortAgain = False Then
38570             Select Case blnAll
                  Case True
38580               strSortNow = Left(strSortAsc, (intPos01 - 2)) & "d" & Mid(strSortAsc, intPos01)
38590             Case False
38600               strSortNow = Left(strSortAsc, (intPos01 - 2)) & "b" & Mid(strSortAsc, intPos01)
38610             End Select
38620             .Controls(strSortLblAsc).Visible = False
38630             .Controls(strSortLblDesc).Visible = True
38640           Else
38650             strSortNow = strSortAsc
38660             .Controls(strSortLblAsc).Visible = True
38670             .Controls(strSortLblDesc).Visible = False
38680           End If
38690         End If
38700       End If
38710       .Controls(strSortLine1).Visible = True
38720       .Controls(strSortLine2).Visible = True
38730       .lbxShortAccountName.RowSource = strSortNow
38740       If .chkSyncListSort = True Then
38750         If IsNull(.cmbSortBy) = False Then
38760           If Trim(.cmbSortBy) <> vbNullString Then
38770             If .cmbSortBy.Column(CBX_SORT_FNAM) <> strCalled Then
38780               For intX = 0 To (.cmbSortBy.ListCount - 1)
38790                 If .cmbSortBy.Column(CBX_SORT_FNAM, intX) = strCalled Then
38800                   .cmbSortBy = .cmbSortBy.Column(CBX_SORT_FID, intX)
38810                   If strCalled = "Check_Count" Or strCalled = "Last_Check_Number" Then
38820                     Beep
38830                     .cmbSortBy_lbl2.Visible = True
38840                     .cmdPrintChecks.Enabled = False
38850                     .cmdPreviewChecks.Enabled = False
38860                     .cmbSortBy.ForeColor = CLR_DISABLED_FG
38870                     .cmbSortBy.BackColor = CLR_DISABLED_BG
38880                   End If
38890                   Exit For
38900                 End If
38910               Next
38920             End If
38930           End If
38940         End If
38950       End If
38960       If strCalled <> "Check_Count" And strCalled <> "Last_Check_Number" And .cmdPrintChecks.Enabled = False Then
38970         .cmbSortBy_lbl2.Visible = False
38980         .cmdPrintChecks.Enabled = True
38990         .cmdPreviewChecks.Enabled = True
39000         .cmbSortBy.ForeColor = CLR_VDKGRY
39010         .cmbSortBy.BackColor = CLR_WHT
39020       End If
39030     End If
39040     blnSortAgain = False
39050   End With

EXITP:
39060   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

39100 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

39110   With Me
39120     blnGTR_Emblem = False
39130     For lngX = 1& To 24&
39140       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
39150     Next
39160   End With

EXITP:
39170   Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub
