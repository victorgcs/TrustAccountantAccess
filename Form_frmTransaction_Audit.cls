VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTransaction_Audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmTransaction_Audit"

'VGC 10/11/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   Form_Load()
' ##   ckgFlds_cmd_Click()
' ##   CriteriaWiden()
' #######################################

' #########################
' ## Use VBA_RenumErrh().  39830
' #########################

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** Shortcut Alt keys responsive from this form:
' **   AccountNo:        A {cmbAccounts on frmTransaction_Audit_Sub_Criteria}
' **   Acct Number:      B {opgAccountSource_optNumber on frmTransaction_Audit_Sub_Criteria}
' **   Close:            C {cmdClose}
' **   Posting Date:     D {TransDateStart on frmTransaction_Audit_Sub_Criteria}
' **   Excel:            E {cmdExcel}
' **   Form View:        F {opgView_optForm}
' **   Include Archive:  H {chkIncludeArchive}
' **   Asset CUSIP:      I {opgAssetSource_optCUSIP on frmTransaction_Audit_Sub_Criteria}
' **   JournalNo:        J {journalno on frmTransaction_Audit_Sub, frmTransaction_Audit_Sub_ds}
' **   Show Criteria:    K {cmdWidenToCriteria}
' **   Select All:       L {cmdSelectAll}
' **   Acct Name:        M {opgAccountSource_optName on frmTransaction_Audit_Sub_Criteria}
' **   Select None:      N {cmdSelectNone}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Clear All:        R {cmdClearAll}
' **   Assets:           S {cmbAssets on frmTransaction_Audit_Sub_Criteria}
' **   Datasheet View:   T {opgView_optDatasheet}
' **   View Fields:      V {ckgFlds_cmd}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}
' **   Asset Type:       Y {opgAssetSource_optType on frmTransaction_Audit_Sub_Criteria}
' **   Size/Position:    Z {chkSaveSizePos}
' **   JournalType1:     1 {cmbJournalType1 on frmTransaction_Audit_Sub_Criteria}
' **   JournalType2:     2 {cmbJournalType2 on frmTransaction_Audit_Sub_Criteria}
' **   JournalType3:     3 {cmbJournalType3 on frmTransaction_Audit_Sub_Criteria}
'G  Q  U

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker Posting - From:  F5  {cmdCalendar1 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Posting - To:    F6  {cmdCalendar2 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Trade - From:    F7  {cmdCalendar3 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Trade - To:      F8  {cmdCalendar4 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Purchase - From: F9  {cmdCalendar5 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Purchase - To:   F10 {cmdCalendar6 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Posted - From:   F11 {cmdCalendar7 on frmTransaction_Audit_Sub_Criteria}
' **   Date Picker Posted - To:     F12 {cmdCalendar8 on frmTransaction_Audit_Sub_Criteria}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Include Hidden:   C {opgHidden_optInclude on frmTransaction_Audit_Sub_Criteria}
' **   Tax Deduction:    D {chkTaxcodeType_Deduction on frmTransaction_Audit_Sub_Criteria}
' **   Exclude Hidden:   E {opgHidden_optExclude on frmTransaction_Audit_Sub_Criteria}
' **   Tax Income:       I {chkTaxcodeType_Income on frmTransaction_Audit_Sub_Criteria}
' **   Asset Name:       M {opgAssetSource_optName on frmTransaction_Audit_Sub_Criteria}
' **   Narrow Font:      N {chkNarrowFont}
' **   Only Hidden:      O {opgHidden_optOnly on frmTransaction_Audit_Sub_Criteria}
' **   Shareface Total:  T {chkIncludeSharefaceTot}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   Inc/Exp Expense:  E {chkRevcodeType_Expense on frmTransaction_Audit_Sub_Criteria}
' **   Inc/Exp Income:   I {chkRevcodeType_Income on frmTransaction_Audit_Sub_Criteria}

'THE DATASHEET MAY NEED ITS VISIBLE FIELD NAMES INSTEAD OF THE TABLE FIELD NAMES IN THE FILTER!
'Journal Number
'Journal Type
'Posting Date
'Account Number
'Name
'CUSIP
'Asset
'Share / Face
'Income Cash
'Principal Cash
'Cost
'Currency
'Trade Date
'Original Trade Date
'Comments
'Recurring Item
'Inc/Exp Codes
'Inc/Exp
'Tax Codes
'Inc/Ded
'Location
'Check Number
'User
'Date Posted
'Hidden

' ** If the filter's bad, the error generated is:
' **   Case 2448  ' ** You can't assign a value to this object.

'####################
' ** Add a button to do additional checks:
' **   1. Future date: qryLedger_05
' **   2. Liability mismatch.
' ** Check for Liability mismatches.
' ** qryLedger_04a (Liability assets), where journaltype = 'Deposit', 'Purchase', 'Sold', 'Withdrawn';
' ** these would be forbidden transactions on a Liability asset; 0.
'Set qdf = dbs.QueryDefs("qryLedger_04b")
' ** qryLedger_02d (Liability journaltypes), not in qryLedger_03 (Liability assets);
' ** these would be forbidden Liability transactions on a non-Liability asset; 0!
'Set qdf = dbs.QueryDefs("qryLedger_04c")
'MsgBox "WARNING!" & vbcrlf & vbcrlf & "It has been determined that you have 'Liability' " & _
'  "transactions associated with non-Liability assets," & vbcrlf & _
'  "or non-Liability transactions associated with Liability assets." & vbcrlf & vbcrlf & _
'  "{some action}", _
'  vbExclamation + vbYesNo, "Liability Error"
'####################

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Filter constants: Both subforms.  ALL IN CRITERIA
'Private Const ANDF         As String = " And "  ' ** Filter 'And'.
'Private Const ORF          As String = " Or "  ' ** Filter 'Or'.
'Private Const JRNL_NUM     As String = "[journalno] = "
'Private Const JRNL_TYPE    As String = "[journaltype] = '"
'Private Const TRANS_START  As String = "[transdate] >= #"
'Private Const TRANS_END    As String = "[transdate] <= #"
'Private Const ACCT_NUM     As String = "[accountno] = '"
'Private Const ASSET_NUM    As String = "[assetno] = "
'Private Const CURR_NUM     As String = "[curr_id] = "
'Private Const ASSET_START  As String = "[assetdate] >= #"
'Private Const ASSET_END    As String = "[assetdate] <= #"
'Private Const PURCH_START  As String = "[PurchaseDate] >= #"
'Private Const PURCH_END    As String = "[PurchaseDate] <= #"
'Private Const COMM_DESC    As String = "[ledger_description] Like '*"
'Private Const RECUR_ITEM   As String = "[RecurringItem] Like '*"
'Private Const REV_CODE     As String = "[revcode_ID] = "
'Private Const TAX_CODE     As String = "[taxcode] = "
'Private Const LOC_NUM      As String = "[Location_ID] = "
'Private Const CHK_NUM      As String = "[CheckNum] = "
'Private Const CHK_NUM1     As String = "[CheckNum] >= "
'Private Const CHK_NUM2     As String = "[CheckNum] <= "
'Private Const JRNL_USER    As String = "[journal_USER] = '"
'Private Const POSTED_START As String = "[posted] >= #"
'Private Const POSTED_END   As String = "[posted] <= #"
'Private Const HIDDEN_TRX1  As String = "[ledger_HIDDEN] = True"
'Private Const HIDDEN_TRX2  As String = "[ledger_HIDDEN] = False"

' ** Array: arr_varFilt(), arr_varFilt_ds().
Private lngFilts As Long, arr_varFilt As Variant
Private lngFilts_ds As Long, arr_varFilt_ds As Variant
'Private Const F_ELEMS As Integer = 13  ' ** Array's first-element UBound().
'Private Const F_IDX   As Integer = 0
'Private Const F_NAM   As Integer = 1
'Private Const F_CONST As Integer = 2
Private Const F_CTL   As Integer = 3
Private Const F_CLBL  As Integer = 4
'Private Const F_FLD   As Integer = 5
'Private Const F_FLBL  As Integer = 6
'Private Const F_CTL2  As Integer = 7
'Private Const F_CLBL2 As Integer = 8
'Private Const F_FLD2  As Integer = 9
'Private Const F_FLBL2 As Integer = 10
Private Const F_CLBL3 As Integer = 11
'Private Const F_FLD3  As Integer = 12
'Private Const F_FLBL3 As Integer = 13

' ** Array: arr_varFrmFld(), arr_varFrmFld_ds().
Private lngFrmFlds As Long, arr_varFrmFld() As Variant
Private lngFrmFlds_ds As Long, arr_varFrmFld_ds() As Variant
Private Const FM_ELEMS As Integer = 7  ' ** Array's first-element UBound().
Private Const FM_FLD_NAM As Integer = 0
Private Const FM_FLD_TAB As Integer = 1
Private Const FM_FLD_VIS As Integer = 2
Private Const FM_CHK_NAM As Integer = 3
Private Const FM_CHK_VAL As Integer = 4
Private Const FM_VIEWCHK As Integer = 5
Private Const FM_TOPO    As Integer = 6
Private Const FM_TOPC    As Integer = 7

Private rstAll1 As DAO.Recordset, rstAll2 As DAO.Recordset

Private blnClearAll_Focus As Boolean, blnClearAll_MouseDown As Boolean
Private blnSelectAll_Focus As Boolean, blnSelectAll_MouseDown As Boolean
Private blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean
Private blnCkgFlds_Focus As Boolean, blnCkgFlds_MouseDown As Boolean
Private blnWidenToCrit_Focus As Boolean, blnWidenToCrit_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngChkBoxLbl_LeftOffset As Long, lngChkBoxLbl_TopOffset As Long, lngCkgFldsHLine04_Top As Long
Private lngCkgFldsBox_Height As Long, lngCkgFldsBox_Offset As Long, lngCkgFldsVLine_Top As Long, lngCkgFldsVLine_Height As Long
Private lngForm_Width As Long, lngSub_Width As Long, lngSub_Height As Long, lngDetail_Height As Long
Private lngForm_Width_New As Long, lngDetail_Height_New As Long, lngDetail_Height_Diff As Long, lngDetail_Height_Min As Long
Private lngSub_Width_New As Long, lngSub_Width_Diff As Long, lngSub_Width_Min As Long, lngSub_Height_Offset As Long
Private lngSizable_Offset As Long, lngHLine_Offset As Long, lngCritSub_Width As Long, lngCritSub_Max As Long
Private lngClose_Left As Long, lngClose_Offset As Long
Private lngNarrowFontBox_Left As Long, lngNarrowFont_Left As Long, lngPageOf_Left As Long

Private lngSelectBox_Top As Long, lngSelectLegendLbl_Top As Long, lngSelectAll_Top As Long, lngSelectNone_Top As Long
Private lngNarrowFont_Top As Long, lngShareface_Top As Long, lngSelectLbl_Top As Long, lngSelectLbl2_Top As Long

Private frmCrit As Access.Form, frmMaint As Access.Form
Private blnIsMaximized As Boolean, blnNoArchive As Boolean, blnFromCkgFlds As Boolean, blnProgBarOff As Boolean
Private blnRecsZero As Boolean, intRecsZero As Integer, blnScrollBar As Boolean
Private lngTotalRecs1 As Long, lngTotalRecs2 As Long, lngChkCnt As Long, lngVisCnt1 As Long, lngVisCnt2 As Long
Private strFilter01 As String, strFilter02 As String, dblFilterRecs As Double
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim qdfAll As DAO.QueryDef
        Dim strFilter As String, strQryName As String
        Dim Button As Integer, Shift As Integer, X As Single, Y As Single
        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim blnRetVal As Boolean, blnMsgSeen As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           DoCmd.Hourglass True  ' ** Make sure it's still running.
170           DoEvents

180           CLR_DISABLED_FG = CLR_DKGRY
190           CLR_DISABLED_BG = CLR_LTTEAL

200           gblnMessage = True: gblnClosing = False: blnRecsZero = False: intRecsZero = 0
210           gblnSetFocus = False: blnProgBarOff = False
220           blnMsgSeen = False: blnFromCkgFlds = False: blnScrollBar = False
230           blnClearAll_Focus = False: blnClearAll_MouseDown = False
240           blnSelectAll_Focus = False: blnSelectAll_MouseDown = False
250           blnSelectNone_Focus = False: blnSelectNone_MouseDown = False
260           blnCkgFlds_Focus = False: blnCkgFlds_MouseDown = False
270           blnWidenToCrit_Focus = False: blnWidenToCrit_MouseDown = False
280           DoEvents

290           .ckgFlds_cmd_raised_img.Visible = False
300           .ckgFlds_cmd_raised_semifocus_dots_img.Visible = False
310           .ckgFlds_cmd_raised_focus_img.Visible = False
320           .ckgFlds_cmd_raised_focus_dots_img.Visible = False
330           .ckgFlds_cmd_sunken_focus_dots_img.Visible = False
340           .ckgFlds_cmd_raised_img_dis.Visible = False

350           Set frmMaint = Forms(strCallingForm)

              ' ****************************************
360           gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
370           frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
380           gctlPBar_Lbl.Caption = "Step 7."
390           DoEvents
400           ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
410           gctlPBar_Lbl_2.Caption = "Step 7."
420           DoEvents
              ' ****************************************

              ' ***************************************
              ' ** Steps 5, 6, 7, & 8 are in subform.
              ' ***************************************

              ' ****************************************
              ' ** Step 9. Query 4
              ' ****************************************
430           gdblPBar_ThisStep = 9#
              ' ****************************************
440           strQryName = "qryTransaction_Audit_01"  ' ** frmTransaction_Audit_Sub RecordSource.
450           Set qdfAll = CurrentDb.QueryDefs(strQryName)
460           Set rstAll1 = qdfAll.OpenRecordset
470           If rstAll1.BOF = True And rstAll1.EOF = True Then
                ' ** New installation, no transactions!
480             blnMsgSeen = True
490             Cancel = -1
500             MsgBox "There are no transactions to view.", vbInformation + vbOKOnly, "Nothing To Do"
510             gblnSetFocus = True
                ' ****************************************
520             gdblPBar_CurWidth = gdblPBar_MaxWidth
530             frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
540             gctlPBar_Lbl.Caption = "Step 9."
550             DoEvents
560             ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
570             gctlPBar_Lbl_2.Caption = "Step 9."
580             DoEvents
                ' ****************************************
590             Forms(strCallingForm).TimerInterval = 500&
600           Else

610             .frmTransaction_Audit_Sub_Criteria.Form.RstAllSet strQryName  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.

620             SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
630             DoEvents

640             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
650             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

                ' ****************************************
660             gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
670             frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
680             gctlPBar_Lbl.Caption = "Step 9."
690             DoEvents
700             ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
710             gctlPBar_Lbl_2.Caption = "Step 9."
720             DoEvents
                ' ****************************************

                ' ****************************************
                ' ** Step 10. Form initialization
                ' ****************************************
730             gdblPBar_ThisStep = 10#
                ' ****************************************

                ' ** Variables are fed empty, then populated ByRef.
740             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

750             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

760             lngSub_Width = .frmTransaction_Audit_Sub.Width
770             lngSub_Height = .frmTransaction_Audit_Sub.Height
780             lngForm_Width = .Width
790             lngDetail_Height = .Detail.Height

800             .form_width_line.Left = 0&  ' ** Just get these out of the way.
810             .detail_height_line.Top = 0&
820             .cmdSelectAll_lbl.Top = 0&
830             .cmdSelectAll_lbl.Left = 0&
840             .cmdSelectNone_lbl.Top = 0&
850             .cmdSelectNone_lbl.Left = 0&
860             .cmdSelect_Legend_tgl_on_raised_img_lbl.Top = 0&
870             .cmdSelect_Legend_tgl_on_raised_img_lbl.Left = 0&
880             .cmdSelect_Legend_tgl_off_raised_img_lbl.Top = 0&
890             .cmdSelect_Legend_tgl_off_raised_img_lbl.Left = 0&
900             .cmdSelect_Legend_lbl_alt_line1.Top = 0&  ' ** Their Left positions are still used.
910             .cmdSelect_Legend_lbl_alt_line2.Top = 0&
920             .cmdSelect_Legend_lbl_alt_line3.Top = 0&
930             .cmdSelect_Legend_lbl_alt_line4.Top = 0&
940             .cmdSelect_Legend_lbl_alt_line5.Top = 0&
950             .cmdSelect_Legend_lbl_alt_line6.Top = 0&
960             .chkNarrowFont_box_alt_line.Top = 0&
970             .ckgFlds_cmd_lbl.Top = 0&
980             .ckgFlds.Top = 0&
990             DoEvents

1000            lngSizable_Offset = (.Header_vline01.Left - .Sizable_lbl1.Left)
1010            lngHLine_Offset = (.Header_hline01.Width - .frmTransaction_Audit_Sub.Width)
1020            lngClose_Left = .cmdClose.Left
1030            lngSub_Height_Offset = (.Detail.Height - .frmTransaction_Audit_Sub.Height)
1040            lngDetail_Height_Min = ((110& * lngTpp) + lngSub_Height_Offset)
1050            lngClose_Offset = ((.frmTransaction_Audit_Sub.Left + .frmTransaction_Audit_Sub.Width) - .cmdClose.Left)
1060            lngNarrowFontBox_Left = .chkNarrowFont_box.Left
1070            lngNarrowFont_Top = .chkNarrowFont.Top
1080            lngNarrowFont_Left = .chkNarrowFont.Left
1090            lngShareface_Top = .chkIncludeSharefaceTot.Top
1100            lngPageOf_Left = .chkPageOf.Left
1110            lngChkBoxLbl_LeftOffset = (.chkNarrowFont_lbl.Left - .chkNarrowFont.Left)
1120            lngChkBoxLbl_TopOffset = (.ckgFlds_chkAccountNo.Top - .ckgFlds_chkAccountNo_lbl.Top)
1130            lngSelectBox_Top = .cmdSelect_box.Top
1140            lngSelectAll_Top = .cmdSelectAll.Top
1150            lngSelectNone_Top = .cmdSelectNone.Top
1160            lngSelectLbl_Top = .cmdSelect_lbl.Top
1170            lngSelectLegendLbl_Top = .cmdSelect_Legend_lbl.Top
1180            lngSelectLbl2_Top = .cmdSelect_lbl2.Top
1190            lngCkgFldsBox_Height = .ckgFlds_box.Height
1200            lngCkgFldsVLine_Top = .ckgFlds_vline09.Top
1210            lngCkgFldsVLine_Height = .ckgFlds_vline09.Height
1220            lngCkgFldsBox_Offset = .Detail.Height - (.ckgFlds_box.Top + .ckgFlds_box.Height)
1230            lngCkgFldsHLine04_Top = .ckgFlds_hline04.Top
1240            lngCritSub_Width = .frmTransaction_Audit_Sub_Criteria.Width
1250            lngCritSub_Max = .frmTransaction_Audit_Sub_Criteria.Form.Width

                ' ****************************************
1260            gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
1270            frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
1280            gctlPBar_Lbl.Caption = "Step 10."
1290            DoEvents
1300            ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
1310            gctlPBar_Lbl_2.Caption = "Step 10."
1320            DoEvents
                ' ****************************************

1330            DoCmd.Hourglass True  ' ** Make sure it's still running.
1340            DoEvents

                ' ****************************************
                ' ** Step 11. Load filter
                ' ****************************************
1350            gdblPBar_ThisStep = 11#
                ' ****************************************

1360            FilterRecs_Load  ' ** Module Procedure: modTransactionAuditFuncs1.
1370            DoEvents

1380            arr_varFilt = Empty
1390            arr_varFilt = FilterRecs_GetArr(1)  ' ** Module Function: modTransactionAuditFuncs1.
1400            lngFilts = (UBound(arr_varFilt, 2) + 1&)
1410            arr_varFilt_ds = Empty
1420            arr_varFilt_ds = FilterRecs_GetArr(2)  ' ** Module Function: modTransactionAuditFuncs1.
1430            lngFilts_ds = (UBound(arr_varFilt_ds, 2) + 1&)

1440            strFilter01 = vbNullString: strFilter02 = vbNullString
1450            dblFilterRecs = 0#
1460            .CurrentFilter1 = vbNullString
1470            .CurrentFilter2 = vbNullString

                ' ****************************************
1480            gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
1490            frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
1500            gctlPBar_Lbl.Caption = "Step 11."
1510            DoEvents
1520            ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
1530            gctlPBar_Lbl_2.Caption = "Step 11."
1540            DoEvents
                ' ****************************************

                ' ****************************************
                ' ** Step 12. Set totals
                ' ****************************************
1550            gdblPBar_ThisStep = 12#
                ' ****************************************

1560            rstAll1.MoveLast
1570            lngTotalRecs1 = rstAll1.RecordCount
1580            .frmTransaction_Audit_Sub.Form.TotalRecs_Set lngTotalRecs1  ' ** Form Procedure: frmTransaction_Audit_Sub.
1590            .frmTransaction_Audit_Sub_ds.Form.TotalRecs_Set lngTotalRecs1  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
1600            rstAll1.MoveFirst
1610            DoEvents

1620            DoCmd.Hourglass True  ' ** Make sure it's still running.
1630            DoEvents

      #If NoExcel Then
1640            .cmdExcel.Enabled = False
      #Else
1650            .cmdExcel.Enabled = True
      #End If

                ' ****************************************
1660            gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
1670            frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
1680            gctlPBar_Lbl.Caption = "Step 12."
1690            DoEvents
1700            ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
1710            gctlPBar_Lbl_2.Caption = "Step 12."
1720            DoEvents
                ' ****************************************

                ' ****************************************
                ' ** Step 13. Set default options
                ' ****************************************
1730            gdblPBar_ThisStep = 13#
                ' ****************************************

                ' ** Form_Load() takes over during this section,
                ' ** with Steps 14-22.

1740            DoCmd.Hourglass True  ' ** Make sure it's still running.
1750            DoEvents

1760            Set frmCrit = .frmTransaction_Audit_Sub_Criteria.Form

                ' ** These are defaults, which can be overridden by preferences loaded in Form_Load(), below.
1770            With frmCrit
                  'LET THE 2ND CALL TAKE CARE OF THESE!
1780              .opgAccountSource = .opgAccountSource_optNumber.OptionValue
1790              .opgAssetSource = .opgAssetSource_optName.OptionValue
1800              .chkRevcodeType_Expense = True
1810              .chkRevcodeType_Income = True
1820              .chkTaxcodeType_Income = True
1830              .chkTaxcodeType_Deduction = True
1840              .opgHidden = .opgHidden_optInclude.OptionValue
1850              .opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
1860              DoEvents
1870              .cmbJournalType2.Enabled = False
1880              .cmbJournalType2.BorderColor = WIN_CLR_DISR
1890              .cmbJournalType2.BackStyle = acBackStyleTransparent
1900              .cmbJournalType3.Enabled = False
1910              .cmbJournalType3.BorderColor = WIN_CLR_DISR
1920              .cmbJournalType3.BackStyle = acBackStyleTransparent
1930            End With
1940            .ckgFlds_chkCurrID = False
1950            .ckgFlds_chkLocationName = False
1960            .ckgFlds_chkCheckNum = False
1970            chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
1980            DoEvents

1990            DoCmd.Hourglass True  ' ** Make sure it's still running.
2000            DoEvents

                ' ****************************************
2010            gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
2020            frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
2030            gctlPBar_Lbl.Caption = "Step 22."
2040            DoEvents
2050            ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
2060            gctlPBar_Lbl_2.Caption = "Step 22."
2070            DoEvents
                ' ****************************************

                ' ****************************************
                ' ** Step 23. Initialize field check boxes
                ' ****************************************
2080            gdblPBar_ThisStep = 23#
                ' ****************************************

2090            DoCmd.Hourglass True  ' ** Make sure it's still running.
2100            DoEvents

                ' *********************************************************************************************
                ' ** The arr_varFld() array absolutely, positively has to be loaded by the time it gets here!
                ' *********************************************************************************************
2110            .frmTransaction_Audit_Sub.Form.ShowFields_Sub THIS_PROC, 1  ' ** Form Procedure: frmTransaction_Audit_Sub.
2120            DoEvents

2130            ckgFlds_chkJournalno_AfterUpdate  ' ** Procedure: Below.
2140            ckgFlds_chkJournalType_AfterUpdate  ' ** Procedure: Below.
2150            ckgFlds_chkTransDate_AfterUpdate  ' ** Procedure: Below.
2160            ckgFlds_chkAccountNo_AfterUpdate  ' ** Procedure: Below.
2170            ckgFlds_chkShortName_AfterUpdate  ' ** Procedure: Below.
2180            ckgFlds_chkCusip_AfterUpdate  ' ** Procedure: Below.
2190            ckgFlds_chkAssetDescription_AfterUpdate  ' ** Procedure: Below.
2200            ckgFlds_chkShareFace_AfterUpdate  ' ** Procedure: Below.
2210            ckgFlds_chkICash_AfterUpdate  ' ** Procedure: Below.
2220            ckgFlds_chkPCash_AfterUpdate  ' ** Procedure: Below.
2230            ckgFlds_chkCost_AfterUpdate  ' ** Procedure: Below.
2240            ckgFlds_chkCurrID_AfterUpdate  ' ** Procedure: Below.
2250            ckgFlds_chkAssetDate_AfterUpdate  ' ** Procedure: Below.
2260            ckgFlds_chkPurchaseDate_AfterUpdate  ' ** Procedure: Below.
2270            ckgFlds_chkLedgerDescription_AfterUpdate  ' ** Procedure: Below.
2280            ckgFlds_chkRecurringItem_AfterUpdate  ' ** Procedure: Below.
2290            ckgFlds_chkRevCodeDesc_AfterUpdate  ' ** Procedure: Below.
2300            ckgFlds_chkRevCodeTypeDescription_AfterUpdate  ' ** Procedure: Below.
2310            ckgFlds_chkTaxCodeDescription_AfterUpdate  ' ** Procedure: Below.
2320            ckgFlds_chkTaxCodeTypeDescription_AfterUpdate  ' ** Procedure: Below.
2330            ckgFlds_chkLocationName_AfterUpdate  ' ** Procedure: Below.
2340            ckgFlds_chkCheckNum_AfterUpdate  ' ** Procedure: Below.
2350            ckgFlds_chkJournalUser_AfterUpdate  ' ** Procedure: Below.
2360            ckgFlds_chkPosted_AfterUpdate  ' ** Procedure: Below.
2370            ckgFlds_chkLedgerHidden_AfterUpdate  ' ** Procedure: Below.
2380            DoEvents

2390            .frmTransaction_Audit_Sub.Form.Detail_MouseMove Button, Shift, X, Y  ' ** Form Procedure: frmTransaction_Audit_Sub.

2400            If .ckgFlds = False Then  ' ** This should be after Pref_Load().
                  ' ** The button only closes it if it's True.
2410              .ckgFlds = True
2420              ckgFlds_cmd_Click  ' ** Procedure: Below.
2430              DoEvents
2440            End If

2450            DoCmd.Hourglass True  ' ** Make sure it's still running.
2460            DoEvents

                ' ****************************************
2470            gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
2480            frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
2490            gctlPBar_Lbl.Caption = "Step 23."
2500            DoEvents
2510            ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
2520            gctlPBar_Lbl_2.Caption = "Step 23."
2530            DoEvents
                ' ****************************************

                ' ****************************************
                ' ** Step 24. Set filters again
                ' ****************************************
2540            gdblPBar_ThisStep = 24#
                ' ****************************************

2550            blnIsMaximized = IsMaximized(Me)  ' ** Module Function: modWindowFunctions.

2560            strFilter = vbNullString
2570            If IsNull(.CurrentFilter1) = False Then
2580              If Trim(.CurrentFilter1) <> vbNullString Then
2590                If .frmTransaction_Audit_Sub.Form.Filter <> .CurrentFilter1 Then
2600  On Error Resume Next
2610                  .frmTransaction_Audit_Sub.Form.Filter = .CurrentFilter1
2620  On Error GoTo ERRH
2630                  DoEvents
2640                End If
2650                .frmTransaction_Audit_Sub.Form.FilterOn = True
2660                DoEvents
2670                If IsNull(.CurrentFilter2) = False Then
2680                  If Trim(.CurrentFilter2) <> vbNullString Then
2690                    .CurrentFilter2 = strFilter02
2700                    strFilter02 = .frmTransaction_Audit_Sub_ds.Form.SwapNames(strFilter02)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
2710                    If .frmTransaction_Audit_Sub_ds.Form.Filter <> strFilter02 Then
2720  On Error Resume Next
2730                      .frmTransaction_Audit_Sub_ds.Form.Filter = strFilter02
2740  On Error GoTo ERRH
2750                    End If
2760                    .frmTransaction_Audit_Sub_ds.Form.FilterOn = True
2770                    .frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
2780                  Else
2790                    If strFilter02 = vbNullString Then
2800                      strFilter02 = .frmTransaction_Audit_Sub_ds.Form.SwapNames(strFilter01)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
2810                    Else
2820                      strFilter02 = .frmTransaction_Audit_Sub_ds.Form.SwapNames(strFilter02)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
2830                    End If
2840                    .CurrentFilter2 = strFilter02
2850                    .frmTransaction_Audit_Sub_ds.Form.Filter = strFilter02
2860                    .frmTransaction_Audit_Sub_ds.Form.FilterOn = True
2870                    .frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
2880                  End If
2890                Else
2900                  If strFilter02 = vbNullString Then
2910                    strFilter02 = .frmTransaction_Audit_Sub_ds.Form.SwapNames(strFilter01)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
2920                  Else
2930                    strFilter02 = .frmTransaction_Audit_Sub_ds.Form.SwapNames(strFilter02)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
2940                  End If
2950                  .CurrentFilter2 = strFilter02
2960                  .frmTransaction_Audit_Sub_ds.Form.Filter = strFilter02
2970                  .frmTransaction_Audit_Sub_ds.Form.FilterOn = True
2980                  .frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
2990                End If
3000              Else
3010                .frmTransaction_Audit_Sub.Form.Filter = vbNullString
3020                .frmTransaction_Audit_Sub.Form.FilterOn = False
3030                .frmTransaction_Audit_Sub_ds.Form.Filter = vbNullString
3040                .frmTransaction_Audit_Sub_ds.Form.FilterOn = False
3050              End If
3060            Else
3070              .frmTransaction_Audit_Sub.Form.Filter = vbNullString
3080              .frmTransaction_Audit_Sub.Form.FilterOn = False
3090              .frmTransaction_Audit_Sub_ds.Form.Filter = vbNullString
3100              .frmTransaction_Audit_Sub_ds.Form.FilterOn = False
3110            End If
3120            DoEvents

3130            DoCmd.Hourglass True  ' ** Make sure it's still running.
3140            DoEvents

3150            If frmCrit.journalno.Enabled = True Then
3160              .frmTransaction_Audit_Sub_Criteria.SetFocus
3170              frmCrit.FocusHolder.SetFocus
3180              If frmCrit.journalno.Enabled = True Then
3190                frmCrit.journalno.SetFocus
3200              End If
3210            End If

                ' ****************************************
3220            gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
3230            frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
3240            gctlPBar_Lbl.Caption = "Step 24."
3250            DoEvents
3260            ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
3270            gctlPBar_Lbl_2.Caption = "Step 24."
3280            DoEvents
                ' ****************************************

                ' ****************************************
                ' ** Step 25. Finish opening
                ' ****************************************
3290            gdblPBar_ThisStep = 25#
                ' ****************************************

3300            If .opgView = .opgView_optForm.OptionValue Then
3310              .frmTransaction_Audit_Sub.SetFocus
3320              .frmTransaction_Audit_Sub.Form.FocusHolder.SetFocus
3330              If .frmTransaction_Audit_Sub.Form.journalno.Visible = True And .frmTransaction_Audit_Sub.Form.journalno.Enabled = True Then
3340                .frmTransaction_Audit_Sub.Form.journalno.SetFocus
3350              End If
3360            End If

3370            DoCmd.SelectObject acForm, THIS_NAME, False
3380            If frmCrit.journalno.Enabled = True Then
3390              .frmTransaction_Audit_Sub_Criteria.SetFocus
3400              frmCrit.FocusHolder.SetFocus
3410              If frmCrit.journalno.Enabled = True Then
3420                frmCrit.journalno.SetFocus
3430              End If
3440            End If
3450            frmCrit.FocusHolder.SetFocus
3460            fSetScrollBarPosHZ frmCrit, 1&  ' ** Module Function: modScrollBarFuncs.

3470            If .ckgFlds = 0 And .chkSaveSizePos = False Then
                  ' ** Make sure window size is correct.
3480              lngTmp01 = ((.ckgFlds_box.Top + lngCkgFldsBox_Height) + lngCkgFldsBox_Offset)
3490              lngTmp01 = (lngTmp01 - (lngCkgFldsHLine04_Top + lngCkgFldsBox_Offset))
3500              lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
3510              lngMonitorNum = 1&: lngTmp02 = 0&
3520              EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
3530              If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
3540              If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
3550              DoCmd.SelectObject acForm, "frmTransaction_Audit", False
3560              DoEvents
3570              DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, (lngFrm_Height - lngTmp01)  'lngFrm_Top
3580              If lngMonitorNum > 1& Then
3590                LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
3600              End If
3610            End If

                ' ****************************************
3620            gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
3630            frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
3640            gctlPBar_Lbl.Caption = "Step 25."
3650            DoEvents
3660            ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
3670            gctlPBar_Lbl_2.Caption = "Step 25."
3680            DoEvents
3690            blnProgBarOff = True
3700            .TimerInterval = 2000&
                ' ****************************************

3710            DoEvents
3720            blnScrollBar = True
                'HAVE OTHER TIMER EVENTS FINISHED BY THIS TIME?
3730            .TimerInterval = 100&

3740            DoCmd.Hourglass False
3750            DoEvents

3760          End If  ' ** BOF, EOF.

              'ckgFlds_chkJournalno_AfterUpdate
              'ckgFlds_chkJournalType_AfterUpdate
              'ckgFlds_chkTransDate_AfterUpdate
              'ckgFlds_chkAccountNo_AfterUpdate
              'ckgFlds_chkShortName_AfterUpdate
              'ckgFlds_chkCusip_AfterUpdate
              'ckgFlds_chkAssetDescription_AfterUpdate
              'ckgFlds_chkShareFace_AfterUpdate
              'ckgFlds_chkICash_AfterUpdate
              'ckgFlds_chkPCash_AfterUpdate
              'ckgFlds_chkCost_AfterUpdate
              'ckgFlds_chkCurrID_AfterUpdate
              'ckgFlds_chkAssetDate_AfterUpdate
              'ckgFlds_chkPurchaseDate_AfterUpdate
              'ckgFlds_chkLedgerDescription_AfterUpdate
              'ckgFlds_chkRecurringItem_AfterUpdate
              'ckgFlds_chkRevCodeDesc_AfterUpdate
              'ckgFlds_chkRevCodeTypeDescription_AfterUpdate
              'ckgFlds_chkTaxCodeDescription_AfterUpdate
              'ckgFlds_chkTaxCodeTypeDescription_AfterUpdate
              'ckgFlds_chkLocationName_AfterUpdate
              'ckgFlds_chkCheckNum_AfterUpdate
              'ckgFlds_chkJournalUser_AfterUpdate
              'ckgFlds_chkPosted_AfterUpdate
              'ckgFlds_chkLedgerHidden_AfterUpdate

3770        Else
3780          Cancel = -1
3790        End If
3800      Else
3810        Cancel = -1
3820      End If
3830    End With

3840    If Cancel = -1 And blnMsgSeen = False Then
3850      DoCmd.Hourglass False
3860      Beep
3870      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
3880    End If

EXITP:
3890    Set qdfAll = Nothing
3900    Exit Sub

ERRH:
3910    DoCmd.Hourglass False
3920    Select Case ERR.Number
        Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub Form_Load()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, ctls As Object, ctl As Access.Control
        Dim frm1 As Access.Form, frm2 As Access.Form                   ' ** Code errors with 'Type mismatch' if this is set to Access.Controls!
        Dim strRptPath As String
        Dim lngRecs As Long
        Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long
        Dim lngX As Long, lngY As Long, lngE As Long

4010    With Me

4020      DoCmd.Hourglass True  ' ** Make sure it's still running.
4030      DoEvents

4040      Set frm1 = .frmTransaction_Audit_Sub.Form
4050      Set frm2 = .frmTransaction_Audit_Sub_ds.Form
4060      Set frmCrit = .frmTransaction_Audit_Sub_Criteria.Form

          ' ****************************************
4070      gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
4080      frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
4090      gctlPBar_Lbl.Caption = "Step 12."
4100      DoEvents
4110      ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
4120      gctlPBar_Lbl_2.Caption = "Step 12."
4130      DoEvents
          ' ****************************************

          ' ****************************************
          ' ** Step 14. Check LedgerArchive
          ' ****************************************
4140      gdblPBar_ThisStep = 14#
          ' ****************************************

4150      varTmp00 = DCount("*", "LedgerArchive")

4160      If IsNull(varTmp00) = False Then
4170        If varTmp00 > 0 Then
4180          blnNoArchive = False
4190          .chkIncludeArchive.Enabled = True
              ' ** The chkIncludeArchive_lbl.Caption is empty,
              ' ** it's chkIncludeArchive_lbl2.Caption that has "Include Archive".
4200          .chkIncludeArchive_lbl2.Visible = True
4210        Else
4220          blnNoArchive = True
4230        End If
4240      Else
4250        blnNoArchive = True
4260      End If
4270      DoEvents
4280      If blnNoArchive = True Then
4290        .chkIncludeArchive.Enabled = False
4300        .chkIncludeArchive_lbl2.Visible = False
4310        .chkIncludeArchive_lbl3.Visible = True
4320        .chkIncludeArchive_lbl3_dim_hi.Visible = True
4330        .chkIncludeArchive_lbl4.Visible = True
4340        .chkIncludeArchive_lbl4_dim_hi.Visible = True
4350      End If
4360      DoEvents

          ' ****************************************
4370      gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
4380      frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
4390      gctlPBar_Lbl.Caption = "Step 14."
4400      DoEvents
4410      ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
4420      gctlPBar_Lbl_2.Caption = "Step 14."
4430      DoEvents
          ' ****************************************

          ' ****************************************
          ' ** Step 15. Load form data fields
          ' ****************************************
4440      gdblPBar_ThisStep = 15#
          ' ****************************************

4450      lngFrmFlds = 0&
4460      ReDim arr_varFrmFld(FM_ELEMS, 0)

4470      lngFrmFlds_ds = 0&
4480      ReDim arr_varFrmFld_ds(FM_ELEMS, 0)

          ' ***************************************************
          ' ** Array: arr_varFrmFld(), arr_varFrmFld_ds()
          ' **
          ' **   Element  Name                     Constant
          ' **   =======  =======================  ==========
          ' **      0     Field Name               FM_FLD_NAM
          ' **      1     Control Tab Index        FM_FLD_TAB
          ' **      2     Field Visible            FM_FLD_VIS
          ' **      3     Checkbox Name            FM_CHK_NAM
          ' **      4     Checkbox Value           FM_CHK_VAL
          ' **      5     ckgFld_chk.. Name        FM_VIEWCHK
          ' **      6     Check Top When Open      FM_TOPO
          ' **      7     Check Top When Closed    FM_TOPC
          ' **
          ' ***************************************************

          ' ** Get a list of the form data fields.
4490      For lngX = 1& To 2&

4500        Select Case lngX
            Case 1&
4510          Set ctls = frm1.Section("Detail").Controls
4520        Case 2&
4530          Set ctls = frm2.Section("Detail").Controls
4540        End Select
4550        For Each ctl In ctls
4560          With ctl
4570            Select Case .Name
                Case "FocusHolder", "FocusHolder2", "assetno", "revcode_ID", "revcode_TYPE", "taxcode", _
                    "taxcode_type", "curr_code", "Location_ID", "alphasort", "currsym_symbol", "icash_str", _
                    "pcash_str", "cost_str", "icash_str_bg", "pcash_str_bg", "cost_str_bg", "curr_id_bg", "ForExCnt"
                  ' ** Skip these.
4580            Case Else
4590              Select Case .ControlType
                  Case acTextBox, acCheckBox, acComboBox, acOptionButton, acCommandButton
4600                Select Case lngX
                    Case 1&
4610                  lngFrmFlds = lngFrmFlds + 1&
4620                  lngE = lngFrmFlds - 1&
4630                  ReDim Preserve arr_varFrmFld(FM_ELEMS, lngE)
4640                  arr_varFrmFld(FM_FLD_NAM, lngE) = .Name
4650                  arr_varFrmFld(FM_FLD_TAB, lngE) = .TabIndex
4660                  arr_varFrmFld(FM_FLD_VIS, lngE) = .Visible
4670                  arr_varFrmFld(FM_CHK_NAM, lngE) = vbNullString
4680                  arr_varFrmFld(FM_CHK_VAL, lngE) = CBool(False)
4690                  arr_varFrmFld(FM_VIEWCHK, lngE) = Null
4700                  arr_varFrmFld(FM_TOPO, lngE) = Null
4710                  arr_varFrmFld(FM_TOPC, lngE) = Null
4720                Case 2&
4730                  lngFrmFlds_ds = lngFrmFlds_ds + 1&
4740                  lngE = lngFrmFlds_ds - 1&
4750                  ReDim Preserve arr_varFrmFld_ds(FM_ELEMS, lngE)
4760                  arr_varFrmFld_ds(FM_FLD_NAM, lngE) = .Name
4770                  arr_varFrmFld_ds(FM_FLD_TAB, lngE) = .TabIndex
4780                  arr_varFrmFld_ds(FM_FLD_VIS, lngE) = .Visible
4790                  arr_varFrmFld_ds(FM_CHK_NAM, lngE) = vbNullString
4800                  arr_varFrmFld_ds(FM_CHK_VAL, lngE) = CBool(False)
4810                  arr_varFrmFld_ds(FM_VIEWCHK, lngE) = Null
4820                  arr_varFrmFld_ds(FM_TOPO, lngE) = Null
4830                  arr_varFrmFld_ds(FM_TOPC, lngE) = Null
4840                End Select
4850              End Select
4860            End Select
4870          End With
4880          DoEvents
4890        Next
4900        Set ctl = Nothing
4910        Set ctls = Nothing

            ' ****************************************
4920        gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
4930        frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
4940        gctlPBar_Lbl.Caption = "Step 15."
4950        DoEvents
4960        ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
4970        gctlPBar_Lbl_2.Caption = "Step 15."
4980        DoEvents
            ' ****************************************

4990        DoCmd.Hourglass True  ' ** Make sure it's still running.
5000        DoEvents

            ' ****************************************
            ' ** Step 16. Load form check boxes
            ' ****************************************
5010        gdblPBar_ThisStep = 16#
            ' ****************************************

            ' ** Get a list of the form data field check boxes.
5020        lngChkCnt = 0&
5030        Select Case lngX
            Case 1&
5040          Set ctls = frm1.Section("FormHeader").Controls
5050        Case 2&
5060          Set ctls = frm2.Section("FormHeader").Controls
5070        End Select
5080        For Each ctl In ctls
5090          With ctl
5100            Select Case .ControlType
                Case acCheckBox
5110              If Right(.Name, 4) = "_chk" Then
                    ' ** For non-visible ID fields, checkbox elements will remain vbNullString and False.
5120                lngChkCnt = lngChkCnt + 1&
5130                strTmp01 = Left(.Name, (Len(.Name) - 4))
5140                Select Case lngX
                    Case 1&
5150                  For lngY = 0& To (lngFrmFlds - 1&)
5160                    If arr_varFrmFld(FM_FLD_NAM, lngY) = strTmp01 Then
5170                      arr_varFrmFld(FM_CHK_NAM, lngY) = .Name
5180                      arr_varFrmFld(FM_CHK_VAL, lngY) = CBool(.Value)
5190                      Exit For
5200                    End If
5210                  Next
5220                Case 2&
5230                  For lngY = 0& To (lngFrmFlds_ds - 1&)
5240                    If arr_varFrmFld_ds(FM_FLD_NAM, lngY) = strTmp01 Then
5250                      arr_varFrmFld_ds(FM_CHK_NAM, lngY) = .Name
5260                      arr_varFrmFld_ds(FM_CHK_VAL, lngY) = CBool(.Value)
5270                      Exit For
5280                    End If
5290                  Next
5300                End Select
5310              End If
5320            End Select
5330          End With
5340          DoEvents
5350        Next

            ' ** Sort the list, left to right on form.
5360        FormFields_Sort lngX  ' ** Procedure: Below.
5370        DoEvents

5380      Next  ' ** lngX.

5390      DoCmd.Hourglass True  ' ** Make sure it's still running.
5400      DoEvents

          ' ****************************************
5410      gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
5420      frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
5430      gctlPBar_Lbl.Caption = "Step 16."
5440      DoEvents
5450      ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
5460      gctlPBar_Lbl_2.Caption = "Step 16."
5470      DoEvents
          ' ****************************************

          ' ****************************************
          ' ** Step 17. Load field array
          ' ****************************************
5480      gdblPBar_ThisStep = 17#
          ' ****************************************

5490      frm1.ShowFields_Sub THIS_PROC, 1  ' ** Form Procedure: frmTransaction_Audit_Sub.

          ' ****************************************
5500      gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
5510      frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
5520      gctlPBar_Lbl.Caption = "Step 17."
5530      DoEvents
5540      ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
5550      gctlPBar_Lbl_2.Caption = "Step 17."
5560      DoEvents
          ' ****************************************

          ' ** Lets sneak this in here.
5570      JC_Frm_LoadMapDropdowns  ' ** Module Procedure: modJrnlCol_Forms.

          ' ****************************************
          ' ** Step 18. Load preferences
          ' ****************************************
5580      gdblPBar_ThisStep = 18#
          ' ****************************************

5590      Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

5600      opgView_AfterUpdate  ' ** Procedure: Below.
5610      DoEvents

          ' ** See if you can clean up any mess.
5620      If Trim(.CurrentFilter1) = "And" Or Trim(.CurrentFilter2) = "And" Then
5630        .CurrentFilter1 = vbNullString
5640        .CurrentFilter2 = vbNullString
5650      ElseIf Left(Trim(.CurrentFilter1), 3) = "And" Or Left(Trim(.CurrentFilter2), 3) = "And" Then
5660        If Left(Trim(.CurrentFilter1), 3) = "And" Then
5670          .CurrentFilter1 = Trim(Mid(Trim(.CurrentFilter1), 4))
5680        ElseIf Left(Trim(.CurrentFilter2), 3) = "And" Then
5690          .CurrentFilter2 = Trim(Mid(Trim(.CurrentFilter2), 4))
5700        End If
5710      ElseIf Right(Trim(.CurrentFilter1), 3) = "And" Or Right(Trim(.CurrentFilter2), 3) = "And" Then
5720        If Right(Trim(.CurrentFilter1), 3) = "And" Then
5730          .CurrentFilter1 = Trim(Left(Trim(.CurrentFilter1), (Len(Trim(.CurrentFilter1)) - 3)))
5740        ElseIf Right(Trim(.CurrentFilter2), 3) = "And" Then
5750          .CurrentFilter2 = Trim(Left(Trim(.CurrentFilter2), (Len(Trim(.CurrentFilter2)) - 3)))
5760        End If
5770      End If

5780      ShowFields_Set lngFrmFlds, arr_varFrmFld, arr_varFrmFld_ds, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
5790      DoEvents

          ' ****************************************
5800      gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
5810      frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
5820      gctlPBar_Lbl.Caption = "Step 18."
5830      DoEvents
5840      ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
5850      gctlPBar_Lbl_2.Caption = "Step 18."
5860      DoEvents
          ' ****************************************

          ' ****************************************
          ' ** Step 19. Set filters
          ' ****************************************
5870      gdblPBar_ThisStep = 19#
          ' ****************************************

          ' ** I guess I'll always base this on Filter1, since they should be after the same fields,
          ' ** and a Filter2 shouldn't happen without a Filter1, unless I want to have independent filters.
          ' ** (Which, by the way, sounds like a good idea!)
5880      If IsNull(.CurrentFilter1) = False Then
5890        If .CurrentFilter1 <> vbNullString Then
5900          strFilter01 = .CurrentFilter1
5910          strFilter02 = frm2.SwapNames(strFilter01)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
5920          .CurrentFilter2 = strFilter02
5930          For lngX = 1& To 2&
5940            Select Case lngX
                Case 1&
5950              frm1.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
5960              frm1.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
5970              DoCmd.Hourglass True
5980              DoEvents
5990              For lngY = 0& To (lngFilts - 1)
6000                If Left(arr_varFilt(F_CTL, lngY), 3) <> "chk" Then
6010                  If IsNull(frmCrit.Controls(arr_varFilt(F_CTL, lngY))) = False Then
6020                    frm1.FilterRecs_Clr arr_varFilt(F_CTL, lngY) & "_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub.
6030                  End If
6040                  If arr_varFilt(F_CTL, lngY) = "cmbJournalType1" Then
6050                    If IsNull(frmCrit.cmbJournalType2) = False Then
6060                      frm1.FilterRecs_Clr "cmbJournalType2_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub.
6070                    End If
6080                    If IsNull(frmCrit.cmbJournalType3) = False Then
6090                      frm1.FilterRecs_Clr "cmbJournalType3_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub.
6100                    End If
6110                  End If
6120                Else
6130                  If .Controls(arr_varFilt(F_CTL, lngY)) = True Then
6140                    frm1.FilterRecs_Clr arr_varFilt(F_CTL, lngY) & "_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub.
6150                  End If
6160                End If
6170              Next  ' ** lngY.
6180            Case 2&
6190              frm2.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
6200              frm2.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
6210              For lngY = 0& To (lngFilts_ds - 1)
6220                If Left(arr_varFilt_ds(F_CTL, lngY), 3) <> "chk" Then
6230                  If IsNull(frmCrit.Controls(arr_varFilt_ds(F_CTL, lngY))) = False Then
6240                    frm2.FilterRecs_Clr arr_varFilt_ds(F_CTL, lngY) & "_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
6250                  End If
6260                  If arr_varFilt_ds(F_CTL, lngY) = "cmbJournalType1" Then
6270                    If IsNull(frmCrit.cmbJournalType2) = False Then
6280                      frm2.FilterRecs_Clr "cmbJournalType2_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
6290                    End If
6300                    If IsNull(frmCrit.cmbJournalType3) = False Then
6310                      frm2.FilterRecs_Clr "cmbJournalType3_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
6320                    End If
6330                  End If
6340                Else
6350                  If .Controls(arr_varFilt_ds(F_CTL, lngY)) = True Then
6360                    frm2.FilterRecs_Clr arr_varFilt_ds(F_CTL, lngY) & "_AfterUpdate", True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
6370                  End If
6380                End If
6390              Next  ' ** lngY.
6400            End Select
6410          Next  ' ** lngX.
6420        End If
6430      End If
6440      DoEvents

6450      DoCmd.Hourglass True  ' ** Make sure it's still running.
6460      DoEvents

          ' ****************************************
6470      gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
6480      frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
6490      gctlPBar_Lbl.Caption = "Step 19."
6500      DoEvents
6510      ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
6520      gctlPBar_Lbl_2.Caption = "Step 19."
6530      DoEvents
          ' ****************************************

          ' ****************************************
          ' ** Step 20. Set user report path
          ' ****************************************
6540      gdblPBar_ThisStep = 20#
          ' ****************************************

6550      If IsNull(.UserReportPath) = True Then
            ' ** if the user has a UserReportPath preference from another screen, use that.
6560        Set dbs = CurrentDb
6570        With dbs
              ' ** qryReport_Path_02 (qryReport_Path_01 (tblPreference_Control, just 'UserReportPath'),
              ' ** linked to tblPreference_User, by specified [usr]), grouped by prefuser_text.
6580          Set qdf = .QueryDefs("qryReport_Path_03")  '##dbs_id
6590          With qdf.Parameters
6600            ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
6610          End With
6620          Set rst = qdf.OpenRecordset
6630          With rst
6640            If .BOF = True And .EOF = True Then
                  ' ** Nope.
6650            Else
6660              .MoveLast
6670              lngRecs = .RecordCount
6680              .MoveFirst
6690              If lngRecs = 1& Then
                    ' ** If they only have one, or all of them are the same, use it!
6700                strRptPath = ![prefuser_text]
6710              End If
6720            End If
6730            .Close
6740          End With
6750          .Close
6760        End With
6770        If strRptPath <> vbNullString Then
6780          .UserReportPath = strRptPath
6790        End If
6800      End If
6810      DoEvents

6820      DoCmd.Hourglass True  ' ** Make sure it's still running.
6830      DoEvents

          ' ****************************************
6840      gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
6850      frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
6860      gctlPBar_Lbl.Caption = "Step 20."
6870      DoEvents
6880      ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
6890      gctlPBar_Lbl_2.Caption = "Step 20."
6900      DoEvents
          ' ****************************************

          ' ****************************************
          ' ** Step 21. Initialize options
          ' ****************************************
6910      gdblPBar_ThisStep = 21#
          ' ****************************************

6920      frmCrit.opgAccountSource_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
6930      DoEvents
6940      frmCrit.opgAssetSource_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
6950      DoEvents
6960      frmCrit.chkRevcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
6970      DoEvents
6980      frmCrit.chkRevcodeType_Expense_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
6990      DoEvents
7000      frmCrit.chkTaxcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
7010      DoEvents
7020      frmCrit.chkTaxcodeType_Deduction_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
7030      DoEvents
7040      frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
7050      DoEvents
7060      chkNarrowFont_AfterUpdate  ' ** Procedure: Below.
7070      DoEvents
7080      chkIncludeSharefaceTot_AfterUpdate  ' ** Procedure: Below.
7090      DoEvents
7100      chkPageOf_AfterUpdate  ' ** Procedure: Below.
7110      DoEvents
7120      chkSaveSizePos_AfterUpdate  ' ** Procedure: Below.
7130      DoEvents
7140      Select Case IsNull(frmCrit.cmbJournalType1)
          Case True
7150        frmCrit.cmbJournalType2.Enabled = False
7160        frmCrit.cmbJournalType2.BorderColor = WIN_CLR_DISR
7170        frmCrit.cmbJournalType2.BackStyle = acBackStyleTransparent
7180        frmCrit.cmbJournalType3.Enabled = False
7190        frmCrit.cmbJournalType3.BorderColor = WIN_CLR_DISR
7200        frmCrit.cmbJournalType3.BackStyle = acBackStyleTransparent
7210      Case False
7220        Select Case IsNull(frmCrit.cmbJournalType2)
            Case True
7230          frmCrit.cmbJournalType2.Enabled = True
7240          frmCrit.cmbJournalType2.BorderColor = CLR_LTBLU2
7250          frmCrit.cmbJournalType2.BackStyle = acBackStyleNormal
7260          frmCrit.cmbJournalType3.Enabled = False
7270          frmCrit.cmbJournalType3.BorderColor = WIN_CLR_DISR
7280          frmCrit.cmbJournalType3.BackStyle = acBackStyleTransparent
7290        Case False
7300          frmCrit.cmbJournalType2.Enabled = True
7310          frmCrit.cmbJournalType2.BorderColor = CLR_LTBLU2
7320          frmCrit.cmbJournalType2.BackStyle = acBackStyleNormal
7330          frmCrit.cmbJournalType3.Enabled = True
7340          frmCrit.cmbJournalType3.BorderColor = CLR_LTBLU2
7350          frmCrit.cmbJournalType3.BackStyle = acBackStyleNormal
7360        End Select
7370      End Select
7380      DoEvents
7390      SubFrmGfx_Load Me  ' ** Module Procedure: modTransactionAuditFuncs1.
7400      DoEvents
7410      frm1.Print_ChkCnt  ' ** Form Function: frmTransaction_Audit_Sub.
7420      DoEvents
7430      frm2.Print_ChkCnt  ' ** Form Function: frmTransaction_Audit_Sub_ds.
7440      DoEvents
7450      PrintTgls_Set frm1  ' ** Module Procedure: modTransactionAuditFuncs1.
7460      DoEvents

7470      DoCmd.Hourglass True  ' ** Make sure it's still running.
7480      DoEvents

          ' ****************************************
7490      gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
7500      frmMaint.ProgBar_Width_Maint gdblPBar_CurWidth, 2  ' ** Form Procedure: frmMenu_Maintenance.
7510      gctlPBar_Lbl.Caption = "Step 21."
7520      DoEvents
7530      ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
7540      gctlPBar_Lbl_2.Caption = "Step 21."
7550      DoEvents
          ' ****************************************

          ' ****************************************
          ' ** Step 22. Set form position and size
          ' ****************************************
7560      gdblPBar_ThisStep = 22#
          ' ****************************************

7570      If .frm_top <> 0 And .frm_left <> 0 And .frm_width <> 0 And .frm_height <> 0 Then
            ' ** The screen should already reflect what's open, what's not.
7580        DoCmd.SelectObject acForm, THIS_NAME, False
7590        lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
7600        lngMonitorNum = 1&: lngTmp02 = 0&: lngTmp03 = 0&
7610        EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
7620        If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
7630        If lngMonitorNum = 1& Then lngTmp02 = .frm_top
7640        lngTmp03 = .frm_left
7650        If lngTmp03 < 0 Then lngTmp03 = 15&
7660        DoCmd.SelectObject acForm, "frmTransaction_Audit", False
7670        DoEvents
7680        DoCmd.MoveSize lngTmp03, lngTmp02, .frm_width, .frm_height  '.frm_top
7690        If lngMonitorNum > 1& Then
7700          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
7710        End If
7720      End If
7730      DoEvents

7740    End With

EXITP:
7750    Set rst = Nothing
7760    Set qdf = Nothing
7770    Set dbs = Nothing
7780    Set ctl = Nothing
7790    Set ctls = Nothing
7800    Set frm1 = Nothing
7810    Set frm2 = Nothing
7820    Exit Sub

ERRH:
3910    DoCmd.Hourglass False
3920    Select Case ERR.Number
        Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7920    intShiftDown = (Shift And acShiftMask) > 0
7930    intAltDown = (Shift And acAltMask) > 0
7940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Date Picker Posting - From:  F5  {cmdCalendar1 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Posting - To:    F6  {cmdCalendar2 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Trade - From:    F7  {cmdCalendar3 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Trade - To:      F8  {cmdCalendar4 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Purchase - From: F9  {cmdCalendar5 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Purchase - To:   F10 {cmdCalendar6 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Posted - From:   F11 {cmdCalendar7 on frmTransaction_Audit_Sub_Criteria}
        ' **   Date Picker Posted - To:     F12 {cmdCalendar8 on frmTransaction_Audit_Sub_Criteria}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Date Posted:      D {TransDateStart on frmTransaction_Audit_Sub_Criteria}
        ' **   Include Archive:  H {chkIncludeArchive}
        ' **   JournalNo:        J {journalno on frmTransaction_Audit_Sub, frmTransaction_Audit_Sub_ds}
        ' **   Show Criteria:    K {cmdWidenToCriteria}
        ' **   Clear All:        R {cmdClearAll}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Include Hidden:   C {opgHidden_optInclude on frmTransaction_Audit_Sub_Criteria}
        ' **   Tax Deduction:    D {chkTaxcodeType_Deduction on frmTransaction_Audit_Sub_Criteria}
        ' **   Exclude Hidden:   E {opgHidden_optExclude on frmTransaction_Audit_Sub_Criteria}
        ' **   Tax Income:       I {chkTaxcodeType_Income on frmTransaction_Audit_Sub_Criteria}
        ' **   Asset Name:       M {opgAssetSource_optName on frmTransaction_Audit_Sub_Criteria}
        ' **   Narrow Font:      N {chkNarrowFont}
        ' **   Only Hidden:      O {opgHidden_optOnly on frmTransaction_Audit_Sub_Criteria}
        ' **   Shareface Total:  T {chkIncludeSharefaceTot}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   Inc/Exp Expense:  E {chkRevcodeType_Expense on frmTransaction_Audit_Sub_Criteria}
        ' **   Inc/Exp Income:   I {chkRevcodeType_Income on frmTransaction_Audit_Sub_Criteria}

        ' ** Plain keys.
7950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7960      Select Case intRetVal
          Case vbKeyF5
            ' **   Date Picker Posting - From:  F5  {cmdCalendar1}
7970        With Me
7980          intRetVal = 0
7990          If frmCrit.cmdCalendar1.Enabled = True Then
8000            .frmTransaction_Audit_Sub_Criteria.SetFocus
8010            frmCrit.cmdCalendar1.SetFocus
8020            frmCrit.cmdCalendar1_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
8030          End If
8040        End With
8050      Case vbKeyF6
            ' **   Date Picker Posting - To:    F6  {cmdCalendar2}
8060        With Me
8070          intRetVal = 0
8080          If frmCrit.cmdCalendar2.Enabled = True Then
8090            .frmTransaction_Audit_Sub_Criteria.SetFocus
8100            frmCrit.cmdCalendar2.SetFocus
8110            frmCrit.cmdCalendar2_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
8120          End If
8130        End With
8140      Case vbKeyF7
            ' **   Date Picker Trade - From:    F7  {cmdCalendar3}
8150        With Me
8160          intRetVal = 0
8170          If frmCrit.cmdCalendar3.Enabled = True Then
8180            .frmTransaction_Audit_Sub_Criteria.SetFocus
8190            frmCrit.cmdCalendar3.SetFocus
8200            frmCrit.cmdCalendar3_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
8210          End If
8220        End With
8230      Case vbKeyF8
            ' **   Date Picker Trade - To:      F8  {cmdCalendar4}
8240        With Me
8250          intRetVal = 0
8260          If frmCrit.cmdCalendar4.Enabled = True Then
8270            .frmTransaction_Audit_Sub_Criteria.SetFocus
8280            frmCrit.cmdCalendar4.SetFocus
8290            frmCrit.cmdCalendar4_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
8300          End If
8310        End With
8320      Case vbKeyF9
            ' **   Date Picker Purchase - From: F9  {cmdCalendar5}
8330        With Me
8340          intRetVal = 0
8350          If frmCrit.cmdCalendar5.Enabled = True Then
8360            .frmTransaction_Audit_Sub_Criteria.SetFocus
8370            frmCrit.cmdCalendar5.SetFocus
8380            frmCrit.cmdCalendar5_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
8390          End If
8400        End With
8410      Case vbKeyF10
            ' **   Date Picker Purchase - To:   F10 {cmdCalendar6}
8420        With Me
8430          intRetVal = 0
8440          If frmCrit.cmdCalendar6.Enabled = True Then
8450            .frmTransaction_Audit_Sub_Criteria.SetFocus
8460            frmCrit.cmdCalendar6.SetFocus
8470            frmCrit.cmdCalendar6_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
8480          End If
8490        End With
8500      Case vbKeyF11
            ' **   Date Picker Posted - From:   F11 {cmdCalendar7}
8510        With Me
8520          intRetVal = 0
8530          If frmCrit.cmdCalendar7.Enabled = True Then
8540            .frmTransaction_Audit_Sub_Criteria.SetFocus
8550            frmCrit.cmdCalendar7.SetFocus
8560            frmCrit.cmdCalendar7_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
8570          End If
8580        End With
8590      Case vbKeyF12
            ' **   Date Picker Posted - To:     F12 {cmdCalendar8}
8600        With Me
8610          intRetVal = 0
8620          If frmCrit.cmdCalendar8.Enabled = True Then
8630            .frmTransaction_Audit_Sub_Criteria.SetFocus
8640            frmCrit.cmdCalendar8.SetFocus
8650            frmCrit.cmdCalendar8_Click  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
8660          End If
8670        End With
8680      End Select
8690    End If

        ' ** Alt keys.
8700    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
8710      Select Case intRetVal
          Case vbKeyD
8720        With Me
8730          intRetVal = 0
8740          .frmTransaction_Audit_Sub_Criteria.SetFocus
8750          frmCrit.TransDateStart.SetFocus
8760        End With
8770      Case vbKeyH
8780        With Me
8790          intRetVal = 0
8800          If .chkIncludeArchive.Enabled = True Then
8810            .chkIncludeArchive.SetFocus
8820            .chkIncludeArchive = (Not .chkIncludeArchive)
8830            chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
8840          Else
8850            Beep
8860          End If
8870        End With
8880      Case vbKeyJ
8890        With Me
8900          intRetVal = 0
8910          Select Case .opgView
              Case .opgView_optForm.OptionValue
8920            .frmTransaction_Audit_Sub.SetFocus
8930            .frmTransaction_Audit_Sub.Form.FocusHolder.SetFocus
8940            If .frmTransaction_Audit_Sub.Form.journalno.Visible = True And .frmTransaction_Audit_Sub.Form.journalno.Enabled = True Then
8950              .frmTransaction_Audit_Sub.Form.journalno.SetFocus
8960            End If
8970          Case .opgView_optDatasheet.OptionValue
8980            .frmTransaction_Audit_Sub_ds.SetFocus
8990            If .frmTransaction_Audit_Sub_ds.Form.[Journal Number].Visible = True And .frmTransaction_Audit_Sub_ds.Form.[Journal Number].Enabled = True Then
9000              .frmTransaction_Audit_Sub_ds.Form.[Journal Number].SetFocus
9010            End If
9020          End Select
9030        End With
9040      Case vbKeyK
9050        With Me
9060          intRetVal = 0
9070          If .cmdWidenToCriteria.Enabled = True Then
9080            .cmdWidenToCriteria.SetFocus
9090            cmdWidenToCriteria_Click  ' ** Procedure: Below.
9100          Else
9110            Beep
9120          End If
9130        End With
9140      Case vbKeyR
            ' ** For some reason, the Alt-R on cmdClearAll doesn't want to invoke when it's sitting
            ' ** on a populated PostedDateEnd (probably others) after no records have been returned.
9150        With Me
9160          If dblFilterRecs = 0# Then
9170            .cmdClearAll.SetFocus
9180            cmdClearAll_Click  ' ** Procedure: Below.
9190          End If
9200        End With
9210      Case vbKeyX
9220        intRetVal = 0
9230        cmdClose_Click  ' ** Procedure: Below.
9240      End Select
9250    End If

        ' ** Alt-Shift keys.
9260    If (Not intCtrlDown) And intAltDown And intShiftDown Then
9270      Select Case intRetVal
          Case vbKeyC
9280        With Me
9290          intRetVal = 0
9300          .frmTransaction_Audit_Sub_Criteria.SetFocus
9310          frmCrit.opgHidden.SetFocus
9320          frmCrit.opgHidden = frmCrit.opgHidden_optInclude.OptionValue
9330          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
9340        End With
9350      Case vbKeyD
9360        With Me
9370          intRetVal = 0
9380          .frmTransaction_Audit_Sub_Criteria.SetFocus
9390          frmCrit.chkTaxcodeType_Deduction.SetFocus
9400          frmCrit.chkTaxcodeType_Deduction = (Not frmCrit.chkTaxcodeType_Deduction)
9410          frmCrit.chkTaxcodeType_Deduction_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
9420        End With
9430      Case vbKeyE
9440        With Me
9450          intRetVal = 0
9460          .frmTransaction_Audit_Sub_Criteria.SetFocus
9470          frmCrit.opgHidden.SetFocus
9480          frmCrit.opgHidden = frmCrit.opgHidden_optExclude.OptionValue
9490          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
9500        End With
9510      Case vbKeyI
9520        With Me
9530          intRetVal = 0
9540          .frmTransaction_Audit_Sub_Criteria.SetFocus
9550          frmCrit.chkTaxcodeType_Income.SetFocus
9560          frmCrit.chkTaxcodeType_Income = (Not frmCrit.chkTaxcodeType_Income)
9570          frmCrit.chkTaxcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
9580        End With
9590      Case vbKeyM
9600        With Me
9610          intRetVal = 0
9620          .frmTransaction_Audit_Sub_Criteria.SetFocus
9630          frmCrit.opgAssetSource = frmCrit.opgAssetSource_optName.OptionValue
9640          frmCrit.opgAssetSource.SetFocus
9650          frmCrit.opgAssetSource_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
9660        End With
9670      Case vbKeyN
9680        With Me
9690          intRetVal = 0
9700          .chkNarrowFont.SetFocus
9710          .chkNarrowFont = (Not .chkNarrowFont)
9720          chkNarrowFont_AfterUpdate  ' ** Procedure: Below.
9730        End With
9740      Case vbKeyO
9750        With Me
9760          intRetVal = 0
9770          .frmTransaction_Audit_Sub_Criteria.SetFocus
9780          frmCrit.opgHidden.SetFocus
9790          frmCrit.opgHidden = frmCrit.opgHidden_optOnly.OptionValue
9800          frmCrit.opgHidden_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
9810        End With
9820      Case vbKeyT
9830        With Me
9840          intRetVal = 0
9850          .chkIncludeSharefaceTot.SetFocus
9860          .chkIncludeSharefaceTot = (Not .chkIncludeSharefaceTot)
9870          chkIncludeSharefaceTot_AfterUpdate  ' ** Procedure: Below.
9880        End With
9890      End Select
9900    End If

        ' ** Ctrl keys.
9910    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
9920      Select Case intRetVal
          Case vbKeyG
9930        With Me
9940          intRetVal = 0
9950          If .chkPageOf.Enabled = True Then
9960            .chkPageOf.SetFocus
9970            .chkPageOf = (Not .chkPageOf)
9980            chkPageOf_AfterUpdate  ' ** Procedure: Below.
9990          End If
10000       End With
10010     End Select
10020   End If

        ' ** Ctrl-Shift keys.
10030   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyF
10050       With Me
10060         intRetVal = 0
10070         .FocusHolder.SetFocus
10080       End With
10090     Case vbKeyT
10100       intRetVal = 0
10110       FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
10120     End Select
10130   End If

        ' ** Ctrl-Alt keys.
10140   If intCtrlDown And intAltDown And (Not intShiftDown) Then
10150     Select Case intRetVal
          Case vbKeyE
10160       With Me
10170         intRetVal = 0
10180         .frmTransaction_Audit_Sub_Criteria.SetFocus
10190         frmCrit.chkRevcodeType_Expense.SetFocus
10200         frmCrit.chkRevcodeType_Expense = (Not frmCrit.chkRevcodeType_Expense)
10210         frmCrit.chkRevcodeType_Expense_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
10220       End With
10230     Case vbKeyI
10240       With Me
10250         intRetVal = 0
10260         .frmTransaction_Audit_Sub_Criteria.SetFocus
10270         frmCrit.chkRevcodeType_Income.SetFocus
10280         frmCrit.chkRevcodeType_Income = (Not frmCrit.chkRevcodeType_Income)
10290         frmCrit.chkRevcodeType_Income_AfterUpdate  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
10300       End With
10310     End Select
10320   End If

EXITP:
10330   KeyCode = intRetVal
10340   Exit Sub

ERRH:
3910    intRetVal = 0
3920    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3930    Case Else
3940      THAT_PROC = THIS_PROC
3950      That_Erl = Erl
3960      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
3970    End Select
3980    Resume EXITP

End Sub

Public Sub Form_Resize()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

10410   FormResize_TransAud blnIsOpen, blnFromCkgFlds, lngFrm_Width, lngSub_Width_Diff, lngSub_Width, lngSub_Width_New, _
          lngForm_Width, lngForm_Width_New, lngSub_Width_Min, lngFrm_Height, lngDetail_Height_Diff, lngDetail_Height, _
          lngDetail_Height_New, lngDetail_Height_Min, lngCkgFldsBox_Height, lngCkgFldsBox_Offset, lngCritSub_Width, _
          lngHLine_Offset, lngClose_Left, lngSizable_Offset, lngCritSub_Max, lngSub_Height, lngSelectBox_Top, lngSelectLbl_Top, _
          lngSelectLegendLbl_Top, lngSelectLbl2_Top, lngNarrowFont_Top, lngShareface_Top, lngSelectAll_Top, lngSelectNone_Top, _
          lngChkBoxLbl_TopOffset, lngCkgFldsVLine_Top, lngFrmFlds, arr_varFrmFld, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
10420   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub Form_Timer()

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strFormName As String

10510   With Me
10520     .TimerInterval = 0&
10530     If gblnClosing = False Then
10540       If gblnSetFocus = True Then
10550         gblnSetFocus = False
10560         strFormName = vbNullString
10570 On Error Resume Next
10580         Do Until strFormName = THIS_NAME
10590           DoCmd.SelectObject acForm, THIS_NAME, False
10600           strFormName = Screen.ActiveForm.Name
10610         Loop
10620 On Error GoTo ERRH
10630         CmdBars_Hide True  ' ** Module Procedure: modWindowFunctions.
10640       ElseIf blnRecsZero = True Then
10650         blnRecsZero = False
10660         If intRecsZero < 20 Then
10670           FilterRecsZero  ' ** Module Procedure: modTransactionAuditFuncs1.
10680         Else
10690           intRecsZero = 0
10700         End If
10710       Else
              ' ** I'm using this from rptTransaction_Audit_01.
              'MAY NOT BE HIT!
10720         gblnSetFocus = False
10730         DoCmd.Restore
10740       End If
10750     End If  ' ** gblnClosing.
10760     If blnProgBarOff = True Then
10770       blnProgBarOff = False
10780       Beep
10790       gctlPBar_Box1_2.Visible = False
10800       gctlPBar_Box2_2.Visible = False
10810       ProgBar_Width_Trans Me, False, 1  ' ** Module Procedure: modTransactionAuditFuncs1.
10820       gctlPBar_Lbl_2.Visible = False
10830       gblnMessage = False
10840       If IsLoaded(strCallingForm) = True Then  ' ** Module Function: modFileUtilities.
10850         DoCmd.Close acForm, strCallingForm
10860         DoEvents
10870       End If
10880       .ProgBar_lbl.Caption = "Step 0."  ' ** Reset for Ledger Archive.
10890       DoCmd.Hourglass False
10900     End If
10910     If gblnGoToReport = True Then
10920       gblnGoToReport = False
10930       If .cmdPrint.Enabled = True Then
10940         .cmdPrint.SetFocus
10950       Else
10960         .cmdClose.SetFocus
10970       End If
10980       Beep
10990       DoCmd.Hourglass False
11000       DoEvents
11010     End If
11020     If blnScrollBar = True Then
11030       blnScrollBar = False
11040       Set frm1 = .frmTransaction_Audit_Sub.Form
11050       Set frm2 = .frmTransaction_Audit_Sub_ds.Form
            ' ** Make sure they're scrolled all the way to the left.
11060       Select Case .opgView
            Case .opgView_optForm.OptionValue
11070         .frmTransaction_Audit_Sub.SetFocus
11080         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
11090       Case .opgView_optDatasheet.OptionValue
11100         .frmTransaction_Audit_Sub_ds.SetFocus
11110         fSetScrollBarPosHZ frm2, 1&  ' ** Module Function: modScrollBarFuncs.
11120       End Select
11130       frm2.ColWidth_Set True  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
11140       If .FilterRecs <> "All" Then
11150         lngRecsCur = .frmTransaction_Audit_Sub.Form.RecCnt  ' ** Form Function: frmTransaction_Audit_Sub
11160         If Val(.FilterRecs) <> lngRecsCur Then
11170           .FilterRecs = lngRecsCur
11180         End If
11190       End If
11200       If gblnTimer = True Then
11210         If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
11220           Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
11230           gblnTimer = False
11240           DoCmd.Hourglass False
11250         End If
11260       End If
11270     End If
11280   End With

EXITP:
11290   Set frm1 = Nothing
11300   Set frm2 = Nothing
11310   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

11410   Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
11420     Response = acDataErrContinue
11430   Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
11440     If Right(Me.ActiveControl.Name, 5) = "Start" Or Right(Me.ActiveControl.Name, 3) = "End" Then
11450       MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Entry"
11460       Response = acDataErrContinue
11470     End If
11480   Case 2237  ' ** 2237: The text you entered isn't an item in the list.
11490     Response = acDataErrDisplay
11500   Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
11510     Beep
11520     MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
11530     Response = acDataErrContinue
11540   Case 2501  ' ** 2501: The '|' action was Canceled.
11550     Response = acDataErrContinue
11560   Case Else
11570     If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
11580     zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
11590   End Select

EXITP:
11600   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngX As Long

11710   With Me

11720     .TimerInterval = 0&

11730     Select Case .chkSaveSizePos
          Case True
            ' ** Save the form's dimensions
11740       GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.
11750       .frm_top = lngTop
11760       .frm_left = lngLeft
11770       .frm_width = lngWidth
11780       .frm_height = lngHeight
11790     Case False
11800       .frm_top = 0&
11810       .frm_left = 0&
11820       .frm_width = 0&
11830       .frm_height = 0&
11840     End Select
11850     DoEvents

          ' ** Update the datasheet column widths for saving in Preferences.
11860     .frmTransaction_Audit_Sub_ds.Form.ColArray_Width  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
11870     DoEvents

11880     FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.

          'IF I KEPT AN ARRAY OF THE PREFS, THEN ONLY SAVED THOSE THAT CHANGED, WOULD THAT SPEED IT UP?
          ' ** It's all these Pref_Save's that are slowing down the Close!
11890     Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
11900     DoEvents
11910     Pref_Save (THIS_NAME & "_Sub_Criteria")  ' ** Module Procedure: modPreferenceFuncs.
11920     DoEvents
11930     Pref_Save (THIS_NAME & "_Sub")  ' ** Module Procedure: modPreferenceFuncs.
11940     DoEvents
11950     Pref_Save (THIS_NAME & "_Sub_ds")  ' ** Module Procedure: modPreferenceFuncs.
11960     DoEvents

          ' ** Make sure all reports are closed.
11970     Do While Reports.Count > 0
11980       DoCmd.Close acReport, Reports(0).Name
11990     Loop
12000     DoEvents

12010     Set dbs = CurrentDb
          ' ** Empty tmpAccount.
12020     Set qdf = dbs.QueryDefs("qryTransaction_Audit_12_01")
12030     qdf.Execute
12040     Set qdf = Nothing
12050     DoEvents
12060     dbs.Close
12070     Set dbs = Nothing
12080     DoEvents

12090     gstrCrtRpt_Ordinal = vbNullString  ' ** Borrowing this variable from Court Reports.

12100     Select Case strCallingForm
          Case "frmMenu_Maintenance"
12110       gblnSetFocus = True
12120       gstrReturningForm = THIS_NAME
12130       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
12140         DoCmd.SelectObject acForm, strCallingForm, False
12150         Forms(strCallingForm).RestoreForm
12160         Forms(strCallingForm).TimerInterval = 250&
12170       Else
12180         DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Utility"
12190       End If
12200       gblnClosing = False
12210     Case vbNullString
12220       gblnSetFocus = True
12230       gstrReturningForm = THIS_NAME
12240       strDocName = "frmMenu_Maintenance"
12250       If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
12260         DoCmd.SelectObject acForm, strDocName, False
12270         Forms(strDocName).RestoreForm
12280         Forms(strDocName).TimerInterval = 250&
12290       Else
12300         DoCmd.OpenForm strDocName, , , , , , "frmMenu_Utility"
12310       End If
12320       gblnClosing = False
12330     Case Else
12340       DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
12350       gblnClosing = False
12360     End Select

12370   End With

EXITP:
12380   Set qdf = Nothing
12390   Set dbs = Nothing
12400   Set frmCrit = Nothing
12410   Set frmMaint = Nothing
12420   Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Public Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

12510   DetailMouse_Main_TA blnClearAll_Focus, blnSelectAll_Focus, _
          blnSelectNone_Focus, blnCkgFlds_Focus, blnWidenToCrit_Focus, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
12520   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

12610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12620   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

12710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12720   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Public Sub cmdClose_Click()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

12810   With Me

12820     DoCmd.Hourglass True
12830     DoEvents

12840     .TimerInterval = 0&  ' ** Make sure nothing is still running.

12850     gblnClosing = True
12860     DoCmd.Close acForm, THIS_NAME

12870   End With

EXITP:
12880   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12920   intShiftDown = (Shift And acShiftMask) > 0
12930   intAltDown = (Shift And acAltMask) > 0
12940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12960     Select Case intRetVal
          Case vbKeyTab
12970       With Me
12980         intRetVal = 0
12990         .cmdClearAll.SetFocus
13000       End With
13010     End Select
13020   End If

        ' ** Shift keys.
13030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13040     Select Case intRetVal
          Case vbKeyTab
13050       With Me
13060         intRetVal = 0
13070         If .cmdExcel.Enabled = True Then
13080           .cmdExcel.SetFocus
13090         ElseIf .cmdWord.Enabled = True Then
13100           .cmdWord.SetFocus
13110         ElseIf .ckgFlds_chkCurrID.Visible = True And .ckgFlds_chkCurrID.Enabled = True Then
13120           .ckgFlds_chkCurrID.SetFocus
13130         Else
13140           .chkSaveSizePos.SetFocus
13150         End If
13160       End With
13170     End Select
13180   End If

EXITP:
13190   KeyCode = intRetVal
13200   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

13310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13320   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub FilterRecs_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_box_MouseMove"

13410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13420   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub FilterRecs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_MouseMove"

13510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13520   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdClearAll_GotFocus()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_GotFocus"

13610   Clear_Handler_TA THIS_PROC, blnClearAll_Focus, blnClearAll_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
13620   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdClearAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_MouseDown"

13710   Clear_Handler_TA THIS_PROC, blnClearAll_Focus, blnClearAll_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
13720   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Public Sub cmdClearAll_Click()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_Click"

13810   ClearAll_TA strFilter01, dblFilterRecs, Me, frmCrit  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
13820   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdClearAll_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13920   intShiftDown = (Shift And acShiftMask) > 0
13930   intAltDown = (Shift And acAltMask) > 0
13940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13960     Select Case intRetVal
          Case vbKeyTab
13970       With Me
13980         intRetVal = 0
13990         .frmTransaction_Audit_Sub_Criteria.SetFocus
14000         frmCrit.FocusHolder.SetFocus
14010         If frmCrit.journalno.Enabled = True Then
14020           frmCrit.journalno.SetFocus
14030         End If
14040       End With
14050     End Select
14060   End If

        ' ** Shift keys.
14070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14080     Select Case intRetVal
          Case vbKeyTab
14090       With Me
14100         intRetVal = 0
14110         .cmdClose.SetFocus
14120       End With
14130     End Select
14140   End If

EXITP:
14150   KeyCode = intRetVal
14160   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub cmdClearAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_MouseMove"

14210   Clear_Handler_TA THIS_PROC, blnClearAll_Focus, blnClearAll_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
14220   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdClearAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_MouseUp"

14310   Clear_Handler_TA THIS_PROC, blnClearAll_Focus, blnClearAll_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
14320   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdClearAll_LostFocus()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearAll_LostFocus"

14410   Clear_Handler_TA THIS_PROC, blnClearAll_Focus, blnClearAll_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
14420   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub chkIncludeArchive_lbl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_lbl3_MouseMove"

14510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14520   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub chkIncludeArchive_lbl4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_lbl4_MouseMove"

14610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14620   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Public Sub chkIncludeArchive_AfterUpdate()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

        Dim qdfAll1 As DAO.QueryDef, qdfAll2 As DAO.QueryDef, frm1 As Access.Form, frm2 As Access.Form
        Dim strRecSource1 As String, strRecSource2 As String
        Dim strSortNow1 As String, strSortNow2 As String

14710   With Me

14720     DoCmd.Hourglass True
14730     DoEvents

14740     If gblnMessage = False Then
            ' ****************************************
            ' ** Progress Bar setup.
            ' ****************************************
14750       gdblPBar_MaxWidth = .ProgBar_box.Width
14760       gdblPBar_Steps = 8#
14770       gdblPBar_Increment = (gdblPBar_MaxWidth / gdblPBar_Steps)
14780       gdblPBar_CurWidth = 0#
14790       gdblPBar_ThisStep = 0#
14800       Set gctlPBar_Box1 = .ProgBar_box
14810       Set gctlPBar_Box2 = .ProgBar_box2
            'Set gctlPBar_Bar = .ProgBar_bar
14820       Set gctlPBar_Lbl = .ProgBar_lbl
            ' ****************************************
            ' ****************************************
14830       gctlPBar_Box1.Visible = True
14840       gctlPBar_Box2.Visible = True
14850       ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
            'gctlPBar_Bar.Width = gdblPBar_CurWidth
14860       ProgBar_Width_Trans Me, True, 1  ' ** Module Procedure: modTransactionAuditFuncs1.
            'gctlPBar_Bar.Visible = True
14870       If CurrentUser = "Superuser" Then
14880         gctlPBar_Lbl.Visible = True
14890       Else
14900         gctlPBar_Lbl.Visible = False
14910       End If
14920       DoEvents
            ' ****************************************
            ' ****************************************
            ' ** Step 1. Clear All
            ' ****************************************
14930       gdblPBar_ThisStep = 1#
            ' ****************************************
14940     End If

14950     cmdClearAll_Click  ' ** Procedure: Below.
14960     DoEvents

14970     ReloadCritPrefs_TA frmCrit  ' ** Module Procedure: modTransactionAuditFuncs2.
14980     DoEvents

14990     Set frm1 = .frmTransaction_Audit_Sub.Form
15000     Set frm2 = .frmTransaction_Audit_Sub_ds.Form

15010     If gblnMessage = False Then
            ' ****************************************
15020       gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
15030       ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
            'gctlPBar_Bar.Width = gdblPBar_CurWidth
15040       gctlPBar_Lbl.Caption = "Step 1."
15050       DoEvents
            ' ****************************************
            ' ****************************************
            ' ** Step 2. SortNow
            ' ****************************************
15060       gdblPBar_ThisStep = 2#
            ' ****************************************
15070     End If

15080     strSortNow1 = frm1.SortNow_Get  ' ** Form Function: frmTransaction_Audit_Sub.
15090     strSortNow2 = frm2.SortNow_Get  ' ** Form Function: frmTransaction_Audit_Sub_ds.

15100     Select Case .chkIncludeArchive  ' ** chkIncludeArchive_lbl.Caption is empty.
          Case True
15110       .chkIncludeArchive_lbl2.FontBold = True
15120       strRecSource1 = "qryTransaction_Audit_03"
15130       strRecSource2 = "qryTransaction_Audit_03a"
15140     Case False
15150       .chkIncludeArchive_lbl2.FontBold = False
15160       strRecSource1 = "qryTransaction_Audit_01"
15170       strRecSource2 = "qryTransaction_Audit_01a"
15180     End Select
15190     DoEvents

15200     If gblnMessage = False Then
            ' ****************************************
15210       gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
15220       ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
            'gctlPBar_Bar.Width = gdblPBar_CurWidth
15230       gctlPBar_Lbl.Caption = "Step 2."
15240       DoEvents
            ' ****************************************
15250     End If

15260     If frm1.RecordSource <> strRecSource1 Or frm2.RecordSource <> strRecSource2 Then

15270       If gblnMessage = False Then
              ' ****************************************
              ' ** Step 3. RecordSource 1
              ' ****************************************
15280         gdblPBar_ThisStep = 3#
              ' ****************************************
15290       End If

15300       frm1.RecordSource = strRecSource1

15310       If gblnMessage = False Then
              ' ****************************************
15320         gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
15330         ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
              'gctlPBar_Bar.Width = gdblPBar_CurWidth
15340         gctlPBar_Lbl.Caption = "Step 3."
15350         DoEvents
              ' ****************************************
              ' ****************************************
              ' ** Step 4. RecordSource 2
              ' ****************************************
15360         gdblPBar_ThisStep = 4#
              ' ****************************************
15370       End If

15380       frm2.RecordSource = strRecSource2

15390       If gblnMessage = False Then
              ' ****************************************
15400         gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
15410         ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
              'gctlPBar_Bar.Width = gdblPBar_CurWidth
15420         gctlPBar_Lbl.Caption = "Step 4."
15430         DoEvents
              ' ****************************************
              ' ****************************************
              ' ** Step 5.
              ' ****************************************
15440         gdblPBar_ThisStep = 5#
              ' ****************************************
15450       End If

15460       Set qdfAll1 = CurrentDb.QueryDefs(strRecSource1)
15470       Set qdfAll2 = CurrentDb.QueryDefs(strRecSource2)
15480       Set rstAll1 = qdfAll1.OpenRecordset
15490       Set rstAll2 = qdfAll2.OpenRecordset
15500       If (rstAll1.BOF = True And rstAll1.EOF = True) Or (rstAll2.BOF = True And rstAll2.EOF = True) Then
              ' ** Shouldn't have been able to get here!
15510       Else
15520         rstAll1.MoveLast
15530         rstAll2.MoveLast
15540         lngTotalRecs1 = rstAll1.RecordCount
15550         lngTotalRecs2 = rstAll2.RecordCount

15560         If gblnMessage = False Then
                ' ****************************************
15570           gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
15580           ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
                'gctlPBar_Bar.Width = gdblPBar_CurWidth
15590           gctlPBar_Lbl.Caption = "Step 5."
15600           DoEvents
                ' ****************************************
                ' ****************************************
                ' ** Step 6.
                ' ****************************************
15610           gdblPBar_ThisStep = 6#
                ' ****************************************
15620         End If

15630         frm1.TotalRecs_Set lngTotalRecs1  ' ** Form Procedure: frmTransaction_Audit_Sub.
15640         frm2.TotalRecs_Set lngTotalRecs2  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
15650         rstAll1.MoveFirst
15660         rstAll2.MoveFirst

15670         If gblnMessage = False Then
                ' ****************************************
15680           gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
15690           ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
                'gctlPBar_Bar.Width = gdblPBar_CurWidth
15700           gctlPBar_Lbl.Caption = "Step 6."
15710           DoEvents
                ' ****************************************
                ' ****************************************
                ' ** Step 7. Sort subform
                ' ****************************************
15720           gdblPBar_ThisStep = 7#
                ' ****************************************
15730         End If

15740         frmCrit.RstAllSet strRecSource1, 1  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
15750         frmCrit.RstAllSet strRecSource2, 2  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.

15760         If gblnMessage = False Then
                ' ****************************************
15770           gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
15780           ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
                'gctlPBar_Bar.Width = gdblPBar_CurWidth
15790           gctlPBar_Lbl.Caption = "Step 7."
15800           DoEvents
                ' ****************************************
15810         End If

15820       End If

15830       If gblnMessage = False Then
              ' ****************************************
              ' ** Step 8. Query 4
              ' ****************************************
15840         gdblPBar_ThisStep = 8#
              ' ****************************************
15850       End If

15860       frm1.OrderBy = strSortNow1
15870       frm2.OrderBy = strSortNow2
15880       If strSortNow1 = vbNullString Or strSortNow2 = vbNullString Then
15890         frm1.OrderByOn = False
15900         frm2.OrderByOn = False
15910       Else
15920         frm1.OrderByOn = True
15930         frm2.OrderByOn = True
15940       End If
15950       DoEvents

15960       If gblnMessage = False Then
              ' ****************************************
15970         gdblPBar_CurWidth = (gdblPBar_ThisStep * gdblPBar_Increment)
15980         ProgBar_Width_Trans Me, gdblPBar_CurWidth, 2  ' ** Module Procedure: modTransactionAuditFuncs1.
              'gctlPBar_Bar.Width = gdblPBar_CurWidth
15990         gctlPBar_Lbl.Caption = "Step 8."
16000         DoEvents
              ' ****************************************
16010       End If

16020     End If

16030     If gblnMessage = False Then
16040       blnProgBarOff = True
16050       .TimerInterval = 3000&
16060     End If

16070     Beep
16080     DoCmd.Hourglass False

16090   End With

EXITP:
16100   Set frm1 = Nothing
16110   Set frm2 = Nothing
16120   Set qdfAll1 = Nothing
16130   Set qdfAll2 = Nothing
16140   Exit Sub

ERRH:
3910    DoCmd.Hourglass False
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub chkIncludeArchive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_MouseMove"

16210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16220   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub opgView_AfterUpdate()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_AfterUpdate"

        Dim strSortNow As String

16310   With Me
16320     Select Case .opgView
          Case .opgView_optForm.OptionValue
16330       .opgView_optForm_lbl.FontBold = True
16340       .opgView_optDatasheet_lbl.FontBold = False
16350       .frmTransaction_Audit_Sub.Visible = True
16360       .frmTransaction_Audit_Sub_ds.Visible = False
16370       strSortNow = .frmTransaction_Audit_Sub_ds.Form.SortNow_Get  ' ** Form Function: frmTransaction_Audit_Sub_ds.
16380       .frmTransaction_Audit_Sub.Form.SortNow_Set strSortNow  ' ** Form Procedure: frmTransaction_Audit_Sub.
16390       .frmTransaction_Audit_Sub.Form.OrderBy = strSortNow
16400       .frmTransaction_Audit_Sub.Form.OrderByOn = True
16410     Case .opgView_optDatasheet.OptionValue
16420       .opgView_optForm_lbl.FontBold = False
16430       .opgView_optDatasheet_lbl.FontBold = True
16440       strSortNow = .frmTransaction_Audit_Sub.Form.SortNow_Get  ' ** Form Function: frmTransaction_Audit_Sub.
16450       .frmTransaction_Audit_Sub_ds.Form.SortNow_Set strSortNow  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
16460       .frmTransaction_Audit_Sub_ds.Form.OrderBy = strSortNow
16470       .frmTransaction_Audit_Sub_ds.Form.OrderByOn = True
16480       With .frmTransaction_Audit_Sub_ds
16490         With .Form
16500           If .journalno.ColumnHidden = False Then
16510             .journalno.ColumnHidden = True
16520             .journaltype.ColumnHidden = True
16530             .transdate.ColumnHidden = True
16540             .accountno.ColumnHidden = True
16550             .shortname.ColumnHidden = True
16560             .asset_description.ColumnHidden = True
16570             .shareface.ColumnHidden = True
16580             .ICash.ColumnHidden = True
16590             .PCash.ColumnHidden = True
16600             .assetdate.ColumnHidden = True
16610             .PurchaseDate.ColumnHidden = True
16620             .ledger_description.ColumnHidden = True
16630             .RecurringItem.ColumnHidden = True
16640             .revcode_DESC.ColumnHidden = True
16650             .revcode_TYPE_Description.ColumnHidden = True
16660             .taxcode_description.ColumnHidden = True
16670             .TaxCode_Type_Description.ColumnHidden = True
16680             .journal_USER.ColumnHidden = True
16690             .posted.ColumnHidden = True
16700             .ledger_HIDDEN.ColumnHidden = True
16710           End If
16720         End With
16730         .Visible = True
16740       End With
16750       .frmTransaction_Audit_Sub.Visible = False
16760     End Select
16770     .frmTransaction_Audit_Sub.Form.FilterOn = True
16780   End With

        ' ** AcDefView Enumeration:
        ' **   0  acDefViewSingle      Displays one record at a time. (Default)
        ' **   1  acDefViewContinuous  The form displays multiple records, each in its own copy of the form's detail section.
        ' **   2  acDefViewDatasheet   Displays the form fields arranged in rows and columns like a spreadsheet.
        ' **   3  acDefViewPivotTable  Displays the form as a PivotTable.
        ' **   4  acDefViewPivotChart  Displays the form as a PivotChart.
        ' **   5  acDefViewSplitForm   Displays the form in Split Form view.

        ' ** AcCurrentView enumeration:
        ' **   0  acCurViewDesign        The object is in Design view.
        ' **   1  acCurViewFormBrowse    The object is in Form view.
        ' **   2  acCurViewDatasheet     The object is in Datasheet view.
        ' **   3  acCurViewPivotTable    The object is in PivotTable view.
        ' **   4  acCurViewPivotChart    The object is in PivotChart view.
        ' **   5  acCurViewPreview       The object is in Print Preview.
        ' **   6  acCurViewReportBrowse  The object is in Report view.
        ' **   7  acCurViewLayout        The object is in Layout view.

EXITP:
16790   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case 2001  ' ** You canceled the previous operation.
          ' ** Ignore.
3920    Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub opgView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_MouseMove"

16810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16820   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub opgView_optForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_optForm_MouseMove"

16910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16920   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub opgView_optDatasheet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_optDatasheet_MouseMove"

17010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17020   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdWidenToCriteria_GotFocus()

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWidenToCriteria_GotFocus"

17110   Widen_Handler_TA THIS_PROC, blnWidenToCrit_Focus, blnWidenToCrit_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
17120   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdWidenToCriteria_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWidenToCriteria_MouseDown"

17210   Widen_Handler_TA THIS_PROC, blnWidenToCrit_Focus, blnWidenToCrit_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
17220   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Public Sub cmdWidenToCriteria_Click()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWidenToCriteria_Click"

17310   With Me
17320     CriteriaWiden lngFrm_Top, lngFrm_Left, lngFrm_Width, lngFrm_Height, Me, frmCrit  ' ** Module Procedure: modTransactionAuditFuncs1.
17330     .frmTransaction_Audit_Sub_Criteria.SetFocus
17340   End With

EXITP:
17350   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdWidenToCriteria_KeyDown(KeyCode As Integer, Shift As Integer)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWidenToCriteria_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17420   intShiftDown = (Shift And acShiftMask) > 0
17430   intAltDown = (Shift And acAltMask) > 0
17440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17460     Select Case intRetVal
          Case vbKeyTab
17470       With Me
17480         intRetVal = 0
17490         .frmTransaction_Audit_Sub_Criteria.SetFocus
17500       End With
17510     End Select
17520   End If

        ' ** Shift keys.
17530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17540     Select Case intRetVal
          Case vbKeyTab
17550       With Me
17560         intRetVal = 0
17570         .opgView.SetFocus
17580       End With
17590     End Select
17600   End If

EXITP:
17610   KeyCode = intRetVal
17620   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub cmdWidenToCriteria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWidenToCriteria_MouseMove"

17710   Widen_Handler_TA THIS_PROC, blnWidenToCrit_Focus, blnWidenToCrit_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
17720   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdWidenToCriteria_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWidenToCriteria_MouseUp"

17810   Widen_Handler_TA THIS_PROC, blnWidenToCrit_Focus, blnWidenToCrit_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
17820   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdWidenToCriteria_LostFocus()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWidenToCriteria_LostFocus"

17910   Widen_Handler_TA THIS_PROC, blnWidenToCrit_Focus, blnWidenToCrit_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
17920   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

18010   Preview_TA strFilter01, dblFilterRecs, lngFrmFlds, arr_varFrmFld, lngFrmFlds_ds, arr_varFrmFld_ds, blnIsMaximized, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18020   Exit Sub

ERRH:
3910    Me.Visible = True
3920    Select Case ERR.Number
        Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
          Case vbKeyTab
18170       With Me
18180         intRetVal = 0
18190         .cmdPrint.SetFocus
18200       End With
18210     End Select
18220   End If

        ' ** Shift keys.
18230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18240     Select Case intRetVal
          Case vbKeyTab
18250       With Me
18260         intRetVal = 0
18270         .cmdSelectNone.SetFocus
18280       End With
18290     End Select
18300   End If

        ' ** Ctrl-Shift keys.
18310   If intCtrlDown And (Not intAltDown) And intShiftDown Then
18320     Select Case intRetVal
          Case vbKeyTab
18330       With Me
18340         intRetVal = 0
18350         If .ckgFlds_chkCurrID.Visible = True And .ckgFlds_chkCurrID.Enabled = True Then
18360           .ckgFlds_chkCurrID.SetFocus
18370         Else
18380           .ckgFlds_cmd.SetFocus
18390         End If
18400       End With
18410     End Select
18420   End If

EXITP:
18430   KeyCode = intRetVal
18440   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

18510   Print_TA strFilter01, dblFilterRecs, lngFrmFlds, arr_varFrmFld, lngFrmFlds_ds, arr_varFrmFld_ds, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18520   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdWord_Click()

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

18610   Word_TA strFilter01, dblFilterRecs, lngFrmFlds, arr_varFrmFld, lngFrmFlds_ds, arr_varFrmFld_ds, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18620   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

18710   Excel_TA strFilter01, dblFilterRecs, lngFrmFlds, arr_varFrmFld, lngFrmFlds_ds, arr_varFrmFld_ds, Me  ' ** Module Procedure: modTransactionAuditFuncs1.

EXITP:
18720   Exit Sub

ERRH:
3910    DoCmd.Hourglass False
3920    Select Case ERR.Number
        Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub cmdSelect_lbl3_cmd_DblClick(Cancel As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_lbl3_cmd_DblClick"

18810   With Me

18820     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
18830       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
18840     End If

18850     Select Case .cmdSelect_Legend_box.Visible
          Case True

18860       .cmdSelect_Legend_lbl.Visible = False
18870       .cmdSelect_Legend_box.Visible = False
18880       .cmdSelect_Legend_Tgl_lbl1.Visible = False
18890       .cmdSelect_Legend_Tgl_lbl2.Visible = False
18900       .cmdSelect_Legend_tgl_on_raised_img.Visible = False
18910       .cmdSelect_Legend_tgl_on_raised_focus_img.Visible = False
18920       .cmdSelect_Legend_tgl_on.Visible = False
18930       .cmdSelect_Legend_tgl_off_raised_img.Visible = False
18940       .cmdSelect_Legend_tgl_off_raised_focus_img.Visible = False
18950       .cmdSelect_Legend_tgl_off.Visible = False
18960       .cmdSelect_Legend_vline01.Visible = False
18970       .cmdSelect_Legend_vline02.Visible = False

18980       .chkNarrowFont_box.Left = .cmdSelect_Legend_lbl_alt_line1.Left
18990       .chkNarrowFont.Left = (.chkNarrowFont_box.Left + (lngNarrowFont_Left - lngNarrowFontBox_Left))
19000       .chkNarrowFont_lbl.Left = (.chkNarrowFont.Left + lngChkBoxLbl_LeftOffset)
19010       .chkIncludeSharefaceTot.Left = .chkNarrowFont.Left
19020       .chkIncludeSharefaceTot_lbl.Left = .chkNarrowFont_lbl.Left
19030       .chkPageOf.Left = .chkNarrowFont.Left + (lngPageOf_Left - lngNarrowFont_Left)
19040       .chkPageOf_lbl.Left = (.chkPageOf.Left + lngChkBoxLbl_LeftOffset)
19050       .chkSaveSizePos.Left = .chkPageOf.Left
19060       .chkSaveSizePos_lbl.Left = .chkPageOf_lbl.Left
19070       .chkNarrowFont_vline02.Left = (.chkPageOf.Left - (4& * lngTpp))
19080       .chkNarrowFont_vline01.Left = (.chkNarrowFont_vline02.Left - lngTpp)

19090       .cmdSelect_Legend_lbl.Left = .cmdSelect_Legend_lbl_alt_line4.Left
19100       .cmdSelect_Legend_box.Left = .cmdSelect_Legend_lbl.Left
19110       .cmdSelect_Legend_Tgl_lbl1.Left = .cmdSelect_Legend_lbl_alt_line5.Left
19120       .cmdSelect_Legend_Tgl_lbl2.Left = .cmdSelect_Legend_Tgl_lbl1.Left
19130       .cmdSelect_Legend_tgl_on_raised_img.Left = .cmdSelect_Legend_lbl_alt_line6.Left
19140       .cmdSelect_Legend_tgl_on_raised_focus_img.Left = .cmdSelect_Legend_tgl_on_raised_img.Left
19150       .cmdSelect_Legend_tgl_on.Left = .cmdSelect_Legend_tgl_on_raised_img.Left
19160       .cmdSelect_Legend_tgl_off_raised_img.Left = .cmdSelect_Legend_lbl_alt_line6.Left
19170       .cmdSelect_Legend_tgl_off_raised_focus_img.Left = .cmdSelect_Legend_tgl_off_raised_img.Left
19180       .cmdSelect_Legend_tgl_off.Left = .cmdSelect_Legend_tgl_off_raised_img.Left
19190       .cmdSelect_Legend_vline01.Left = (.cmdSelect_Legend_Tgl_lbl1.Left + (7& * lngTpp))
19200       .cmdSelect_Legend_vline02.Left = (.cmdSelect_Legend_vline01.Left + lngTpp)

19210     Case False

19220       .chkNarrowFont_box.Left = .chkNarrowFont_box_alt_line.Left
19230       .chkNarrowFont.Left = (.chkNarrowFont_box.Left + (lngNarrowFont_Left - lngNarrowFontBox_Left))
19240       .chkNarrowFont_lbl.Left = (.chkNarrowFont.Left + lngChkBoxLbl_LeftOffset)
19250       .chkIncludeSharefaceTot.Left = .chkNarrowFont.Left
19260       .chkIncludeSharefaceTot_lbl.Left = .chkNarrowFont_lbl.Left
19270       .chkPageOf.Left = .chkNarrowFont.Left + (lngPageOf_Left - lngNarrowFont_Left)
19280       .chkPageOf_lbl.Left = (.chkPageOf.Left + lngChkBoxLbl_LeftOffset)
19290       .chkSaveSizePos.Left = .chkPageOf.Left
19300       .chkSaveSizePos_lbl.Left = .chkPageOf_lbl.Left
19310       .chkNarrowFont_vline02.Left = (.chkPageOf.Left - (4& * lngTpp))
19320       .chkNarrowFont_vline01.Left = (.chkNarrowFont_vline02.Left - lngTpp)

19330       .cmdSelect_Legend_lbl.Left = .cmdSelect_Legend_lbl_alt_line1.Left
19340       .cmdSelect_Legend_box.Left = .cmdSelect_Legend_lbl.Left
19350       .cmdSelect_Legend_Tgl_lbl1.Left = .cmdSelect_Legend_lbl_alt_line2.Left
19360       .cmdSelect_Legend_Tgl_lbl2.Left = .cmdSelect_Legend_Tgl_lbl1.Left
19370       .cmdSelect_Legend_tgl_on_raised_img.Left = .cmdSelect_Legend_lbl_alt_line3.Left
19380       .cmdSelect_Legend_tgl_on_raised_focus_img.Left = .cmdSelect_Legend_tgl_on_raised_img.Left
19390       .cmdSelect_Legend_tgl_on.Left = .cmdSelect_Legend_tgl_on_raised_img.Left
19400       .cmdSelect_Legend_tgl_off_raised_img.Left = .cmdSelect_Legend_lbl_alt_line3.Left
19410       .cmdSelect_Legend_tgl_off_raised_focus_img.Left = .cmdSelect_Legend_tgl_off_raised_img.Left
19420       .cmdSelect_Legend_tgl_off.Left = .cmdSelect_Legend_tgl_off_raised_img.Left
19430       .cmdSelect_Legend_vline01.Left = (.cmdSelect_Legend_Tgl_lbl1.Left + (7& * lngTpp))
19440       .cmdSelect_Legend_vline02.Left = (.cmdSelect_Legend_vline01.Left + lngTpp)

19450       .cmdSelect_Legend_lbl.Visible = True
19460       .cmdSelect_Legend_box.Visible = True
19470       .cmdSelect_Legend_Tgl_lbl1.Visible = True
19480       .cmdSelect_Legend_Tgl_lbl2.Visible = True
19490       .cmdSelect_Legend_tgl_on_raised_img.Visible = True
19500       .cmdSelect_Legend_tgl_on.Visible = True
19510       .cmdSelect_Legend_tgl_off_raised_img.Visible = True
19520       .cmdSelect_Legend_tgl_off.Visible = True
19530       .cmdSelect_Legend_vline01.Visible = True
19540       .cmdSelect_Legend_vline02.Visible = True

19550     End Select
19560   End With

EXITP:
19570   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelect_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_lbl_MouseMove"

19610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19620   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_lbl_DblClick(Cancel As Integer)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_lbl_DblClick"

19710   cmdSelect_lbl3_cmd_DblClick Cancel  ' ** Procedure: Above.

EXITP:
19720   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_lbl2_MouseMove"

19810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19820   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelectAll_GotFocus()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_GotFocus"

19910   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
19920   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseDown"

20010   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
20020   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectAll_Click()

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_Click"

        Dim frm1 As Access.Form, frm2 As Access.Form, ctl As Access.Control
        Dim strTmp01 As String, strTmp02 As String

20110   With Me
20120     Set frm1 = .frmTransaction_Audit_Sub.Form
20130     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
20140     For Each ctl In frm1.Section("FormHeader").Controls  ' ** No need to do frmTransaction_Audit_Sub_ds.
20150       With ctl
20160         If Right(.Name, 4) = "_chk" Then
20170           strTmp01 = .Name
20180           strTmp01 = Left(strTmp01, (Len(strTmp01) - 4))
20190           strTmp01 = strTmp01 & "_lbl"
20200           strTmp02 = strTmp01 & "_line"
20210           .Value = True
20220           frm2.Controls(.Name).Value = True
20230           frm1.Controls(strTmp01).ForeColor = CLR_DKGRY2
20240           frm1.Controls(strTmp02).BorderColor = CLR_DKGRY
20250           Select Case strTmp01
                Case "journalno_lbl", "journaltype_lbl", "transdate_lbl", "accountno_lbl"
20260             frm1.Controls(strTmp01 & "2").ForeColor = CLR_DKGRY2
20270           End Select
20280         End If
20290       End With
20300     Next
20310     frm1.Print_ChkCnt  ' ** Form Function: frmTransaction_Audit_Sub.
20320     frm2.Print_ChkCnt  ' ** Form Function: frmTransaction_Audit_Sub_ds.
20330     PrintTgls_Set Me  ' ** Module Procedure: modTransactionAuditFuncs1.
20340   End With

EXITP:
20350   Set ctl = Nothing
20360   Set frm1 = Nothing
20370   Set frm2 = Nothing
20380   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20420   intShiftDown = (Shift And acShiftMask) > 0
20430   intAltDown = (Shift And acAltMask) > 0
20440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20460     Select Case intRetVal
          Case vbKeyTab
20470       With Me
20480         intRetVal = 0
20490         .cmdSelectNone.SetFocus
20500       End With
20510     End Select
20520   End If

        ' ** Shift keys.
20530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20540     Select Case intRetVal
          Case vbKeyTab
20550       With Me
20560         intRetVal = 0
20570         lngRecsCur = .frmTransaction_Audit_Sub.Form.RecCnt  ' ** Form Function: frmTransaction_Audit_Sub.
20580         If lngRecsCur > 0& Then
20590           Select Case .opgView
                Case .opgView_optForm.OptionValue
20600             .frmTransaction_Audit_Sub.SetFocus
20610             .frmTransaction_Audit_Sub.Form.FocusHolder.SetFocus
20620           Case .opgView_optDatasheet.OptionValue
20630             .frmTransaction_Audit_Sub_ds.SetFocus
20640           End Select
20650         Else
20660           DoCmd.SelectObject acForm, .Parent.Name, False
20670           .Parent.frmTransaction_Audit_Sub_Criteria.SetFocus
20680           .Parent.frmTransaction_Audit_Sub_Criteria.Form.opgHidden.SetFocus
20690         End If
20700       End With
20710     End Select
20720   End If

        ' ** Ctrl keys.
20730   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
20740     Select Case intRetVal
          Case vbKeyTab
20750       With Me
20760         intRetVal = 0
20770         lngRecsCur = .frmTransaction_Audit_Sub.Form.RecCnt  ' ** Form Function: frmTransaction_Audit_Sub.
20780         If lngRecsCur > 0& Then
20790           Select Case .opgView
                Case .opgView_optForm.OptionValue
20800             .frmTransaction_Audit_Sub.SetFocus
20810             .frmTransaction_Audit_Sub.Form.journalno_tgl.SetFocus
20820           Case .opgView_optDatasheet.OptionValue
20830             .frmTransaction_Audit_Sub_ds.SetFocus
20840           End Select
20850         Else
20860           Beep
20870         End If
20880       End With
20890     End Select
20900   End If

EXITP:
20910   KeyCode = intRetVal
20920   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseMove"

21010   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
21020   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseUp"

21110   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
21120   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectAll_LostFocus()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_LostFocus"

21210   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
21220   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectNone_GotFocus()

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_GotFocus"

21310   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
21320   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

21400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseDown"

21410   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
21420   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectNone_Click()

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_Click"

        Dim frm1 As Access.Form, frm2 As Access.Form, ctl As Access.Control
        Dim strTmp01 As String, strTmp02 As String

21510   With Me
21520     Set frm1 = .frmTransaction_Audit_Sub.Form
21530     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
21540     For Each ctl In frm1.Section("FormHeader").Controls  ' ** No need to do frmTransaction_Audit_Sub_ds.
21550       With ctl
21560         If Right(.Name, 4) = "_chk" Then
21570           strTmp01 = .Name
21580           strTmp01 = Left(strTmp01, (Len(strTmp01) - 4))
21590           strTmp01 = strTmp01 & "_lbl"
21600           strTmp02 = strTmp01 & "_line"
21610           .Value = False
21620           frm2.Controls(.Name).Value = False
21630           frm1.Controls(strTmp01).ForeColor = WIN_CLR_DISF
21640           frm1.Controls(strTmp02).BorderColor = CLR_GRY4
21650           Select Case strTmp01
                Case "journalno_lbl", "journaltype_lbl", "transdate_lbl", "accountno_lbl"
21660             frm1.Controls(strTmp01 & "2").ForeColor = WIN_CLR_DISF
21670           End Select
21680         End If
21690       End With
21700     Next
21710     frm1.Print_ChkCnt  ' ** Form Function: frmTransaction_Audit_Sub.
21720     frm2.Print_ChkCnt  ' ** Form Function: frmTransaction_Audit_Sub_ds.
21730     PrintTgls_Set Me  ' ** Module Procedure: modTransactionAuditFuncs1.
21740   End With

EXITP:
21750   Set ctl = Nothing
21760   Set frm1 = Nothing
21770   Set frm2 = Nothing
21780   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21820   intShiftDown = (Shift And acShiftMask) > 0
21830   intAltDown = (Shift And acAltMask) > 0
21840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21860     Select Case intRetVal
          Case vbKeyTab
21870       With Me
21880         intRetVal = 0
21890         If .cmdPreview.Enabled = True Then
21900           .cmdPreview.SetFocus
21910         Else
21920           .cmdClose.SetFocus
21930         End If
21940       End With
21950     End Select
21960   End If

        ' ** Shift keys.
21970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21980     Select Case intRetVal
          Case vbKeyTab
21990       With Me
22000         intRetVal = 0
22010         .cmdSelectAll.SetFocus
22020       End With
22030     End Select
22040   End If

        ' ** Ctrl keys.
22050   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
22060     Select Case intRetVal
          Case vbKeyTab
22070       With Me
22080         intRetVal = 0
22090         .chkNarrowFont.SetFocus
22100       End With
22110     End Select
22120   End If

        ' ** Ctrl-Shift keys.
22130   If intCtrlDown And (Not intAltDown) And intShiftDown Then
22140     Select Case intRetVal
          Case vbKeyTab
22150       With Me
22160         intRetVal = 0
22170         lngRecsCur = .frmTransaction_Audit_Sub.Form.RecCnt  ' ** Form Function: frmTransaction_Audit_Sub.
22180         If lngRecsCur > 0& Then
22190           Select Case .opgView
                Case .opgView_optForm.OptionValue
22200             .frmTransaction_Audit_Sub.SetFocus
22210             .frmTransaction_Audit_Sub.Form.FocusHolder2.SetFocus
22220             .frmTransaction_Audit_Sub.Form.ledger_HIDDEN_tgl.SetFocus
22230           Case .opgView_optDatasheet.OptionValue
22240             .frmTransaction_Audit_Sub_ds.SetFocus
22250           End Select
22260         Else
22270           Beep
22280         End If
22290       End With
22300     End Select
22310   End If

EXITP:
22320   KeyCode = intRetVal
22330   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseMove"

22410   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
22420   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseUp"

22510   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
22520   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub cmdSelectNone_LostFocus()

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_LostFocus"

22610   Select_Handler_TA THIS_PROC, blnSelectAll_Focus, blnSelectAll_MouseDown, _
          blnSelectNone_Focus, blnSelectNone_MouseDown, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
22620   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Public Sub chkNarrowFont_AfterUpdate()

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkNarrowFont_AfterUpdate"

22710   With Me
22720     Select Case .chkNarrowFont
          Case True
22730       .chkNarrowFont_lbl.FontBold = True
22740     Case False
22750       .chkNarrowFont_lbl.FontBold = False
22760     End Select
22770   End With

EXITP:
22780   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub chkNarrowFont_KeyDown(KeyCode As Integer, Shift As Integer)

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkNarrowFont_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22820   intShiftDown = (Shift And acShiftMask) > 0
22830   intAltDown = (Shift And acAltMask) > 0
22840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22860     Select Case intRetVal
          Case vbKeyTab
22870       With Me
22880         intRetVal = 0
22890         .chkIncludeSharefaceTot.SetFocus
22900       End With
22910     End Select
22920   End If

        ' ** Shift keys.
22930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22940     Select Case intRetVal
          Case vbKeyTab
22950       With Me
22960         intRetVal = 0
22970         .cmdSelectNone.SetFocus
22980       End With
22990     End Select
23000   End If

EXITP:
23010   KeyCode = intRetVal
23020   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub chkNarrowFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkNarrowFont_MouseMove"

23110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23120   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Public Sub chkIncludeSharefaceTot_AfterUpdate()

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeSharefaceTot_AfterUpdate"

23210   With Me
23220     Select Case .chkIncludeSharefaceTot
          Case True
23230       .chkIncludeSharefaceTot_lbl.FontBold = True
23240     Case False
23250       .chkIncludeSharefaceTot_lbl.FontBold = False
23260     End Select
23270   End With

EXITP:
23280   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Private Sub chkIncludeSharefaceTot_KeyDown(KeyCode As Integer, Shift As Integer)

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeSharefaceTot_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23320   intShiftDown = (Shift And acShiftMask) > 0
23330   intAltDown = (Shift And acAltMask) > 0
23340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23360     Select Case intRetVal
          Case vbKeyTab
23370       With Me
23380         intRetVal = 0
23390         .chkPageOf.SetFocus
23400       End With
23410     End Select
23420   End If

        ' ** Shift keys.
23430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23440     Select Case intRetVal
          Case vbKeyTab
23450       With Me
23460         intRetVal = 0
23470         .chkNarrowFont.SetFocus
23480       End With
23490     End Select
23500   End If

EXITP:
23510   KeyCode = intRetVal
23520   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub chkIncludeSharefaceTot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeSharefaceTot_MouseMove"

23610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23620   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Public Sub chkPageOf_AfterUpdate()

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

23710   With Me
23720     Select Case .chkPageOf
          Case True
23730       .chkPageOf_lbl.FontBold = True
23740     Case False
23750       .chkPageOf_lbl.FontBold = False
23760     End Select
23770   End With

EXITP:
23780   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

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
23890         .chkSaveSizePos.SetFocus
23900       End With
23910     End Select
23920   End If

        ' ** Shift keys.
23930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23940     Select Case intRetVal
          Case vbKeyTab
23950       With Me
23960         intRetVal = 0
23970         .chkIncludeSharefaceTot.SetFocus
23980       End With
23990     End Select
24000   End If

EXITP:
24010   KeyCode = intRetVal
24020   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub chkSaveSizePos_AfterUpdate()

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSaveSizePos_AfterUpdate"

24110   With Me
24120     Select Case .chkSaveSizePos
          Case True
24130       .chkSaveSizePos_lbl.FontBold = True
24140     Case False
24150       .chkSaveSizePos_lbl.FontBold = False
24160     End Select
24170   End With

EXITP:
24180   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub chkSaveSizePos_KeyDown(KeyCode As Integer, Shift As Integer)

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkSaveSizePos_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24220   intShiftDown = (Shift And acShiftMask) > 0
24230   intAltDown = (Shift And acAltMask) > 0
24240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24260     Select Case intRetVal
          Case vbKeyTab
24270       With Me
24280         intRetVal = 0
24290         If .ckgFlds_chkJournalno.Visible = True And .ckgFlds_chkJournalno.Enabled = True Then
24300           .ckgFlds_chkJournalno.SetFocus
24310         ElseIf .cmdPreview.Enabled = True Then
24320           .cmdPreview.SetFocus
24330         Else
24340           .cmdClose.SetFocus
24350         End If
24360       End With
24370     End Select
24380   End If

        ' ** Shift keys.
24390   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24400     Select Case intRetVal
          Case vbKeyTab
24410       With Me
24420         intRetVal = 0
24430         .chkPageOf.SetFocus
24440       End With
24450     End Select
24460   End If

EXITP:
24470   KeyCode = intRetVal
24480   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_box_DblClick(Cancel As Integer)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_box_DblClick"

24510   cmdSelect_lbl3_cmd_DblClick Cancel  ' ** Procedure: Above.

EXITP:
24520   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_lbl_DblClick(Cancel As Integer)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_lbl_DblClick"

24610   cmdSelect_lbl3_cmd_DblClick Cancel  ' ** Procedure: Above.

EXITP:
24620   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_lbl_MouseMove"

24710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
24720   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_Tgl_lbl1_DblClick(Cancel As Integer)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_Tgl_lbl1_DblClick"

24810   cmdSelect_lbl3_cmd_DblClick Cancel  ' ** Procedure: Above.

EXITP:
24820   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_tgl_on_DblClick(Cancel As Integer)

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_tgl_on_DblClick"

24910   With Me
24920     If .cmdSelectAll.Enabled = True Then
24930       .cmdSelectAll.SetFocus
24940     ElseIf .chkNarrowFont.Enabled = True Then
24950       .chkNarrowFont.SetFocus
24960     Else
24970       .cmdClose.SetFocus
24980     End If
24990     cmdSelect_lbl3_cmd_DblClick Cancel  ' ** Procedure: Above.
25000   End With

EXITP:
25010   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_tgl_on_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_tgl_on_MouseMove"

25110   With Me
25120     .cmdSelect_Legend_tgl_on_raised_focus_img.Visible = True
25130     .cmdSelect_Legend_tgl_on_raised_img.Visible = False
25140   End With

EXITP:
25150   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_Tgl_lbl2_DblClick(Cancel As Integer)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_Tgl_lbl2_DblClick"

25210   cmdSelect_lbl3_cmd_DblClick Cancel  ' ** Procedure: Above.

EXITP:
25220   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_tgl_off_DblClick(Cancel As Integer)

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_tgl_off_DblClick"

25310   With Me
25320     If .cmdSelectNone.Enabled = True Then
25330       .cmdSelectNone.SetFocus
25340     ElseIf .chkIncludeSharefaceTot.Enabled = True Then
25350       .chkIncludeSharefaceTot.SetFocus
25360     Else
25370       .cmdClose.SetFocus
25380     End If
25390     cmdSelect_lbl3_cmd_DblClick Cancel  ' ** Procedure: Above.
25400   End With

EXITP:
25410   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub cmdSelect_Legend_tgl_off_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelect_Legend_tgl_off_MouseMove"

25510   With Me
25520     .cmdSelect_Legend_tgl_off_raised_focus_img.Visible = True
25530     .cmdSelect_Legend_tgl_off_raised_img.Visible = False
25540   End With

EXITP:
25550   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_box_MouseMove"

25610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
25620   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_cmd_GotFocus()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_cmd_GotFocus"

25710   CkgFlds_Handler_TA THIS_PROC, blnCkgFlds_Focus, blnCkgFlds_MouseDown, _
          blnFromCkgFlds, lngMonitorCnt, lngMonitorNum, lngCkgFldsBox_Offset, _
          lngChkBoxLbl_TopOffset, lngCkgFldsBox_Height, lngSelectBox_Top, _
          lngCkgFldsVLine_Top, lngCkgFldsVLine_Height, lngFrmFlds, arr_varFrmFld, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
25720   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_cmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_cmd_MouseDown"

25810   CkgFlds_Handler_TA THIS_PROC, blnCkgFlds_Focus, blnCkgFlds_MouseDown, _
          blnFromCkgFlds, lngMonitorCnt, lngMonitorNum, lngCkgFldsBox_Offset, _
          lngChkBoxLbl_TopOffset, lngCkgFldsBox_Height, lngSelectBox_Top, _
          lngCkgFldsVLine_Top, lngCkgFldsVLine_Height, lngFrmFlds, arr_varFrmFld, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
25820   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_cmd_Click()

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_cmd_Click"

25910   CkgFlds_Handler_TA THIS_PROC, blnCkgFlds_Focus, blnCkgFlds_MouseDown, _
          blnFromCkgFlds, lngMonitorCnt, lngMonitorNum, lngCkgFldsBox_Offset, _
          lngChkBoxLbl_TopOffset, lngCkgFldsBox_Height, lngSelectBox_Top, _
          lngCkgFldsVLine_Top, lngCkgFldsVLine_Height, lngFrmFlds, arr_varFrmFld, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
25920   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_cmd_KeyDown(KeyCode As Integer, Shift As Integer)

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_cmd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26020   intShiftDown = (Shift And acShiftMask) > 0
26030   intAltDown = (Shift And acAltMask) > 0
26040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26060     Select Case intRetVal
          Case vbKeyTab
26070       With Me
26080         intRetVal = 0
26090         If .ckgFlds_chkJournalno.Visible = True And .ckgFlds_chkJournalno.Enabled = True Then
26100           .ckgFlds_chkJournalno.SetFocus
26110         ElseIf .cmdPreview.Enabled = True Then
26120           .cmdPreview.SetFocus
26130         Else
26140           .cmdClose.SetFocus
26150         End If
26160       End With
26170     End Select
26180   End If

        ' ** Shift keys.
26190   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26200     Select Case intRetVal
          Case vbKeyTab
26210       With Me
26220         intRetVal = 0
26230         .chkSaveSizePos.SetFocus
26240       End With
26250     End Select
26260   End If

EXITP:
26270   KeyCode = intRetVal
26280   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub ckgFlds_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_cmd_MouseMove"

26310   CkgFlds_Handler_TA THIS_PROC, blnCkgFlds_Focus, blnCkgFlds_MouseDown, _
          blnFromCkgFlds, lngMonitorCnt, lngMonitorNum, lngCkgFldsBox_Offset, _
          lngChkBoxLbl_TopOffset, lngCkgFldsBox_Height, lngSelectBox_Top, _
          lngCkgFldsVLine_Top, lngCkgFldsVLine_Height, lngFrmFlds, arr_varFrmFld, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
26320   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_cmd_MouseUp"

26410   CkgFlds_Handler_TA THIS_PROC, blnCkgFlds_Focus, blnCkgFlds_MouseDown, _
          blnFromCkgFlds, lngMonitorCnt, lngMonitorNum, lngCkgFldsBox_Offset, _
          lngChkBoxLbl_TopOffset, lngCkgFldsBox_Height, lngSelectBox_Top, _
          lngCkgFldsVLine_Top, lngCkgFldsVLine_Height, lngFrmFlds, arr_varFrmFld, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
26420   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_cmd_LostFocus()

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_cmd_LostFocus"

26510   CkgFlds_Handler_TA THIS_PROC, blnCkgFlds_Focus, blnCkgFlds_MouseDown, _
          blnFromCkgFlds, lngMonitorCnt, lngMonitorNum, lngCkgFldsBox_Offset, _
          lngChkBoxLbl_TopOffset, lngCkgFldsBox_Height, lngSelectBox_Top, _
          lngCkgFldsVLine_Top, lngCkgFldsVLine_Height, lngFrmFlds, arr_varFrmFld, Me  ' ** Module Procedure: modTransactionAuditFuncs2.

EXITP:
26520   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkJournalno_AfterUpdate()

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkJournalno_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

26610   With Me
26620     blnContinue = True
26630     Set frm1 = .frmTransaction_Audit_Sub.Form
26640     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
26650     If .ckgFlds_chkJournalno = True Then
26660       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
26670         blnContinue = False
26680         Beep
26690         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
26700         .ckgFlds_chkJournalno = False
26710       End If
26720     End If

26730     If blnContinue = True Then
26740       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
26750       Select Case .Controls(strCtlName)
            Case True
26760         .Controls(strCtlName & "_lbl").FontBold = True
26770         frmCrit.journalno.Enabled = True
26780         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
26790         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
26800       Case False
26810         .Controls(strCtlName & "_lbl").FontBold = False
26820         frmCrit.FocusHolder.SetFocus
26830         frmCrit.journalno = Null
26840         frmCrit.journalno.Enabled = False
26850         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
26860         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
26870       End Select
26880       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the left where the field is.
26890         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
26900       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the left where the field is.
26910         fSetScrollBarPosHZ frm2, 1&  ' ** Module Function: modScrollBarFuncs.
26920       End Select
26930     End If  ' ** blnContinue.
26940     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
26950   End With

EXITP:
26960   Set frm1 = Nothing
26970   Set frm2 = Nothing
26980   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkJournalno_KeyDown(KeyCode As Integer, Shift As Integer)

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkJournalno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27020   intShiftDown = (Shift And acShiftMask) > 0
27030   intAltDown = (Shift And acAltMask) > 0
27040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27060     Select Case intRetVal
          Case vbKeyTab
27070       With Me
27080         intRetVal = 0
27090         .ckgFlds_chkJournalType.SetFocus
27100       End With
27110     End Select
27120   End If

        ' ** Shift keys.
27130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27140     Select Case intRetVal
          Case vbKeyTab
27150       With Me
27160         intRetVal = 0
27170         .chkSaveSizePos.SetFocus
27180       End With
27190     End Select
27200   End If

EXITP:
27210   KeyCode = intRetVal
27220   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Private Sub ckgFlds_chkJournalno_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkJournalno_MouseMove"

27310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
27320   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkJournalType_AfterUpdate()

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkJournalType_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

27410   With Me
27420     blnContinue = True
27430     Set frm1 = .frmTransaction_Audit_Sub.Form
27440     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
27450     If .ckgFlds_chkJournalType = True Then
27460       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
27470         blnContinue = False
27480         Beep
27490         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
27500         .ckgFlds_chkJournalType = False
27510       End If
27520     End If
27530     If blnContinue = True Then
27540       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
27550       Select Case .Controls(strCtlName)
            Case True
27560         .Controls(strCtlName & "_lbl").FontBold = True
27570         frmCrit.cmbJournalType1.Enabled = True
27580         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
27590         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
27600       Case False
27610         .Controls(strCtlName & "_lbl").FontBold = False
27620         frmCrit.FocusHolder.SetFocus
27630         frmCrit.cmbJournalType1.Enabled = False
27640         frmCrit.cmbJournalType1 = Null
27650         frmCrit.cmbJournalType2 = Null
27660         frmCrit.cmbJournalType3 = Null
27670         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
27680         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
27690       End Select
27700       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
27710         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
27720       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
27730         fSetScrollBarPosHZ frm2, 1&  ' ** Module Function: modScrollBarFuncs.
27740       End Select
27750     End If  ' ** blnContinue.
27760     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
27770   End With

EXITP:
27780   Set frm1 = Nothing
27790   Set frm2 = Nothing
27800   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkTransDate_AfterUpdate()

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkTransDate_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

27910   With Me
27920     blnContinue = True
27930     Set frm1 = .frmTransaction_Audit_Sub.Form
27940     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
27950     If .ckgFlds_chkTransDate = True Then
27960       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
27970         blnContinue = False
27980         Beep
27990         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
28000         .ckgFlds_chkTransDate = False
28010       End If
28020     End If
28030     If blnContinue = True Then
28040       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
28050       Select Case .Controls(strCtlName)
            Case True
28060         .Controls(strCtlName & "_lbl").FontBold = True
28070         frmCrit.TransDateStart.Enabled = True
28080         frmCrit.TransDateEnd.Enabled = True
28090         frmCrit.cmdCalendar1.Enabled = True
28100         frmCrit.Calendar_Handler frmCrit.cmdCalendar1.Name & "_Disable", True  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
28110         frmCrit.cmdCalendar2.Enabled = True
28120         frmCrit.Calendar_Handler frmCrit.cmdCalendar2.Name & "_Disable", True  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
28130         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
28140         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
28150       Case False
28160         .Controls(strCtlName & "_lbl").FontBold = False
28170         frmCrit.FocusHolder.SetFocus
28180         frmCrit.TransDateStart.Enabled = False
28190         frmCrit.TransDateEnd.Enabled = False
28200         frmCrit.TransDateStart = Null
28210         frmCrit.TransDateEnd = Null
28220         frmCrit.cmdCalendar1.Enabled = False
28230         frmCrit.Calendar_Handler frmCrit.cmdCalendar1.Name & "_Disable", False  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
28240         frmCrit.cmdCalendar2.Enabled = False
28250         frmCrit.Calendar_Handler frmCrit.cmdCalendar2.Name & "_Disable", False  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
28260         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
28270         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
28280       End Select
28290       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
28300         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
28310       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
28320         fSetScrollBarPosHZ frm2, 1&  ' ** Module Function: modScrollBarFuncs.
28330       End Select
28340     End If  ' ** blnContinue.
28350     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
28360   End With

EXITP:
28370   Set frm1 = Nothing
28380   Set frm2 = Nothing
28390   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkAccountNo_AfterUpdate()

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkAccountNo_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

28410   With Me
28420     blnContinue = True
28430     Set frm1 = .frmTransaction_Audit_Sub.Form
28440     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
28450     If .ckgFlds_chkAccountNo = True Then
28460       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
28470         blnContinue = False
28480         Beep
28490         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
28500         .ckgFlds_chkAccountNo = False
28510       End If
28520     End If
28530     If blnContinue = True Then
28540       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
28550       Select Case .Controls(strCtlName)
            Case True
28560         .Controls(strCtlName & "_lbl").FontBold = True
28570         frmCrit.cmbAccounts.Enabled = True
28580         frmCrit.opgAccountSource.Enabled = True
28590         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
28600         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
28610       Case False
28620         .Controls(strCtlName & "_lbl").FontBold = False
28630         Select Case .ckgFlds_chkShortName
              Case True
28640           frmCrit.cmbAccounts.Enabled = True
28650           frmCrit.opgAccountSource.Enabled = True
28660           ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
28670         Case False
28680           frmCrit.FocusHolder.SetFocus
28690           frmCrit.cmbAccounts.Enabled = False
28700           frmCrit.cmbAccounts = Null
28710           frmCrit.opgAccountSource.Enabled = False
28720           ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
28730         End Select
28740         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
28750       End Select
28760       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
28770         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
28780       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
28790         fSetScrollBarPosHZ frm2, 1& ' ** Module Function: modScrollBarFuncs.
28800       End Select
28810     End If  ' ** blnContinue.
28820     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
28830   End With

EXITP:
28840   Set frm1 = Nothing
28850   Set frm2 = Nothing
28860   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkAccountNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkAccountNo_MouseMove"

28910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
28920   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkShortName_AfterUpdate()

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkShortName_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

29010   With Me
29020     blnContinue = True
29030     Set frm1 = .frmTransaction_Audit_Sub.Form
29040     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
29050     If .ckgFlds_chkShortName = True Then
29060       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
29070         blnContinue = False
29080         Beep
29090         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
29100         .ckgFlds_chkShortName = False
29110       End If
29120     End If
29130     If blnContinue = True Then
29140       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
29150       Select Case .Controls(strCtlName)
            Case True
29160         .Controls(strCtlName & "_lbl").FontBold = True
29170         frmCrit.cmbAccounts.Enabled = True
29180         frmCrit.opgAccountSource.Enabled = True
29190         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
29200         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
29210       Case False
29220         .Controls(strCtlName & "_lbl").FontBold = False
29230         Select Case .ckgFlds_chkAccountNo
              Case True
29240           frmCrit.cmbAccounts.Enabled = True
29250           frmCrit.opgAccountSource.Enabled = True
29260           ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
29270         Case False
29280           frmCrit.FocusHolder.SetFocus
29290           frmCrit.cmbAccounts.Enabled = False
29300           frmCrit.cmbAccounts = Null
29310           frmCrit.opgAccountSource.Enabled = False
29320           ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
29330         End Select
29340         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
29350       End Select
29360       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
29370         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
29380       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
29390         fSetScrollBarPosHZ frm2, 1& ' ** Module Function: modScrollBarFuncs.
29400       End Select
29410     End If  ' ** blnContinue.
29420     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
29430   End With

EXITP:
29440   Set frm1 = Nothing
29450   Set frm2 = Nothing
29460   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkCusip_AfterUpdate()
' ** There is no CUSIP search.

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkCusip_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

29510   With Me
29520     blnContinue = True
29530     Set frm1 = .frmTransaction_Audit_Sub.Form
29540     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
29550     If .ckgFlds_chkCusip = True Then
29560       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
29570         blnContinue = False
29580         Beep
29590         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
29600         .ckgFlds_chkCusip = False
29610       End If
29620     End If
29630     If blnContinue = True Then
29640       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
29650       Select Case .Controls(strCtlName)
            Case True
29660         .Controls(strCtlName & "_lbl").FontBold = True
29670         frmCrit.cmbAssets.Enabled = True
29680         frmCrit.opgAssetSource.Enabled = True
29690         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
29700         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
29710       Case False
29720         .Controls(strCtlName & "_lbl").FontBold = False
29730         Select Case .ckgFlds_chkAssetDescription
              Case True
29740           frmCrit.cmbAssets.Enabled = True
29750           frmCrit.opgAssetSource.Enabled = True
29760           ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
29770         Case False
29780           frmCrit.FocusHolder.SetFocus
29790           frmCrit.cmbAssets.Enabled = False
29800           frmCrit.opgAssetSource.Enabled = False
29810           frmCrit.cmbAssets = Null
29820           ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
29830         End Select
29840         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
29850       End Select
29860       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
29870         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
29880       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
29890         fSetScrollBarPosHZ frm2, 1& ' ** Module Function: modScrollBarFuncs.
29900       End Select
29910     End If  ' ** blnContinue.
29920     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
29930   End With

EXITP:
29940   Set frm1 = Nothing
29950   Set frm2 = Nothing
29960   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkAssetDescription_AfterUpdate()

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkAssetDescription_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

30010   With Me
30020     blnContinue = True
30030     Set frm1 = .frmTransaction_Audit_Sub.Form
30040     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
30050     If .ckgFlds_chkAssetDescription = True Then
30060       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
30070         blnContinue = False
30080         Beep
30090         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
30100         .ckgFlds_chkAssetDescription = False
30110       End If
30120     End If
30130     If blnContinue = True Then
30140       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
30150       Select Case .Controls(strCtlName)
            Case True
30160         .Controls(strCtlName & "_lbl").FontBold = True
30170         frmCrit.cmbAssets.Enabled = True
30180         frmCrit.opgAssetSource.Enabled = True
30190         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
30200         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
30210       Case False
30220         .Controls(strCtlName & "_lbl").FontBold = False
30230         Select Case .ckgFlds_chkCusip
              Case True
30240           frmCrit.cmbAssets.Enabled = True
30250           frmCrit.opgAssetSource.Enabled = True
30260           ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
30270         Case False
30280           frmCrit.FocusHolder.SetFocus
30290           frmCrit.cmbAssets.Enabled = False
30300           frmCrit.opgAssetSource.Enabled = False
30310           frmCrit.cmbAssets = Null
30320           ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
30330         End Select
30340         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
30350       End Select
30360       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
30370         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
30380       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
30390         fSetScrollBarPosHZ frm2, 1& ' ** Module Function: modScrollBarFuncs.
30400       End Select
30410     End If  ' ** blnContinue.
30420     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
30430   End With

EXITP:
30440   Set frm1 = Nothing
30450   Set frm2 = Nothing
30460   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkShareFace_AfterUpdate()
' ** There is no shareface search.

30500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkShareFace_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

30510   With Me
30520     blnContinue = True
30530     Set frm1 = .frmTransaction_Audit_Sub.Form
30540     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
30550     If .ckgFlds_chkShareFace = True Then
30560       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
30570         blnContinue = False
30580         Beep
30590         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
30600         .ckgFlds_chkShareFace = False
30610       End If
30620     End If
30630     If blnContinue = True Then
30640       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
30650       Select Case .Controls(strCtlName)
            Case True
30660         .Controls(strCtlName & "_lbl").FontBold = True
30670         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
30680         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
30690       Case False
30700         .Controls(strCtlName & "_lbl").FontBold = False
30710         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
30720         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
30730       End Select
30740       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
30750         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
30760       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
30770         fSetScrollBarPosHZ frm2, 1& ' ** Module Function: modScrollBarFuncs.
30780       End Select
30790     End If  ' ** blnContinue.
30800     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
30810   End With

EXITP:
30820   Set frm1 = Nothing
30830   Set frm2 = Nothing
30840   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkICash_AfterUpdate()
' ** There is no icash search.

30900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkICash_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

30910   With Me
30920     blnContinue = True
30930     Set frm1 = .frmTransaction_Audit_Sub.Form
30940     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
30950     If .ckgFlds_chkICash = True Then
30960       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
30970         blnContinue = False
30980         Beep
30990         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
31000         .ckgFlds_chkICash = False
31010       End If
31020     End If
31030     If blnContinue = True Then
31040       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
31050       Select Case .Controls(strCtlName)
            Case True
31060         .Controls(strCtlName & "_lbl").FontBold = True
31070         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
31080         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
31090       Case False
31100         .Controls(strCtlName & "_lbl").FontBold = False
31110         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
31120         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
31130       End Select
31140       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
31150         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
31160       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
31170         fSetScrollBarPosHZ frm2, 1& ' ** Module Function: modScrollBarFuncs.
31180       End Select
31190     End If  ' ** blnContinue.
31200     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
31210   End With

EXITP:
31220   Set frm1 = Nothing
31230   Set frm2 = Nothing
31240   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkPCash_AfterUpdate()
' ** There is no pcash search.

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkPCash_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

31310   With Me
31320     blnContinue = True
31330     Set frm1 = .frmTransaction_Audit_Sub.Form
31340     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
31350     If .ckgFlds_chkPCash = True Then
31360       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
31370         blnContinue = False
31380         Beep
31390         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
31400         .ckgFlds_chkPCash = False
31410       End If
31420     End If
31430     If blnContinue = True Then
31440       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
31450       Select Case .Controls(strCtlName)
            Case True
31460         .Controls(strCtlName & "_lbl").FontBold = True
31470         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
31480         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
31490       Case False
31500         .Controls(strCtlName & "_lbl").FontBold = False
31510         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
31520         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
31530       End Select
31540       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
31550         fSetScrollBarPosHZ frm1, 1&  ' ** Module Function: modScrollBarFuncs.
31560       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
31570         fSetScrollBarPosHZ frm2, 1& ' ** Module Function: modScrollBarFuncs.
31580       End Select
31590     End If  ' ** blnContinue.
31600     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
31610   End With

EXITP:
31620   Set frm1 = Nothing
31630   Set frm2 = Nothing
31640   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkCost_AfterUpdate()
' ** There is no cost search.

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkCost_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

31710   With Me
31720     blnContinue = True
31730     Set frm1 = .frmTransaction_Audit_Sub.Form
31740     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
31750     If .ckgFlds_chkCost = True Then
31760       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
31770         blnContinue = False
31780         Beep
31790         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
31800         .ckgFlds_chkCost = False
31810       End If
31820     End If
31830     If blnContinue = True Then
31840       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
31850       Select Case .Controls(strCtlName)
            Case True
31860         .Controls(strCtlName & "_lbl").FontBold = True
31870         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
31880         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
31890       Case False
31900         .Controls(strCtlName & "_lbl").FontBold = False
31910         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
31920         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
31930       End Select
31940       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
31950         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
31960       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
31970         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
31980       End Select
31990     End If  ' ** blnContinue.
32000     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
32010   End With

EXITP:
32020   Set frm1 = Nothing
32030   Set frm2 = Nothing
32040   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkCurrID_AfterUpdate()

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkCurrID_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

32110   With Me
32120     blnContinue = True
32130     Set frm1 = .frmTransaction_Audit_Sub.Form
32140     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
32150     If .ckgFlds_chkCurrID = True Then
32160       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
32170         blnContinue = False
32180         Beep
32190         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
32200         .ckgFlds_chkCurrID = False
32210       End If
32220     End If
32230     If blnContinue = True Then
32240       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
32250       Select Case .Controls(strCtlName)
            Case True
32260         .Controls(strCtlName & "_lbl").FontBold = True
32270         frmCrit.cmbCurrencies.Enabled = True
32280         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
32290         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
32300       Case False
32310         .Controls(strCtlName & "_lbl").FontBold = False
32320         frmCrit.FocusHolder2.SetFocus
32330         frmCrit.cmbCurrencies.Enabled = False
32340         frmCrit.cmbCurrencies = 150&  ' ** Default to USD.
32350         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
32360         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
32370       End Select
32380       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
32390         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
32400       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
32410         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
32420       End Select
32430     End If  ' ** blnContinue.
32440     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
32450   End With

EXITP:
32460   Set frm1 = Nothing
32470   Set frm2 = Nothing
32480   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkAssetDate_AfterUpdate()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkAssetDate_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

32510   With Me
32520     blnContinue = True
32530     Set frm1 = .frmTransaction_Audit_Sub.Form
32540     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
32550     If .ckgFlds_chkAssetDate = True Then
32560       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
32570         blnContinue = False
32580         Beep
32590         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
32600         .ckgFlds_chkAssetDate = False
32610       End If
32620     End If
32630     If blnContinue = True Then
32640       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
32650       Select Case .Controls(strCtlName)
            Case True
32660         .Controls(strCtlName & "_lbl").FontBold = True
32670         frmCrit.AssetDateStart.Enabled = True
32680         frmCrit.AssetDateEnd.Enabled = True
32690         frmCrit.cmdCalendar3.Enabled = True
32700         frmCrit.Calendar_Handler frmCrit.cmdCalendar3.Name & "_Disable", True  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
32710         frmCrit.cmdCalendar4.Enabled = True
32720         frmCrit.Calendar_Handler frmCrit.cmdCalendar4.Name & "_Disable", True  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
32730         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
32740         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
32750       Case False
32760         .Controls(strCtlName & "_lbl").FontBold = False
32770         frmCrit.FocusHolder3.SetFocus
32780         frmCrit.AssetDateStart.Enabled = False
32790         frmCrit.AssetDateEnd.Enabled = False
32800         frmCrit.AssetDateStart = Null
32810         frmCrit.AssetDateEnd = Null
32820         frmCrit.cmdCalendar3.Enabled = False
32830         frmCrit.Calendar_Handler frmCrit.cmdCalendar3.Name & "_Disable", False  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
32840         frmCrit.cmdCalendar4.Enabled = False
32850         frmCrit.Calendar_Handler frmCrit.cmdCalendar4.Name & "_Disable", False  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
32860         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
32870         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
32880       End Select
32890       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
32900         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
32910       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
32920         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
32930       End Select
32940     End If  ' ** blnContinue.
32950     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
32960   End With

EXITP:
32970   Set frm1 = Nothing
32980   Set frm2 = Nothing
32990   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkPurchaseDate_AfterUpdate()

33000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkPurchaseDate_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

33010   With Me
33020     blnContinue = True
33030     Set frm1 = .frmTransaction_Audit_Sub.Form
33040     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
33050     If .ckgFlds_chkPurchaseDate = True Then
33060       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
33070         blnContinue = False
33080         Beep
33090         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
33100         .ckgFlds_chkPurchaseDate = False
33110       End If
33120     End If
33130     If blnContinue = True Then
33140       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
33150       Select Case .Controls(strCtlName)
            Case True
33160         .Controls(strCtlName & "_lbl").FontBold = True
33170         frmCrit.PurchaseDateStart.Enabled = True
33180         frmCrit.PurchaseDateEnd.Enabled = True
33190         frmCrit.cmdCalendar5.Enabled = True
33200         frmCrit.Calendar_Handler frmCrit.cmdCalendar5.Name & "_Disable", True  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
33210         frmCrit.cmdCalendar6.Enabled = True
33220         frmCrit.Calendar_Handler frmCrit.cmdCalendar6.Name & "_Disable", True  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
33230         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
33240         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
33250       Case False
33260         .Controls(strCtlName & "_lbl").FontBold = False
33270         frmCrit.FocusHolder3.SetFocus
33280         frmCrit.PurchaseDateStart.Enabled = False
33290         frmCrit.PurchaseDateEnd.Enabled = False
33300         frmCrit.PurchaseDateStart = Null
33310         frmCrit.PurchaseDateEnd = Null
33320         frmCrit.cmdCalendar5.Enabled = False
33330         frmCrit.Calendar_Handler frmCrit.cmdCalendar5.Name & "_Disable", False  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
33340         frmCrit.cmdCalendar6.Enabled = False
33350         frmCrit.Calendar_Handler frmCrit.cmdCalendar6.Name & "_Disable", False  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
33360         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
33370         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
33380       End Select
33390       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
33400         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
33410       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
33420         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
33430       End Select
33440     End If  ' ** blnContinue.
33450     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
33460   End With

EXITP:
33470   Set frm1 = Nothing
33480   Set frm2 = Nothing
33490   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkLedgerDescription_AfterUpdate()

33500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkLedgerDescription_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

33510   With Me
33520     blnContinue = True
33530     Set frm1 = .frmTransaction_Audit_Sub.Form
33540     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
33550     If .ckgFlds_chkLedgerDescription = True Then
33560       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
33570         blnContinue = False
33580         Beep
33590         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
33600         .ckgFlds_chkLedgerDescription = False
33610       End If
33620     End If
33630     If blnContinue = True Then
33640       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
33650       Select Case .Controls(strCtlName)
            Case True
33660         .Controls(strCtlName & "_lbl").FontBold = True
33670         frmCrit.ledger_description.Enabled = True
33680         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
33690         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
33700       Case False
33710         .Controls(strCtlName & "_lbl").FontBold = False
33720         frmCrit.FocusHolder3.SetFocus
33730         frmCrit.ledger_description.Enabled = False
33740         frmCrit.ledger_description = Null
33750         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
33760         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
33770       End Select
33780       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
33790         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
33800       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
33810         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
33820       End Select
33830     End If  ' ** blnContinue.
33840     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
33850   End With

EXITP:
33860   Set frm1 = Nothing
33870   Set frm2 = Nothing
33880   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkRecurringItem_AfterUpdate()

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkRecurringItem_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

33910   With Me
33920     blnContinue = True
33930     Set frm1 = .frmTransaction_Audit_Sub.Form
33940     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
33950     If .ckgFlds_chkRecurringItem = True Then
33960       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
33970         blnContinue = False
33980         Beep
33990         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
34000         .ckgFlds_chkRecurringItem = False
34010       End If
34020     End If
34030     If blnContinue = True Then
34040       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
34050       Select Case .Controls(strCtlName)
            Case True
34060         .Controls(strCtlName & "_lbl").FontBold = True
34070         frmCrit.cmbRecurringItems.Enabled = True
34080         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
34090         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
34100       Case False
34110         .Controls(strCtlName & "_lbl").FontBold = False
34120         frmCrit.FocusHolder3.SetFocus
34130         frmCrit.cmbRecurringItems.Enabled = False
34140         frmCrit.cmbRecurringItems = Null
34150         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
34160         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
34170       End Select
34180       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
34190         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
34200       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
34210         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
34220       End Select
34230     End If  ' ** blnContinue.
34240     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
34250   End With

EXITP:
34260   Set frm1 = Nothing
34270   Set frm2 = Nothing
34280   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkRevCodeDesc_AfterUpdate()

34300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkRevCodeDesc_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

34310   With Me
34320     blnContinue = True
34330     Set frm1 = .frmTransaction_Audit_Sub.Form
34340     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
34350     If .ckgFlds_chkRevCodeDesc = True Then
34360       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
34370         blnContinue = False
34380         Beep
34390         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
34400         .ckgFlds_chkRevCodeDesc = False
34410       End If
34420     End If
34430     If blnContinue = True Then
34440       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
34450       Select Case .Controls(strCtlName)
            Case True
34460         .Controls(strCtlName & "_lbl").FontBold = True
34470         frmCrit.cmbRevenueCodes.Enabled = True
34480         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
34490         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
34500       Case False
34510         .Controls(strCtlName & "_lbl").FontBold = False
34520         frmCrit.FocusHolder2.SetFocus
34530         frmCrit.cmbRevenueCodes.Enabled = False
34540         frmCrit.cmbRevenueCodes = Null
34550         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
34560         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
34570       End Select
34580       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
34590         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
34600       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
34610         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
34620       End Select
34630     End If  ' ** blnContinue.
34640     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
34650   End With

EXITP:
34660   Set frm1 = Nothing
34670   Set frm2 = Nothing
34680   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkRevCodeTypeDescription_AfterUpdate()

34700 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkRevCodeTypeDescription_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

34710   With Me
34720     blnContinue = True
34730     Set frm1 = .frmTransaction_Audit_Sub.Form
34740     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
34750     If .ckgFlds_chkRevCodeTypeDescription = True Then
34760       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
34770         blnContinue = False
34780         Beep
34790         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
34800         .ckgFlds_chkRevCodeTypeDescription = False
34810       End If
34820     End If
34830     If blnContinue = True Then
34840       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
34850       Select Case .Controls(strCtlName)
            Case True
34860         .Controls(strCtlName & "_lbl").FontBold = True
34870         frmCrit.chkRevcodeType_Income.Enabled = True
34880         frmCrit.chkRevcodeType_Expense.Enabled = True
34890         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
34900         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
34910       Case False
34920         .Controls(strCtlName & "_lbl").FontBold = False
34930         frmCrit.FocusHolder2.SetFocus
34940         If frmCrit.chkRevcodeType_Income = False Then
34950           frmCrit.chkRevcodeType_Income = True
                'chkRevcodeType_Income_AfterUpdate 3  NOT HIT ON OPEN!
34960           frmCrit.chkRevcodeType_Income_AfterUpdate  ' ** Procedure: Above.
34970         End If
34980         frmCrit.chkRevcodeType_Income.Enabled = False
34990         If frmCrit.chkRevcodeType_Expense = False Then
35000           frmCrit.chkRevcodeType_Expense = True
                'chkRevcodeType_Expense_AfterUpdate 3  NOT HIT ON OPEN!
35010           frmCrit.chkRevcodeType_Expense_AfterUpdate  ' ** Procedure: Above.
35020         End If
35030         frmCrit.chkRevcodeType_Expense.Enabled = False
35040         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
35050         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
35060       End Select
35070       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
35080         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
35090       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
35100         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
35110       End Select
35120     End If  ' ** blnContinue.
35130     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
35140   End With

EXITP:
35150   Set frm1 = Nothing
35160   Set frm2 = Nothing
35170   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkTaxCodeDescription_AfterUpdate()

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkTaxCodeDescription_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

35210   With Me
35220     blnContinue = True
35230     Set frm1 = .frmTransaction_Audit_Sub.Form
35240     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
35250     If .ckgFlds_chkTaxCodeDescription = True Then
35260       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
35270         blnContinue = False
35280         Beep
35290         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
35300         .ckgFlds_chkTaxCodeDescription = False
35310       End If
35320     End If
35330     If blnContinue = True Then
35340       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
35350       Select Case .Controls(strCtlName)
            Case True
35360         .Controls(strCtlName & "_lbl").FontBold = True
35370         frmCrit.cmbTaxCodes.Enabled = True
35380         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
35390         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
35400       Case False
35410         .Controls(strCtlName & "_lbl").FontBold = False
35420         frmCrit.FocusHolder2.SetFocus
35430         frmCrit.cmbTaxCodes.Enabled = False
35440         frmCrit.cmbTaxCodes = Null
35450         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
35460         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
35470       End Select
35480       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
35490         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
35500       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
35510         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
35520       End Select
35530     End If  ' ** blnContinue.
35540     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
35550   End With

EXITP:
35560   Set frm1 = Nothing
35570   Set frm2 = Nothing
35580   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkTaxCodeTypeDescription_AfterUpdate()

35600 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkTaxCodeTypeDescription_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

35610   With Me
35620     blnContinue = True
35630     Set frm1 = .frmTransaction_Audit_Sub.Form
35640     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
35650     If .ckgFlds_chkTaxCodeTypeDescription = True Then
35660       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
35670         blnContinue = False
35680         Beep
35690         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
35700         .ckgFlds_chkTaxCodeTypeDescription = False
35710       End If
35720     End If
35730     If blnContinue = True Then
35740       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
35750       Select Case .Controls(strCtlName)
            Case True
35760         .Controls(strCtlName & "_lbl").FontBold = True
35770         frmCrit.chkTaxcodeType_Income.Enabled = True
35780         frmCrit.chkTaxcodeType_Deduction.Enabled = True
35790         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
35800         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
35810       Case False
35820         .Controls(strCtlName & "_lbl").FontBold = False
35830         frmCrit.FocusHolder2.SetFocus
35840         If frmCrit.chkTaxcodeType_Income = False Then
35850           frmCrit.chkTaxcodeType_Income = True
                'chkTaxcodeType_Income_AfterUpdate 3  NOT HIT ON OPEN!
35860           frmCrit.chkTaxcodeType_Income_AfterUpdate  ' ** Procedure: Above.
35870         End If
35880         frmCrit.chkTaxcodeType_Income.Enabled = False
35890         If frmCrit.chkTaxcodeType_Deduction = False Then
35900           frmCrit.chkTaxcodeType_Deduction = True
                'chkTaxcodeType_Deduction_AfterUpdate 3  NOT HIT ON OPEN!
35910           frmCrit.chkTaxcodeType_Deduction_AfterUpdate  ' ** Procedure: Above.
35920         End If
35930         frmCrit.chkTaxcodeType_Deduction.Enabled = False
35940         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
35950         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
35960       End Select
35970       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
35980         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
35990       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
36000         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
36010       End Select
36020     End If  ' ** blnContinue.
36030     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
36040   End With

EXITP:
36050   Set frm1 = Nothing
36060   Set frm2 = Nothing
36070   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkLocationName_AfterUpdate()

36100 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkLocationName_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

36110   With Me
36120     blnContinue = True
36130     Set frm1 = .frmTransaction_Audit_Sub.Form
36140     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
36150     If .ckgFlds_chkLocationName = True Then
36160       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
36170         blnContinue = False
36180         Beep
36190         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
36200         .ckgFlds_chkLocationName = False
36210       End If
36220     End If
36230     If blnContinue = True Then
36240       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
36250       Select Case .Controls(strCtlName)
            Case True
36260         .Controls(strCtlName & "_lbl").FontBold = True
36270         frmCrit.cmbLocations.Enabled = True
36280         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
36290         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
36300       Case False
36310         .Controls(strCtlName & "_lbl").FontBold = False
36320         frmCrit.FocusHolder2.SetFocus
36330         frmCrit.cmbLocations.Enabled = False
36340         frmCrit.cmbLocations = Null
36350         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
36360         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
36370       End Select
36380       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
36390         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
36400       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
36410         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
36420       End Select
36430     End If  ' ** blnContinue.
36440     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
36450   End With

EXITP:
36460   Set frm1 = Nothing
36470   Set frm2 = Nothing
36480   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkCheckNum_AfterUpdate()

36500 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkCheckNum_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

36510   With Me
36520     blnContinue = True
36530     Set frm1 = .frmTransaction_Audit_Sub.Form
36540     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
36550     If .ckgFlds_chkCheckNum = True Then
36560       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
36570         blnContinue = False
36580         Beep
36590         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
36600         .ckgFlds_chkCheckNum = False
36610       End If
36620     End If
36630     If blnContinue = True Then
36640       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
36650       Select Case .Controls(strCtlName)
            Case True
36660         .Controls(strCtlName & "_lbl").FontBold = True
36670         frmCrit.CheckNum.Enabled = True
36680         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
36690         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
36700       Case False
36710         .Controls(strCtlName & "_lbl").FontBold = False
36720         frmCrit.FocusHolder2.SetFocus
36730         frmCrit.CheckNum.Enabled = False
36740         frmCrit.CheckNum = Null
36750         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
36760         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
36770       End Select
36780       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
36790         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
36800       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
36810         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
36820       End Select
36830     End If  ' ** blnContinue.
36840     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
36850   End With

EXITP:
36860   Set frm1 = Nothing
36870   Set frm2 = Nothing
36880   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkJournalUser_AfterUpdate()

36900 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkJournalUser_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

36910   With Me
36920     blnContinue = True
36930     Set frm1 = .frmTransaction_Audit_Sub.Form
36940     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
36950     If .ckgFlds_chkJournalUser = True Then
36960       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
36970         blnContinue = False
36980         Beep
36990         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
37000         .ckgFlds_chkJournalUser = False
37010       End If
37020     End If
37030     If blnContinue = True Then
37040       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
37050       Select Case .Controls(strCtlName)
            Case True
37060         .Controls(strCtlName & "_lbl").FontBold = True
37070         frmCrit.cmbUsers.Enabled = True
37080         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
37090         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
37100       Case False
37110         .Controls(strCtlName & "_lbl").FontBold = False
37120         frmCrit.FocusHolder2.SetFocus
37130         frmCrit.cmbUsers.Enabled = False
37140         frmCrit.cmbUsers = Null
37150         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
37160         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
37170       End Select
37180       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
37190         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
37200       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
37210         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
37220       End Select
37230     End If  ' ** blnContinue.
37240     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
37250   End With

EXITP:
37260   Set frm1 = Nothing
37270   Set frm2 = Nothing
37280   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkPosted_AfterUpdate()

37300 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkPosted_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

37310   With Me
37320     blnContinue = True
37330     Set frm1 = .frmTransaction_Audit_Sub.Form
37340     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
37350     If .ckgFlds_chkPosted = True Then
37360       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
37370         blnContinue = False
37380         Beep
37390         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
37400         .ckgFlds_chkPosted = False
37410       End If
37420     End If
37430     If blnContinue = True Then
37440       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
37450       Select Case .Controls(strCtlName)
            Case True
37460         .Controls(strCtlName & "_lbl").FontBold = True
37470         frmCrit.PostedDateStart.Enabled = True
37480         frmCrit.PostedDateEnd.Enabled = True
37490         frmCrit.cmdCalendar7.Enabled = True
37500         frmCrit.Calendar_Handler frmCrit.cmdCalendar7.Name & "_Disable", True  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
37510         frmCrit.cmdCalendar8.Enabled = True
37520         frmCrit.Calendar_Handler frmCrit.cmdCalendar8.Name & "_Disable", True  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
37530         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
37540         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
37550       Case False
37560         .Controls(strCtlName & "_lbl").FontBold = False
37570         frmCrit.FocusHolder2.SetFocus
37580         frmCrit.PostedDateStart.Enabled = False
37590         frmCrit.PostedDateEnd.Enabled = False
37600         frmCrit.PostedDateStart = Null
37610         frmCrit.PostedDateEnd = Null
37620         frmCrit.cmdCalendar7.Enabled = False
37630         frmCrit.Calendar_Handler frmCrit.cmdCalendar7.Name & "_Disable", False  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
37640         frmCrit.cmdCalendar8.Enabled = False
37650         frmCrit.Calendar_Handler frmCrit.cmdCalendar8.Name & "_Disable", False  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
37660         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
37670         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
37680       End Select
37690       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
37700         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
37710       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
37720         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
37730       End Select
37740     End If  ' ** blnContinue.
37750     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
37760   End With

EXITP:
37770   Set frm1 = Nothing
37780   Set frm2 = Nothing
37790   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkLedgerHidden_AfterUpdate()

37800 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkLedgerHidden_AfterUpdate"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim strCtlName As String
        Dim blnContinue As Boolean

37810   With Me
37820     blnContinue = True
37830     Set frm1 = .frmTransaction_Audit_Sub.Form
37840     Set frm2 = .frmTransaction_Audit_Sub_ds.Form
37850     If .ckgFlds_chkLedgerHidden = True Then
37860       If frm1.CalcSpace(THIS_PROC) = False Then  ' ** Form Function: frmTransaction_Audit_Sub.
37870         blnContinue = False
37880         Beep
37890         MsgBox "There is not enough room to show this field." & vbCrLf & _
                "You must uncheck another field before you can show this one.", _
                vbInformation + vbOKOnly, "Too Many Fields"
37900         .ckgFlds_chkLedgerHidden = False
37910       End If
37920     End If
37930     If blnContinue = True Then
37940       strCtlName = Left(THIS_PROC, (CharPos(THIS_PROC, 2, "_") - 1))
37950       Select Case .Controls(strCtlName)
            Case True
37960         .Controls(strCtlName & "_lbl").FontBold = True
37970         frmCrit.opgHidden.Enabled = True
37980         ShowFields THIS_PROC, True, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
37990         frm1.ShowFields_Sub THIS_PROC, 2  ' ** Form Procedure: frmTransaction_Audit_Sub.
38000       Case False
38010         .Controls(strCtlName & "_lbl").FontBold = False
38020         frmCrit.FocusHolder2.SetFocus
38030         frmCrit.opgHidden.Enabled = False
38040         ShowFields THIS_PROC, False, Me  ' ** Module Procedure: modTransactionAuditFuncs1.
38050         frm1.ShowFields_Sub THIS_PROC, 3  ' ** Form Procedure: frmTransaction_Audit_Sub.
38060       End Select
38070       Select Case .opgView
            Case .opgView_optForm.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
38080         fSetScrollBarPosHZ frm1, 999&  ' ** Module Function: modScrollBarFuncs.
38090       Case .opgView_optDatasheet.OptionValue
              ' ** Move the subform scrollbar to the right where the field is.
38100         fSetScrollBarPosHZ frm2, 999&  ' ** Module Function: modScrollBarFuncs.
38110       End Select
38120     End If  ' ** blnContinue.
38130     ChecksChecked Me  ' ** Module Function: modTransactionAuditFuncs1.
38140   End With

EXITP:
38150   Set frm1 = Nothing
38160   Set frm2 = Nothing
38170   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Private Sub ckgFlds_chkLedgerHidden_KeyDown(KeyCode As Integer, Shift As Integer)

38200 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgFlds_chkLedgerHidden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

38210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
38220   intShiftDown = (Shift And acShiftMask) > 0
38230   intAltDown = (Shift And acAltMask) > 0
38240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
38250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
38260     Select Case intRetVal
          Case vbKeyTab
38270       With Me
38280         intRetVal = 0
38290         If .cmdPreview.Enabled = True Then
38300           .cmdPreview.SetFocus
38310         Else
38320           .cmdClose.SetFocus
38330         End If
38340       End With
38350     End Select
38360   End If

        ' ** Shift keys.
38370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
38380     Select Case intRetVal
          Case vbKeyTab
38390       With Me
38400         intRetVal = 0
38410         .ckgFlds_chkPosted.SetFocus
38420       End With
38430     End Select
38440   End If

EXITP:
38450   KeyCode = intRetVal
38460   Exit Sub

ERRH:
3910    intRetVal = 0
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Sub

Public Sub FormFields_Sort(lngXCalled As Long)
' ** Sorted by TabIndex.
' ** lngXCalled is lngX from a For/Next loop.

38500 On Error GoTo ERRH

        Const THIS_PROC As String = "FormFields_Sort"

        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long

        ' ****************************************************
        ' ** Array: arr_varFrmFld(), arr_varFrmFld_ds()
        ' **
        ' **   Element  Name                    Constant
        ' **   =======  ======================  ============
        ' **      0     Field Name              FM_FLD_NAM
        ' **      1     Control Tab Index       FM_FLD_TAB
        ' **      2     Field Visible           FM_FLD_VIS
        ' **      3     Checkbox Name           FM_CHK_NAM
        ' **      4     Checkbox Value          FM_CHK_VAL
        ' **
        ' ****************************************************

38510   Select Case lngXCalled
        Case 1&
          ' ** Binary Sort arr_varFrmFld() array by tab order (which is the left-to-right order on the form).
38520     For lngX = UBound(arr_varFrmFld, 2) To 1 Step -1
38530       For lngY = 0 To (lngX - 1)
38540         If arr_varFrmFld(FM_FLD_TAB, lngY) > arr_varFrmFld(FM_FLD_TAB, (lngY + 1)) Then
38550           For lngZ = 0& To FM_ELEMS
38560             varTmp00 = arr_varFrmFld(lngZ, lngY)
38570             arr_varFrmFld(lngZ, lngY) = arr_varFrmFld(lngZ, (lngY + 1))
38580             arr_varFrmFld(lngZ, (lngY + 1)) = varTmp00
38590             varTmp00 = Empty
38600           Next  ' ** lngZ.
38610         End If
38620       Next  ' ** lngY.
38630     Next  ' ** lngX.
38640   Case 2&
38650     For lngX = UBound(arr_varFrmFld_ds, 2) To 1 Step -1
38660       For lngY = 0 To (lngX - 1)
38670         If arr_varFrmFld_ds(FM_FLD_TAB, lngY) > arr_varFrmFld_ds(FM_FLD_TAB, (lngY + 1)) Then
38680           For lngZ = 0& To FM_ELEMS
38690             varTmp00 = arr_varFrmFld_ds(lngZ, lngY)
38700             arr_varFrmFld_ds(lngZ, lngY) = arr_varFrmFld_ds(lngZ, (lngY + 1))
38710             arr_varFrmFld_ds(lngZ, (lngY + 1)) = varTmp00
38720             varTmp00 = Empty
38730           Next  ' ** lngZ.
38740         End If
38750       Next  ' ** lngY.
38760     Next  ' ** lngX.
38770   End Select

EXITP:
38780   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Public Function FormFields_GetArr(intMode As Integer) As Variant
' ** intMode is opgView.

38800 On Error GoTo ERRH

        Const THIS_PROC As String = "FormFields_GetArr"

        Dim blnContinue As Boolean
        Dim varTmp00 As Variant, arr_varTmp01() As Variant
        Dim varRetVal As Variant

38810   blnContinue = True

38820   Select Case intMode
        Case 1
38830 On Error Resume Next
38840     varTmp00 = UBound(arr_varFrmFld, 2)  ' ** I don't really know why this is causing so much trouble!
38850     If ERR.Number <> 0 Then
38860       blnContinue = False
38870       If IsEmpty(arr_varFrmFld) = True Then
              'MsgBox "FRM: arr_varFrmFld IS EMPTY!"
38880       Else
38890         If IsNull(arr_varFrmFld) = True Then
                'MsgBox "FRM: arr_varFrmFld IS NULL!"
38900         Else
38910           If ERR.Number = 9 Then  ' ** Subscript out of range.
                  'MsgBox "FRM arr_varFrmFld(0): " & Nz(arr_varFrmFld(0), "Null")
38920           Else
                  'MsgBox "FRM ERROR: " & CStr(ERR.Number) & "  " & ERR.description
38930           End If
38940         End If
38950       End If
38960 On Error GoTo ERRH
38970     Else
38980 On Error GoTo ERRH
38990     End If
39000     If blnContinue = True Then
39010       varRetVal = arr_varFrmFld
39020     Else
39030       ReDim arr_varTmp01(0, 0)
39040       varRetVal = arr_varTmp01
39050     End If
39060   Case 2
39070     varTmp00 = UBound(arr_varFrmFld_ds, 2)
39080     If ERR.Number <> 0 Then
39090       blnContinue = False
39100       If IsEmpty(arr_varFrmFld_ds) = True Then
              'MsgBox "FRM: arr_varFrmFld IS EMPTY!"
39110       Else
39120         If IsNull(arr_varFrmFld_ds) = True Then
                'MsgBox "FRM: arr_varFrmFld IS NULL!"
39130         Else
39140           If ERR.Number = 9 Then  ' ** Subscript out of range.
                  'MsgBox "FRM arr_varFrmFld(0): " & Nz(arr_varFrmFld(0), "Null")
39150           Else
                  'MsgBox "FRM ERROR: " & CStr(ERR.Number) & "  " & ERR.description
39160           End If
39170         End If
39180       End If
39190 On Error GoTo ERRH
39200     Else
39210 On Error GoTo ERRH
39220     End If
39230     If blnContinue = True Then
39240       varRetVal = arr_varFrmFld_ds
39250     Else
39260       ReDim arr_varTmp01(0, 0)
39270       varRetVal = arr_varTmp01
39280     End If
39290   End Select

EXITP:
39300   FormFields_GetArr = varRetVal
39310   Exit Function

ERRH:
3910    varRetVal = Null
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Function

Public Sub FilterRecs_Set(strFltr As String, dblFltrRecs As Double)

39400 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Set"

39410   With Me
39420     strFilter01 = strFltr
39430     strFilter01 = .frmTransaction_Audit_Sub_ds.Form.SwapNames(strFilter01, True)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
39440     strFilter02 = strFltr
39450     strFilter02 = .frmTransaction_Audit_Sub_ds.Form.SwapNames(strFilter02)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
39460     dblFilterRecs = dblFltrRecs
39470   End With

EXITP:
39480   Exit Sub

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Sub

Public Sub FilterRecsZero()
' ** Flash the Records Returned label.

39500 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecsZero"

39510   With Me
39520     blnRecsZero = True
39530     intRecsZero = intRecsZero + 1
39540     Select Case .FilterRecs_lbl.Visible
          Case True
39550       .FilterRecs_lbl.Visible = False
39560       .FilterRecs.ForeColor = CLR_LTTEAL
39570       .TimerInterval = 200&
39580     Case False
39590       .FilterRecs_lbl.Visible = True
39600       .FilterRecs.ForeColor = CLR_DKRED
39610       .TimerInterval = 200&
39620     End Select
39630   End With

EXITP:
39640   Exit Sub

ERRH:
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Public Function FrmIsMax() As Boolean

39700 On Error GoTo ERRH

        Const THIS_PROC As String = "FrmIsMax"

        Dim blnRetVal As Boolean

39710   blnRetVal = IsMaximized(Me)  ' ** Module Function: modWindowFunctions.

EXITP:
39720   FrmIsMax = blnRetVal
39730   Exit Function

ERRH:
3910    THAT_PROC = THIS_PROC
3920    That_Erl = Erl
3930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3940    Resume EXITP

End Function

Public Function VisCnts_Get() As String

39800 On Error GoTo ERRH

        Const THIS_PROC As String = "VisCnts_Get"

        Dim strRetVal As String

39810   strRetVal = CStr(lngChkCnt) & "~" & CStr(lngVisCnt1) & "~" & CStr(lngVisCnt2)

EXITP:
39820   VisCnts_Get = strRetVal
39830   Exit Function

ERRH:
3910    strRetVal = "0~0~0"
3920    THAT_PROC = THIS_PROC
3930    That_Erl = Erl
3940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3950    Resume EXITP

End Function
