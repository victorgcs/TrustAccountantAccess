Attribute VB_Name = "modTransactionAuditFuncs1"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modTransactionAuditFuncs1"

'VGC 10/05/2017: CHANGES!

' ** Filter constants: frmTransaction_Audit_Sub.
Private Const ANDF         As String = " And "  ' ** Filter 'And'.
Private Const ORF          As String = " Or "  ' ** Filter 'Or'.
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
Private Const POSTED_START As String = "[posted] >= #"
Private Const POSTED_END   As String = "[posted] <= #"
Private Const HIDDEN_TRX1  As String = "[ledger_HIDDEN] = True"
Private Const HIDDEN_TRX2  As String = "[ledger_HIDDEN] = False"

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Array: arr_varCal().
Private lngCals As Long, arr_varCal() As Variant

' ** Array: arr_varFld().
Private lngFlds As Long, arr_varFld() As Variant
Private Const FLD_ELEMS As Integer = 13  ' ** Array's first-element UBound().
Private Const F_CNAM     As Integer = 0
Private Const F_FNAM     As Integer = 1
Private Const F_LFT      As Integer = 2
Private Const F_WDT      As Integer = 3
Private Const F_LBL1     As Integer = 4
Private Const F_LBL2     As Integer = 5
Private Const F_LBL_LFT  As Integer = 6
Private Const F_LBL_WDT  As Integer = 7
Private Const F_LIN      As Integer = 8
Private Const F_LIN_LFT  As Integer = 9
Private Const F_LIN_WDT  As Integer = 10
Private Const F_SRT_ADJ  As Integer = 11
Private Const F_VIS      As Integer = 12
Private Const F_CHK_ELEM As Integer = 13

 ' ** Array: arr_varFilt(), arr_varFilt_ds().
Private lngFilts As Long, arr_varFilt As Variant
Private lngFilts_ds As Long, arr_varFilt_ds As Variant
Private Const F_ELEMS As Integer = 13  ' ** Array's first-element UBound().
Private Const F_IDX   As Integer = 0
Private Const F_NAM   As Integer = 1
Private Const F_CONST As Integer = 2
Private Const F_CTL   As Integer = 3
Private Const F_CLBL  As Integer = 4
Private Const F_FLD   As Integer = 5
Private Const F_FLBL  As Integer = 6
Private Const F_CTL2  As Integer = 7
Private Const F_CLBL2 As Integer = 8
Private Const F_FLD2  As Integer = 9
Private Const F_FLBL2 As Integer = 10
Private Const F_CLBL3 As Integer = 11
Private Const F_FLD3  As Integer = 12
Private Const F_FLBL3 As Integer = 13

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

' ** Array: arr_varChk().
Private lngChks As Long, arr_varChk() As Variant
Private Const C_ELEMS As Integer = 17  ' ** Array's first-element UBound().
Private Const C_FNAM  As Integer = 0
Private Const C_CHKBX As Integer = 1
Private Const C_INCL  As Integer = 2
Private Const C_FOC   As Integer = 3
Private Const C_MOUS  As Integer = 4
Private Const C_DIS   As Integer = 5
Private Const C_CMD   As Integer = 6
Private Const C_OFR   As Integer = 7
Private Const C_OFRD  As Integer = 8
Private Const C_OFRF  As Integer = 9
Private Const C_OFRFD As Integer = 10
Private Const C_OFDIS As Integer = 11
Private Const C_ONR   As Integer = 12
Private Const C_ONRD  As Integer = 13
Private Const C_ONRF  As Integer = 14
Private Const C_ONRFD As Integer = 15
Private Const C_ONSD  As Integer = 16
Private Const C_ONDIS As Integer = 17

Private frmCrit As Access.Form

Private blnFromSet As Boolean
Private lngChkCnt As Long, lngVisCnt1 As Long, lngVisCnt2 As Long
Private lngTpp As Long, lngMonitorCnt As Long, lngMonitorNum As Long
' **

Public Sub FilterRecs_Load()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngX As Long

110     For lngX = 1& To 2&
120       Select Case lngX
          Case 1&

            ' ** We'll be automating this better when it's complete.  OH YAH?
            'lngFilts = 21&  ' ** 0 - 20.
            'lngE = lngFilts - 1&
            'ReDim arr_varFilt(F_ELEMS, lngE)

130         Set dbs = CurrentDb
140         With dbs
              ' ** tblTransaction_Audit_Filter, just 'frmTransaction_Audit_Sub'.
150           Set qdf = .QueryDefs("qryTransaction_Audit_08_01")
160           Set rst = qdf.OpenRecordset
170           With rst
180             .MoveLast
190             lngFilts = .RecordCount
200             .MoveFirst
210             arr_varFilt = .GetRows(lngFilts)
                ' ******************************************************
                ' ** Array: arr_varFilt()
                ' **
                ' **   Field  Element  Name                 Constant
                ' **   =====  =======  ===================  ==========
                ' **     1       0     taf_index            F_NAM
                ' **     2       1     vbdec_name1          F_NAM
                ' **     3       2     vbdec_value1         F_CONST
                ' **     4       3     ctl_name1            F_CTL
                ' **     5       4     ctl_name_lbl1        F_CLBL
                ' **     6       5     fld_ctl_name1        F_FLD
                ' **     7       6     fld_ctl_name_lbl1    F_FLBL
                ' **     8       7     ctl_name2            F_CTL2
                ' **     9       8     ctl_name_lbl2        F_CLBL2
                ' **    10       9     fld_ctl_name2        F_FLD2
                ' **    11      10     fld_ctl_name_lbl2    F_FLBL2
                ' **    12      11     ctl_name_lbl3        F_CLBL3
                ' **    13      12     fld_ctl_name3        F_FLD3
                ' **    14      13     fld_ctl_name_lbl3    F_FLBL3
                ' **
                ' ******************************************************
220             .Close
230           End With  ' ** rst.
240           Set rst = Nothing
250           Set qdf = Nothing
260           .Close
270         End With  ' ** dbs.
280         Set dbs = Nothing
290         DoEvents

            ' ** The filter array has only the fields that can be filtered.
            'arr_varFilt(F_NAM, 0) = "JRNL_NUM"
            'arr_varFilt(F_CONST, 0) = JRNL_NUM
            'arr_varFilt(F_CTL, 0) = "journalno"
            'arr_varFilt(F_CLBL, 0) = "journalno_lbl"
            'arr_varFilt(F_FLD, 0) = "journalno"
            'arr_varFilt(F_FLBL, 0) = "journalno_lbl" & "~" & "journalno_lbl2"
            'arr_varFilt(F_CTL2, 0) = vbNullString
            'arr_varFilt(F_CLBL2, 0) = vbNullString
            'arr_varFilt(F_FLD2, 0) = vbNullString
            'arr_varFilt(F_FLBL2, 0) = vbNullString
            'arr_varFilt(F_CLBL3, 0) = vbNullString
            'arr_varFilt(F_FLD3, 0) = vbNullString
            'arr_varFilt(F_FLBL3, 0) = vbNullString

            'arr_varFilt(F_NAM, 1) = "JRNL_TYPE"
            'arr_varFilt(F_CONST, 1) = JRNL_TYPE
            'arr_varFilt(F_CTL, 1) = "cmbJournalType1"
            'arr_varFilt(F_CLBL, 1) = "cmbJournalType1_lbl"
            'arr_varFilt(F_FLD, 1) = "journaltype"
            'arr_varFilt(F_FLBL, 1) = "journaltype_lbl" & "~" & "journaltype_lbl2"
            'arr_varFilt(F_CTL2, 1) = vbNullString
            'arr_varFilt(F_CLBL2, 1) = "cmbJournalType_lbl"
            'arr_varFilt(F_FLD2, 1) = vbNullString
            'arr_varFilt(F_FLBL2, 1) = vbNullString
            'arr_varFilt(F_CLBL3, 1) = vbNullString
            'arr_varFilt(F_FLD3, 1) = vbNullString
            'arr_varFilt(F_FLBL3, 1) = vbNullString

            'arr_varFilt(F_NAM, 2) = "TRANS_START"
            'arr_varFilt(F_CONST, 2) = TRANS_START
            'arr_varFilt(F_CTL, 2) = "TransDateStart"
            'arr_varFilt(F_CLBL, 2) = "TransDateStart_lbl"
            'arr_varFilt(F_FLD, 2) = "transdate"
            'arr_varFilt(F_FLBL, 2) = "transdate_lbl" & "~" & "transdate_lbl2"
            'arr_varFilt(F_CTL2, 2) = "TransDateEnd"
            'arr_varFilt(F_CLBL2, 2) = "TransDateEnd_lbl"
            'arr_varFilt(F_FLD2, 2) = vbNullString
            'arr_varFilt(F_FLBL2, 2) = "transdate_lbl2"
            'arr_varFilt(F_CLBL3, 2) = "TransDateStart_lbl2"
            'arr_varFilt(F_FLD3, 2) = vbNullString
            'arr_varFilt(F_FLBL3, 2) = vbNullString

            'arr_varFilt(F_NAM, 3) = "TRANS_END"
            'arr_varFilt(F_CONST, 3) = TRANS_END
            'arr_varFilt(F_CTL, 3) = "TransDateEnd"
            'arr_varFilt(F_CLBL, 3) = "TransDateEnd_lbl"
            'arr_varFilt(F_FLD, 3) = "transdate"
            'arr_varFilt(F_FLBL, 3) = "transdate_lbl" & "~" & "transdate_lbl2"
            'arr_varFilt(F_CTL2, 3) = "TransDateStart"
            'arr_varFilt(F_CLBL2, 3) = "TransDateStart_lbl"
            'arr_varFilt(F_FLD2, 3) = vbNullString
            'arr_varFilt(F_FLBL2, 3) = "transdate_lbl2"
            'arr_varFilt(F_CLBL3, 3) = "TransDateStart_lbl2"
            'arr_varFilt(F_FLD3, 3) = vbNullString
            'arr_varFilt(F_FLBL3, 3) = vbNullString

            'arr_varFilt(F_NAM, 4) = "ACCT_NUM"
            'arr_varFilt(F_CONST, 4) = ACCT_NUM
            'arr_varFilt(F_CTL, 4) = "cmbAccounts"
            'arr_varFilt(F_CLBL, 4) = "cmbAccounts_lbl"
            'arr_varFilt(F_FLD, 4) = "accountno"
            'arr_varFilt(F_FLBL, 4) = "accountno_lbl" & "~" & "accountno_lbl2"
            'arr_varFilt(F_CTL2, 4) = vbNullString
            'arr_varFilt(F_CLBL2, 4) = vbNullString
            'arr_varFilt(F_FLD2, 4) = "shortname"
            'arr_varFilt(F_FLBL2, 4) = "shortname_lbl"
            'arr_varFilt(F_CLBL3, 4) = vbNullString
            'arr_varFilt(F_FLD3, 4) = vbNullString
            'arr_varFilt(F_FLBL3, 4) = vbNullString

            'arr_varFilt(F_NAM, 5) = "ASSET_NUM"
            'arr_varFilt(F_CONST, 5) = ASSET_NUM
            'arr_varFilt(F_CTL, 5) = "cmbAssets"
            'arr_varFilt(F_CLBL, 5) = "cmbAssets_lbl"
            'arr_varFilt(F_FLD, 5) = "assetno"
            'arr_varFilt(F_FLBL, 5) = vbNullString  ' ** Field not displayed.
            'arr_varFilt(F_CTL2, 5) = vbNullString
            'arr_varFilt(F_CLBL2, 5) = vbNullString
            'arr_varFilt(F_FLD2, 5) = "asset_description"
            'arr_varFilt(F_FLBL2, 5) = "asset_description_lbl"
            'arr_varFilt(F_CLBL3, 5) = vbNullString
            'arr_varFilt(F_FLD3, 5) = "cusip"
            'arr_varFilt(F_FLBL3, 5) = "cusip_lbl"

            'arr_varFilt(F_NAM, 6) = "CURR_NUM"
            'arr_varFilt(F_CONST, 6) = CURR_NUM
            'arr_varFilt(F_CTL, 6) = "cmbCurrencies"
            'arr_varFilt(F_CLBL, 6) = "cmbCurrencies_lbl"
            'arr_varFilt(F_FLD, 6) = "curr_id"
            'arr_varFilt(F_FLBL, 6) = "curr_id_lbl"
            'arr_varFilt(F_CTL2, 6) = vbNullString
            'arr_varFilt(F_CLBL2, 6) = vbNullString
            'arr_varFilt(F_FLD2, 6) = vbNullString
            'arr_varFilt(F_FLBL2, 6) = vbNullString
            'arr_varFilt(F_CLBL3, 6) = vbNullString
            'arr_varFilt(F_FLD3, 6) = vbNullString
            'arr_varFilt(F_FLBL3, 6) = vbNullString

            'arr_varFilt(F_NAM, 7) = "ASSET_START"
            'arr_varFilt(F_CONST, 7) = ASSET_START
            'arr_varFilt(F_CTL, 7) = "AssetDateStart"
            'arr_varFilt(F_CLBL, 7) = "AssetDateStart_lbl"
            'arr_varFilt(F_FLD, 7) = "assetdate"
            'arr_varFilt(F_FLBL, 7) = "assetdate_lbl"
            'arr_varFilt(F_CTL2, 7) = "AssetDateEnd"
            'arr_varFilt(F_CLBL2, 7) = "AssetDateEnd_lbl"
            'arr_varFilt(F_FLD2, 7) = vbNullString
            'arr_varFilt(F_FLBL2, 7) = vbNullString
            'arr_varFilt(F_CLBL3, 7) = "AssetDateStart_lbl2"
            'arr_varFilt(F_FLD3, 7) = vbNullString
            'arr_varFilt(F_FLBL3, 7) = vbNullString

            'arr_varFilt(F_NAM, 8) = "ASSET_END"
            'arr_varFilt(F_CONST, 8) = ASSET_END
            'arr_varFilt(F_CTL, 8) = "AssetDateEnd"
            'arr_varFilt(F_CLBL, 8) = "AssetDateEnd_lbl"
            'arr_varFilt(F_FLD, 8) = "assetdate"
            'arr_varFilt(F_FLBL, 8) = "assetdate_lbl"
            'arr_varFilt(F_CTL2, 8) = "AssetDateStart"
            'arr_varFilt(F_CLBL2, 8) = "AssetDateStart_lbl"
            'arr_varFilt(F_FLD2, 8) = vbNullString
            'arr_varFilt(F_FLBL2, 8) = vbNullString
            'arr_varFilt(F_CLBL3, 8) = "AssetDateStart_lbl2"
            'arr_varFilt(F_FLD3, 8) = vbNullString
            'arr_varFilt(F_FLBL3, 8) = vbNullString

            'arr_varFilt(F_NAM, 9) = "PURCH_START"
            'arr_varFilt(F_CONST, 9) = PURCH_START
            'arr_varFilt(F_CTL, 9) = "PurchaseDateStart"
            'arr_varFilt(F_CLBL, 9) = "PurchaseDateStart_lbl"
            'arr_varFilt(F_FLD, 9) = "PurchaseDate"
            'arr_varFilt(F_FLBL, 9) = "PurchaseDate_lbl"
            'arr_varFilt(F_CTL2, 9) = "PurchaseDateEnd"
            'arr_varFilt(F_CLBL2, 9) = "PurchaseDateEnd_lbl"
            'arr_varFilt(F_FLD2, 9) = vbNullString
            'arr_varFilt(F_FLBL2, 9) = vbNullString
            'arr_varFilt(F_CLBL3, 9) = "PurchaseDateStart_lbl2"
            'arr_varFilt(F_FLD3, 9) = vbNullString
            'arr_varFilt(F_FLBL3, 9) = vbNullString

            'arr_varFilt(F_NAM, 10) = "PURCH_END"
            'arr_varFilt(F_CONST, 10) = PURCH_END
            'arr_varFilt(F_CTL, 10) = "PurchaseDateEnd"
            'arr_varFilt(F_CLBL, 10) = "PurchaseDateEnd_lbl"
            'arr_varFilt(F_FLD, 10) = "PurchaseDate"
            'arr_varFilt(F_FLBL, 10) = "PurchaseDate_lbl"
            'arr_varFilt(F_CTL2, 10) = "PurchaseDateStart"
            'arr_varFilt(F_CLBL2, 10) = "PurchaseDateStart_lbl"
            'arr_varFilt(F_FLD2, 10) = vbNullString
            'arr_varFilt(F_FLBL2, 10) = vbNullString
            'arr_varFilt(F_CLBL3, 10) = "PurchaseDateStart_lbl2"
            'arr_varFilt(F_FLD3, 10) = vbNullString
            'arr_varFilt(F_FLBL3, 10) = vbNullString

            'arr_varFilt(F_NAM, 11) = "COMM_DESC"
            'arr_varFilt(F_CONST, 11) = COMM_DESC
            'arr_varFilt(F_CTL, 11) = "ledger_description"
            'arr_varFilt(F_CLBL, 11) = "ledger_description_lbl"
            'arr_varFilt(F_FLD, 11) = "ledger_description"
            'arr_varFilt(F_FLBL, 11) = "ledger_description_lbl"
            'arr_varFilt(F_CTL2, 11) = vbNullString
            'arr_varFilt(F_CLBL2, 11) = vbNullString
            'arr_varFilt(F_FLD2, 11) = vbNullString
            'arr_varFilt(F_FLBL2, 11) = vbNullString
            'arr_varFilt(F_CLBL3, 11) = vbNullString
            'arr_varFilt(F_FLD3, 11) = vbNullString
            'arr_varFilt(F_FLBL3, 11) = vbNullString

            'arr_varFilt(F_NAM, 12) = "RECUR_ITEM"
            'arr_varFilt(F_CONST, 12) = RECUR_ITEM
            'arr_varFilt(F_CTL, 12) = "cmbRecurringItems"
            'arr_varFilt(F_CLBL, 12) = "cmbRecurringItems_lbl"
            'arr_varFilt(F_FLD, 12) = "RecurringItem"
            'arr_varFilt(F_FLBL, 12) = "RecurringItem_lbl"
            'arr_varFilt(F_CTL2, 12) = vbNullString
            'arr_varFilt(F_CLBL2, 12) = vbNullString
            'arr_varFilt(F_FLD2, 12) = vbNullString
            'arr_varFilt(F_FLBL2, 12) = vbNullString
            'arr_varFilt(F_CLBL3, 12) = vbNullString
            'arr_varFilt(F_FLD3, 12) = vbNullString
            'arr_varFilt(F_FLBL3, 12) = vbNullString

            'arr_varFilt(F_NAM, 13) = "REV_CODE"
            'arr_varFilt(F_CONST, 13) = REV_CODE
            'arr_varFilt(F_CTL, 13) = "cmbRevenueCodes"
            'arr_varFilt(F_CLBL, 13) = "cmbRevenueCodes_lbl"
            'arr_varFilt(F_FLD, 13) = "revcode_ID"
            'arr_varFilt(F_FLBL, 13) = vbNullString  ' ** Field not displayed.
            'arr_varFilt(F_CTL2, 13) = vbNullString
            'arr_varFilt(F_CLBL2, 13) = vbNullString
            'arr_varFilt(F_FLD2, 13) = "revcode_DESC"
            'arr_varFilt(F_FLBL2, 13) = "revcode_DESC_lbl"
            'arr_varFilt(F_CLBL3, 13) = vbNullString
            'arr_varFilt(F_FLD3, 13) = vbNullString
            'arr_varFilt(F_FLBL3, 13) = vbNullString

            'arr_varFilt(F_NAM, 14) = "TAX_CODE"
            'arr_varFilt(F_CONST, 14) = TAX_CODE
            'arr_varFilt(F_CTL, 14) = "cmbTaxCodes"
            'arr_varFilt(F_CLBL, 14) = "cmbTaxCodes_lbl"
            'arr_varFilt(F_FLD, 14) = "taxcode"
            'arr_varFilt(F_FLBL, 14) = vbNullString  ' ** Field not displayed.
            'arr_varFilt(F_CTL2, 14) = vbNullString
            'arr_varFilt(F_CLBL2, 14) = vbNullString
            'arr_varFilt(F_FLD2, 14) = "taxcode_description"
            'arr_varFilt(F_FLBL2, 14) = "taxcode_description_lbl"
            'arr_varFilt(F_CLBL3, 14) = vbNullString
            'arr_varFilt(F_FLD3, 14) = vbNullString
            'arr_varFilt(F_FLBL3, 14) = vbNullString

            'arr_varFilt(F_NAM, 15) = "LOC_NUM"
            'arr_varFilt(F_CONST, 15) = LOC_NUM
            'arr_varFilt(F_CTL, 15) = "cmbLocations"
            'arr_varFilt(F_CLBL, 15) = "cmbLocations_lbl"
            'arr_varFilt(F_FLD, 15) = "Location_ID"
            'arr_varFilt(F_FLBL, 15) = "Location_ID_lbl"
            'arr_varFilt(F_CTL2, 15) = vbNullString
            'arr_varFilt(F_CLBL2, 15) = vbNullString
            'arr_varFilt(F_FLD2, 15) = vbNullString
            'arr_varFilt(F_FLBL2, 15) = vbNullString
            'arr_varFilt(F_CLBL3, 15) = vbNullString
            'arr_varFilt(F_FLD3, 15) = vbNullString
            'arr_varFilt(F_FLBL3, 15) = vbNullString

            'arr_varFilt(F_NAM, 16) = "CHK_NUM"  'CHK_NUM1  'CHK_NUM2
            'arr_varFilt(F_CONST, 16) = CHK_NUM
            'arr_varFilt(F_CTL, 16) = "CheckNum"
            'arr_varFilt(F_CLBL, 16) = "CheckNum_lbl"
            'arr_varFilt(F_FLD, 16) = "CheckNum"
            'arr_varFilt(F_FLBL, 16) = "CheckNum_lbl" & "~" & "CheckNum_lbl2"
            'arr_varFilt(F_CTL2, 16) = vbNullString
            'arr_varFilt(F_CLBL2, 16) = vbNullString
            'arr_varFilt(F_FLD2, 16) = vbNullString
            'arr_varFilt(F_FLBL2, 16) = vbNullString
            'arr_varFilt(F_CLBL3, 16) = vbNullString
            'arr_varFilt(F_FLD3, 16) = vbNullString
            'arr_varFilt(F_FLBL3, 16) = vbNullString

            'arr_varFilt(F_NAM, 17) = "JRNL_USER"
            'arr_varFilt(F_CONST, 17) = JRNL_USER
            'arr_varFilt(F_CTL, 17) = "cmbUsers"
            'arr_varFilt(F_CLBL, 17) = "cmbUsers_lbl"
            'arr_varFilt(F_FLD, 17) = "journal_USER"
            'arr_varFilt(F_FLBL, 17) = "journal_USER_lbl"
            'arr_varFilt(F_CTL2, 17) = vbNullString
            'arr_varFilt(F_CLBL2, 17) = vbNullString
            'arr_varFilt(F_FLD2, 17) = vbNullString
            'arr_varFilt(F_FLBL2, 17) = vbNullString
            'arr_varFilt(F_CLBL3, 17) = vbNullString
            'arr_varFilt(F_FLD3, 17) = vbNullString
            'arr_varFilt(F_FLBL3, 17) = vbNullString

            'arr_varFilt(F_NAM, 18) = "POSTED_START"
            'arr_varFilt(F_CONST, 18) = POSTED_START
            'arr_varFilt(F_CTL, 18) = "PostedDateStart"
            'arr_varFilt(F_CLBL, 18) = "PostedDateStart_lbl"
            'arr_varFilt(F_FLD, 18) = "posted"
            'arr_varFilt(F_FLBL, 18) = "posted_lbl"
            'arr_varFilt(F_CTL2, 18) = "PostedDateEnd"
            'arr_varFilt(F_CLBL2, 18) = "PostedDateEnd_lbl"
            'arr_varFilt(F_FLD2, 18) = vbNullString
            'arr_varFilt(F_FLBL2, 18) = vbNullString
            'arr_varFilt(F_CLBL3, 18) = "PostedDateStart_lbl2"
            'arr_varFilt(F_FLD3, 18) = vbNullString
            'arr_varFilt(F_FLBL3, 18) = vbNullString

            'arr_varFilt(F_NAM, 19) = "POSTED_END"
            'arr_varFilt(F_CONST, 19) = POSTED_END
            'arr_varFilt(F_CTL, 19) = "PostedDateEnd"
            'arr_varFilt(F_CLBL, 19) = "PostedDateEnd_lbl"
            'arr_varFilt(F_FLD, 19) = "posted"
            'arr_varFilt(F_FLBL, 19) = "posted_lbl"
            'arr_varFilt(F_CTL2, 19) = "PostedDateStart"
            'arr_varFilt(F_CLBL2, 19) = "PostedDateStart_lbl"
            'arr_varFilt(F_FLD2, 19) = vbNullString
            'arr_varFilt(F_FLBL2, 19) = vbNullString
            'arr_varFilt(F_CLBL3, 19) = "PostedDateStart_lbl2"
            'arr_varFilt(F_FLD3, 19) = vbNullString
            'arr_varFilt(F_FLBL3, 19) = vbNullString

            'arr_varFilt(F_NAM, 20) = "HIDDEN_TRX1"  'HIDDEN_TRX2
            'arr_varFilt(F_CONST, 20) = HIDDEN_TRX1
            'arr_varFilt(F_CTL, 20) = "opgHidden"
            'arr_varFilt(F_CLBL, 20) = "opgHidden_lbl"
            'arr_varFilt(F_FLD, 20) = "ledger_HIDDEN"
            'arr_varFilt(F_FLBL, 20) = "ledger_HIDDEN_lbl"
            'arr_varFilt(F_CTL2, 20) = vbNullString
            'arr_varFilt(F_CLBL2, 20) = vbNullString
            'arr_varFilt(F_FLD2, 20) = vbNullString
            'arr_varFilt(F_FLBL2, 20) = vbNullString
            'arr_varFilt(F_CLBL3, 20) = vbNullString
            'arr_varFilt(F_FLD3, 20) = vbNullString
            'arr_varFilt(F_FLBL3, 20) = vbNullString

300       Case 2&

            ' ** We'll be automating this better when it's complete.
            'lngFilts_ds = 21&  ' ** 0 - 20.
            'lngE = lngFilts_ds - 1&
            'ReDim arr_varFilt_ds(F_ELEMS, lngE)

310         Set dbs = CurrentDb
320         With dbs
              ' ** tblTransaction_Audit_Filter, just 'frmTransaction_Audit_Sub_ds'.
330           Set qdf = .QueryDefs("qryTransaction_Audit_08_02")
340           Set rst = qdf.OpenRecordset
350           With rst
360             .MoveLast
370             lngFilts_ds = .RecordCount
380             .MoveFirst
390             arr_varFilt_ds = .GetRows(lngFilts_ds)
                ' ******************************************************
                ' ** Array: arr_varFilt_ds()
                ' **
                ' **   Field  Element  Name                 Constant
                ' **   =====  =======  ===================  ==========
                ' **     1       0     taf_index            F_NAM
                ' **     2       1     vbdec_name1          F_NAM
                ' **     3       2     vbdec_value1         F_CONST
                ' **     4       3     ctl_name1            F_CTL
                ' **     5       4     ctl_name_lbl1        F_CLBL
                ' **     6       5     fld_ctl_name1        F_FLD
                ' **     7       6     fld_ctl_name_lbl1    F_FLBL
                ' **     8       7     ctl_name2            F_CTL2
                ' **     9       8     ctl_name_lbl2        F_CLBL2
                ' **    10       9     fld_ctl_name2        F_FLD2
                ' **    11      10     fld_ctl_name_lbl2    F_FLBL2
                ' **    12      11     ctl_name_lbl3        F_CLBL3
                ' **    13      12     fld_ctl_name3        F_FLD3
                ' **    14      13     fld_ctl_name_lbl3    F_FLBL3
                ' **
                ' ******************************************************
400             .Close
410           End With  ' ** rst.
420           Set rst = Nothing
430           Set qdf = Nothing
440           .Close
450         End With  ' ** dbs.
460         Set dbs = Nothing
470         DoEvents

            ' ** The filter array has only the fields that can be filtered.
            'arr_varFilt_ds(F_NAM, 0) = "JRNL_NUM"
            'arr_varFilt_ds(F_CONST, 0) = JRNL_NUM
            'arr_varFilt_ds(F_CTL, 0) = "journalno"
            'arr_varFilt_ds(F_CLBL, 0) = "journalno_lbl"
            'arr_varFilt_ds(F_FLD, 0) = "Journal Number"
            'arr_varFilt_ds(F_FLBL, 0) = "journalno_lbl" & "~" & "journalno_lbl2"  '"Journal_Number_lbl" & "~" & "Journal_Number_lbl2"
            'arr_varFilt_ds(F_CTL2, 0) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 0) = vbNullString
            'arr_varFilt_ds(F_FLD2, 0) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 0) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 0) = vbNullString
            'arr_varFilt_ds(F_FLD3, 0) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 0) = vbNullString

            'arr_varFilt_ds(F_NAM, 1) = "JRNL_TYPE"
            'arr_varFilt_ds(F_CONST, 1) = JRNL_TYPE
            'arr_varFilt_ds(F_CTL, 1) = "cmbJournalType1"
            'arr_varFilt_ds(F_CLBL, 1) = "cmbJournalType1_lbl"
            'arr_varFilt_ds(F_FLD, 1) = "Journal Type"
            'arr_varFilt_ds(F_FLBL, 1) = "journaltype_lbl" & "~" & "journaltype_lbl2"  '"Journal_Type_lbl" & "~" & "Journal_Type_lbl2"
            'arr_varFilt_ds(F_CTL2, 1) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 1) = "cmbJournalType_lbl"
            'arr_varFilt_ds(F_FLD2, 1) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 1) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 1) = vbNullString
            'arr_varFilt_ds(F_FLD3, 1) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 1) = vbNullString

            'arr_varFilt_ds(F_NAM, 2) = "TRANS_START"
            'arr_varFilt_ds(F_CONST, 2) = TRANS_START
            'arr_varFilt_ds(F_CTL, 2) = "TransDateStart"
            'arr_varFilt_ds(F_CLBL, 2) = "TransDateStart_lbl"
            'arr_varFilt_ds(F_FLD, 2) = "Posting Date"
            'arr_varFilt_ds(F_FLBL, 2) = "transdate_lbl" & "~" & "transdate_lbl2"  '"Posting_Date_lbl"
            'arr_varFilt_ds(F_CTL2, 2) = "TransDateEnd"
            'arr_varFilt_ds(F_CLBL2, 2) = "TransDateEnd_lbl"
            'arr_varFilt_ds(F_FLD2, 2) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 2) = "transdate_lbl2"
            'arr_varFilt_ds(F_CLBL3, 2) = "TransDateStart_lbl2"
            'arr_varFilt_ds(F_FLD3, 2) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 2) = vbNullString

            'arr_varFilt_ds(F_NAM, 3) = "TRANS_END"
            'arr_varFilt_ds(F_CONST, 3) = TRANS_END
            'arr_varFilt_ds(F_CTL, 3) = "TransDateEnd"
            'arr_varFilt_ds(F_CLBL, 3) = "TransDateEnd_lbl"
            'arr_varFilt_ds(F_FLD, 3) = "Posting Date"
            'arr_varFilt_ds(F_FLBL, 3) = "transdate_lbl" & "~" & "transdate_lbl2"  '"Posting_Date_lbl" & "~" & "Posting_Date_lbl2"
            'arr_varFilt_ds(F_CTL2, 3) = "TransDateStart"
            'arr_varFilt_ds(F_CLBL2, 3) = "TransDateStart_lbl"
            'arr_varFilt_ds(F_FLD2, 3) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 3) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 3) = "TransDateStart_lbl2"
            'arr_varFilt_ds(F_FLD3, 3) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 3) = vbNullString

            'arr_varFilt_ds(F_NAM, 4) = "ACCT_NUM"
            'arr_varFilt_ds(F_CONST, 4) = ACCT_NUM
            'arr_varFilt_ds(F_CTL, 4) = "cmbAccounts"
            'arr_varFilt_ds(F_CLBL, 4) = "cmbAccounts_lbl"
            'arr_varFilt_ds(F_FLD, 4) = "Account Number"
            'arr_varFilt_ds(F_FLBL, 4) = "accountno_lbl" & "~" & "accountno_lbl2"  '"Account_Number_lbl" & "~" & "Account_Number_lbl2"
            'arr_varFilt_ds(F_CTL2, 4) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 4) = vbNullString
            'arr_varFilt_ds(F_FLD2, 4) = "Name"
            'arr_varFilt_ds(F_FLBL2, 4) = "shortname_lbl"  '"Name_lbl"
            'arr_varFilt_ds(F_CLBL3, 4) = vbNullString
            'arr_varFilt_ds(F_FLD3, 4) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 4) = vbNullString

            'arr_varFilt_ds(F_NAM, 5) = "ASSET_NUM"
            'arr_varFilt_ds(F_CONST, 5) = ASSET_NUM
            'arr_varFilt_ds(F_CTL, 5) = "cmbAssets"
            'arr_varFilt_ds(F_CLBL, 5) = "cmbAssets_lbl"
            'arr_varFilt_ds(F_FLD, 5) = "assetno"
            'arr_varFilt_ds(F_FLBL, 5) = vbNullString  ' ** Field not displayed.
            'arr_varFilt_ds(F_CTL2, 5) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 5) = vbNullString
            'arr_varFilt_ds(F_FLD2, 5) = "Asset"
            'arr_varFilt_ds(F_FLBL2, 5) = "asset_description_lbl"  '"Asset_lbl"
            'arr_varFilt_ds(F_CLBL3, 5) = vbNullString
            'arr_varFilt_ds(F_FLD3, 5) = "CUSIP"
            'arr_varFilt_ds(F_FLBL3, 5) = "CUSIP_lbl"

            'arr_varFilt_ds(F_NAM, 6) = "CURR_NUM"
            'arr_varFilt_ds(F_CONST, 6) = CURR_NUM
            'arr_varFilt_ds(F_CTL, 6) = "cmbCurrencies"
            'arr_varFilt_ds(F_CLBL, 6) = "cmbCurrencies_lbl"
            'arr_varFilt_ds(F_FLD, 6) = "Currency"
            'arr_varFilt_ds(F_FLBL, 6) = "curr_id_lbl"  '"Currency_lbl"
            'arr_varFilt_ds(F_CTL2, 6) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 6) = vbNullString
            'arr_varFilt_ds(F_FLD2, 6) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 6) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 6) = vbNullString
            'arr_varFilt_ds(F_FLD3, 6) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 6) = vbNullString

            'arr_varFilt_ds(F_NAM, 7) = "ASSET_START"
            'arr_varFilt_ds(F_CONST, 7) = ASSET_START
            'arr_varFilt_ds(F_CTL, 7) = "AssetDateStart"
            'arr_varFilt_ds(F_CLBL, 7) = "AssetDateStart_lbl"
            'arr_varFilt_ds(F_FLD, 7) = "Trade Date"
            'arr_varFilt_ds(F_FLBL, 7) = "assetdate_lbl"  '"Trade_Date_lbl"
            'arr_varFilt_ds(F_CTL2, 7) = "AssetDateEnd"
            'arr_varFilt_ds(F_CLBL2, 7) = "AssetDateEnd_lbl"
            'arr_varFilt_ds(F_FLD2, 7) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 7) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 7) = "AssetDateStart_lbl2"
            'arr_varFilt_ds(F_FLD3, 7) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 7) = vbNullString

            'arr_varFilt_ds(F_NAM, 8) = "ASSET_END"
            'arr_varFilt_ds(F_CONST, 8) = ASSET_END
            'arr_varFilt_ds(F_CTL, 8) = "AssetDateEnd"
            'arr_varFilt_ds(F_CLBL, 8) = "AssetDateEnd_lbl"
            'arr_varFilt_ds(F_FLD, 8) = "Trade Date"
            'arr_varFilt_ds(F_FLBL, 8) = "assetdate_lbl"  '"Trade_Date_lbl"
            'arr_varFilt_ds(F_CTL2, 8) = "AssetDateStart"
            'arr_varFilt_ds(F_CLBL2, 8) = "AssetDateStart_lbl"
            'arr_varFilt_ds(F_FLD2, 8) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 8) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 8) = "AssetDateStart_lbl2"
            'arr_varFilt_ds(F_FLD3, 8) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 8) = vbNullString

            'arr_varFilt_ds(F_NAM, 9) = "PURCH_START"
            'arr_varFilt_ds(F_CONST, 9) = PURCH_START
            'arr_varFilt_ds(F_CTL, 9) = "PurchaseDateStart"
            'arr_varFilt_ds(F_CLBL, 9) = "PurchaseDateStart_lbl"
            'arr_varFilt_ds(F_FLD, 9) = "Original Trade Date"
            'arr_varFilt_ds(F_FLBL, 9) = "PurchaseDate_lbl"  '"Original_Trade_Date_lbl"
            'arr_varFilt_ds(F_CTL2, 9) = "PurchaseDateEnd"
            'arr_varFilt_ds(F_CLBL2, 9) = "PurchaseDateEnd_lbl"
            'arr_varFilt_ds(F_FLD2, 9) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 9) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 9) = "PurchaseDateStart_lbl2"
            'arr_varFilt_ds(F_FLD3, 9) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 9) = vbNullString

            'arr_varFilt_ds(F_NAM, 10) = "PURCH_END"
            'arr_varFilt_ds(F_CONST, 10) = PURCH_END
            'arr_varFilt_ds(F_CTL, 10) = "PurchaseDateEnd"
            'arr_varFilt_ds(F_CLBL, 10) = "PurchaseDateEnd_lbl"
            'arr_varFilt_ds(F_FLD, 10) = "Original Trade Date"
            'arr_varFilt_ds(F_FLBL, 10) = "PurchaseDate_lbl"  '"Original_Trade_Date_lbl"
            'arr_varFilt_ds(F_CTL2, 10) = "PurchaseDateStart"
            'arr_varFilt_ds(F_CLBL2, 10) = "PurchaseDateStart_lbl"
            'arr_varFilt_ds(F_FLD2, 10) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 10) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 10) = "PurchaseDateStart_lbl2"
            'arr_varFilt_ds(F_FLD3, 10) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 10) = vbNullString

            'arr_varFilt_ds(F_NAM, 11) = "COMM_DESC"
            'arr_varFilt_ds(F_CONST, 11) = COMM_DESC
            'arr_varFilt_ds(F_CTL, 11) = "ledger_description"
            'arr_varFilt_ds(F_CLBL, 11) = "ledger_description_lbl"
            'arr_varFilt_ds(F_FLD, 11) = "Comments"
            'arr_varFilt_ds(F_FLBL, 11) = "ledger_description_lbl"  '"Comments_lbl"
            'arr_varFilt_ds(F_CTL2, 11) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 11) = vbNullString
            'arr_varFilt_ds(F_FLD2, 11) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 11) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 11) = vbNullString
            'arr_varFilt_ds(F_FLD3, 11) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 11) = vbNullString

            'arr_varFilt_ds(F_NAM, 12) = "RECUR_ITEM"
            'arr_varFilt_ds(F_CONST, 12) = RECUR_ITEM
            'arr_varFilt_ds(F_CTL, 12) = "cmbRecurringItems"
            'arr_varFilt_ds(F_CLBL, 12) = "cmbRecurringItems_lbl"
            'arr_varFilt_ds(F_FLD, 12) = "Recurring Item"
            'arr_varFilt_ds(F_FLBL, 12) = "RecurringItem_lbl"  '"Recurring_Item_lbl"
            'arr_varFilt_ds(F_CTL2, 12) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 12) = vbNullString
            'arr_varFilt_ds(F_FLD2, 12) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 12) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 12) = vbNullString
            'arr_varFilt_ds(F_FLD3, 12) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 12) = vbNullString

            'arr_varFilt_ds(F_NAM, 13) = "REV_CODE"
            'arr_varFilt_ds(F_CONST, 13) = REV_CODE
            'arr_varFilt_ds(F_CTL, 13) = "cmbRevenueCodes"
            'arr_varFilt_ds(F_CLBL, 13) = "cmbRevenueCodes_lbl"
            'arr_varFilt_ds(F_FLD, 13) = "revcode_ID"
            'arr_varFilt_ds(F_FLBL, 13) = vbNullString  ' ** Field not displayed.
            'arr_varFilt_ds(F_CTL2, 13) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 13) = vbNullString
            'arr_varFilt_ds(F_FLD2, 13) = "Inc/Exp Codes"
            'arr_varFilt_ds(F_FLBL2, 13) = "revcode_DESC_lbl"  '"Inc/Exp_Codes_lbl"
            'arr_varFilt_ds(F_CLBL3, 13) = vbNullString
            'arr_varFilt_ds(F_FLD3, 13) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 13) = vbNullString

            'arr_varFilt_ds(F_NAM, 14) = "TAX_CODE"
            'arr_varFilt_ds(F_CONST, 14) = TAX_CODE
            'arr_varFilt_ds(F_CTL, 14) = "cmbTaxCodes"
            'arr_varFilt_ds(F_CLBL, 14) = "cmbTaxCodes_lbl"
            'arr_varFilt_ds(F_FLD, 14) = "taxcode"
            'arr_varFilt_ds(F_FLBL, 14) = vbNullString  ' ** Field not displayed.
            'arr_varFilt_ds(F_CTL2, 14) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 14) = vbNullString
            'arr_varFilt_ds(F_FLD2, 14) = "Tax Codes"
            'arr_varFilt_ds(F_FLBL2, 14) = "taxcode_description_lbl"  '"Tax_Codes_lbl"
            'arr_varFilt_ds(F_CLBL3, 14) = vbNullString
            'arr_varFilt_ds(F_FLD3, 14) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 14) = vbNullString

            'arr_varFilt(F_NAM, 15) = "LOC_NUM"
            'arr_varFilt(F_CONST, 15) = LOC_NUM
            'arr_varFilt(F_CTL, 15) = "cmbLocations"
            'arr_varFilt(F_CLBL, 15) = "cmbLocations_lbl"
            'arr_varFilt(F_FLD, 15) = "Location_Name"
            'arr_varFilt(F_FLBL, 15) = "Location_Name_lbl"
            'arr_varFilt(F_CTL2, 15) = vbNullString
            'arr_varFilt(F_CLBL2, 15) = vbNullString
            'arr_varFilt(F_FLD2, 15) = vbNullString
            'arr_varFilt(F_FLBL2, 15) = vbNullString
            'arr_varFilt(F_CLBL3, 15) = vbNullString
            'arr_varFilt(F_FLD3, 15) = vbNullString
            'arr_varFilt(F_FLBL3, 15) = vbNullString

            'arr_varFilt_ds(F_NAM, 16) = "CHK_NUM"  'CHK_NUM1  'CHK_NUM2
            'arr_varFilt_ds(F_CONST, 16) = CHK_NUM
            'arr_varFilt_ds(F_CTL, 16) = "CheckNum"
            'arr_varFilt_ds(F_CLBL, 16) = "CheckNum_lbl"
            'arr_varFilt_ds(F_FLD, 16) = "Check Number"
            'arr_varFilt_ds(F_FLBL, 16) = "CheckNum_lbl" & "~" & "CheckNum_lbl2"  '"Check_Number_lbl" & "~" & "Check_Number_lbl2"
            'arr_varFilt_ds(F_CTL2, 16) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 16) = vbNullString
            'arr_varFilt_ds(F_FLD2, 16) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 16) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 16) = vbNullString
            'arr_varFilt_ds(F_FLD3, 16) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 16) = vbNullString

            'arr_varFilt_ds(F_NAM, 17) = "JRNL_USER"
            'arr_varFilt_ds(F_CONST, 17) = JRNL_USER
            'arr_varFilt_ds(F_CTL, 17) = "cmbUsers"
            'arr_varFilt_ds(F_CLBL, 17) = "cmbUsers_lbl"
            'arr_varFilt_ds(F_FLD, 17) = "User"
            'arr_varFilt_ds(F_FLBL, 17) = "journal_USER_lbl"  '"User_lbl"
            'arr_varFilt_ds(F_CTL2, 17) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 17) = vbNullString
            'arr_varFilt_ds(F_FLD2, 17) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 17) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 17) = vbNullString
            'arr_varFilt_ds(F_FLD3, 17) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 17) = vbNullString

            'arr_varFilt_ds(F_NAM, 18) = "POSTED_START"
            'arr_varFilt_ds(F_CONST, 18) = POSTED_START
            'arr_varFilt_ds(F_CTL, 18) = "PostedDateStart"
            'arr_varFilt_ds(F_CLBL, 18) = "PostedDateStart_lbl"
            'arr_varFilt_ds(F_FLD, 18) = "Date Posted"
            'arr_varFilt_ds(F_FLBL, 18) = "posted_lbl"  '"Date_Posted_lbl"
            'arr_varFilt_ds(F_CTL2, 18) = "PostedDateEnd"
            'arr_varFilt_ds(F_CLBL2, 18) = "PostedDateEnd_lbl"
            'arr_varFilt_ds(F_FLD2, 18) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 18) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 18) = "PostedDateStart_lbl2"
            'arr_varFilt_ds(F_FLD3, 18) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 18) = vbNullString

            'arr_varFilt_ds(F_NAM, 19) = "POSTED_END"
            'arr_varFilt_ds(F_CONST, 19) = POSTED_END
            'arr_varFilt_ds(F_CTL, 19) = "PostedDateEnd"
            'arr_varFilt_ds(F_CLBL, 19) = "PostedDateEnd_lbl"
            'arr_varFilt_ds(F_FLD, 19) = "Date Posted"
            'arr_varFilt_ds(F_FLBL, 19) = "posted_lbl"  '"Date_Posted_lbl"
            'arr_varFilt_ds(F_CTL2, 19) = "PostedDateStart"
            'arr_varFilt_ds(F_CLBL2, 19) = "PostedDateStart_lbl"
            'arr_varFilt_ds(F_FLD2, 19) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 19) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 19) = "PostedDateStart_lbl2"
            'arr_varFilt_ds(F_FLD3, 19) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 19) = vbNullString

            'arr_varFilt_ds(F_NAM, 20) = "HIDDEN_TRX1"  'HIDDEN_TRX2
            'arr_varFilt_ds(F_CONST, 20) = HIDDEN_TRX1
            'arr_varFilt_ds(F_CTL, 20) = "opgHidden"
            'arr_varFilt_ds(F_CLBL, 20) = "opgHidden_lbl"
            'arr_varFilt_ds(F_FLD, 20) = "Hidden"
            'arr_varFilt_ds(F_FLBL, 20) = "ledger_HIDDEN_lbl"  '"Hidden_lbl"
            'arr_varFilt_ds(F_CTL2, 20) = vbNullString
            'arr_varFilt_ds(F_CLBL2, 20) = vbNullString
            'arr_varFilt_ds(F_FLD2, 20) = vbNullString
            'arr_varFilt_ds(F_FLBL2, 20) = vbNullString
            'arr_varFilt_ds(F_CLBL3, 20) = vbNullString
            'arr_varFilt_ds(F_FLD3, 20) = vbNullString
            'arr_varFilt_ds(F_FLBL3, 20) = vbNullString

480       End Select
490     Next  ' ** lngX.

EXITP:
500     Set rst = Nothing
510     Set qdf = Nothing
520     Set dbs = Nothing
530     Exit Sub

ERRH:
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Sub

Public Function FilterRecs_GetArr(intMode As Integer) As Variant
' ** intMode is opgView.

600   On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_GetArr"

        Dim varRetVal As Variant

610     Select Case intMode
        Case 1
620       If IsEmpty(arr_varFilt) = True Then
630         FilterRecs_Load  ' ** Procedure: Above.
640         DoEvents
650       End If
660       varRetVal = arr_varFilt
670     Case 2
680       If IsEmpty(arr_varFilt_ds) = True Then
690         FilterRecs_Load  ' ** Procedure: Above.
700         DoEvents
710       End If
720       varRetVal = arr_varFilt_ds
730     End Select

EXITP:
740     FilterRecs_GetArr = varRetVal
750     Exit Function

ERRH:
760     varRetVal = Null
770     Select Case ERR.Number
        Case Else
780       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
790     End Select
800     Resume EXITP

End Function

Public Sub ShowFields(strProc As String, blnShow As Boolean, frm As Access.Form)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "ShowFields"

910     With frm
920       Set frmCrit = .frmTransaction_Audit_Sub_Criteria.Form
930       With frmCrit
940         Select Case strProc
            Case "ckgFlds_chkJournalno_AfterUpdate"
950           Select Case blnShow
              Case True
960             .journalno.BackStyle = acBackStyleNormal
970             .journalno.BorderColor = CLR_LTBLU2
980           Case False
990             .journalno.BackStyle = acBackStyleTransparent
1000            .journalno.BorderColor = WIN_CLR_DISR
1010          End Select
1020        Case "ckgFlds_chkJournalType_AfterUpdate"
1030          Select Case blnShow
              Case True
1040            .cmbJournalType_lbl.ForeColor = CLR_VDKGRY
1050            .cmbJournalType_lbl_dim_hi.Visible = False
1060            .cmbJournalType1.BackStyle = acBackStyleNormal
1070            .cmbJournalType1.BorderColor = CLR_LTBLU2
1080          Case False
1090            .cmbJournalType_lbl.ForeColor = WIN_CLR_DISF
1100            .cmbJournalType_lbl_dim_hi.Visible = True
1110            .cmbJournalType1.BackStyle = acBackStyleTransparent
1120            .cmbJournalType1.BorderColor = WIN_CLR_DISR
1130          End Select
              'THE OTHERS?
1140        Case "ckgFlds_chkTransDate_AfterUpdate"
1150          Select Case blnShow
              Case True
1160            .TransDateStart_lbl2.ForeColor = CLR_VDKGRY
1170            .TransDateStart_lbl2_dim_hi.Visible = False
1180            .TransDateStart.BackStyle = acBackStyleNormal
1190            .TransDateStart.BorderColor = CLR_LTBLU2
1200            .TransDateEnd.BackStyle = acBackStyleNormal
1210            .TransDateEnd.BorderColor = CLR_LTBLU2
1220          Case False
1230            .TransDateStart_lbl2.ForeColor = WIN_CLR_DISF
1240            .TransDateStart_lbl2_dim_hi.Visible = True
1250            .TransDateStart.BackStyle = acBackStyleTransparent
1260            .TransDateStart.BorderColor = WIN_CLR_DISR
1270            .TransDateEnd.BackStyle = acBackStyleTransparent
1280            .TransDateEnd.BorderColor = WIN_CLR_DISR
1290          End Select
1300        Case "ckgFlds_chkAccountNo_AfterUpdate", "ckgFlds_chkShortName_AfterUpdate"
1310          Select Case blnShow
              Case True
1320            .cmbAccounts.BackStyle = acBackStyleNormal
1330            .cmbAccounts.BorderColor = CLR_LTBLU2
1340          Case False
1350            .cmbAccounts.BackStyle = acBackStyleTransparent
1360            .cmbAccounts.BorderColor = WIN_CLR_DISR
1370          End Select
1380        Case "ckgFlds_chkCusip_AfterUpdate", "ckgFlds_chkAssetDescription_AfterUpdate"
1390          Select Case blnShow
              Case True
1400            .cmbAssets.BackStyle = acBackStyleNormal
1410            .cmbAssets.BorderColor = CLR_LTBLU2
1420          Case False
1430            .cmbAssets.BackStyle = acBackStyleTransparent
1440            .cmbAssets.BorderColor = WIN_CLR_DISR
1450          End Select
1460        Case "ckgFlds_chkShareFace_AfterUpdate"
              ' ** No search fields.
1470        Case "ckgFlds_chkICash_AfterUpdate"
              ' ** No search fields.
1480        Case "ckgFlds_chkPCash_AfterUpdate"
              ' ** No search fields.
1490        Case "ckgFlds_chkCost_AfterUpdate"
              ' ** No search fields.
1500        Case "ckgFlds_chkCurrID_AfterUpdate"
1510          Select Case blnShow
              Case True
1520            .cmbCurrencies.Enabled = True
1530            .cmbCurrencies.BackStyle = acBackStyleNormal
1540            .cmbCurrencies.BorderColor = CLR_LTBLU2
1550          Case False
1560            .cmbCurrencies.Enabled = False
1570            .cmbCurrencies.BackStyle = acBackStyleTransparent
1580            .cmbCurrencies.BorderColor = WIN_CLR_DISR
1590          End Select
1600        Case "ckgFlds_chkAssetDate_AfterUpdate"
1610          Select Case blnShow
              Case True
1620            .AssetDateStart_lbl2.ForeColor = CLR_VDKGRY
1630            .AssetDateStart_lbl2_dim_hi.Visible = False
1640            .AssetDateStart.BackStyle = acBackStyleNormal
1650            .AssetDateStart.BorderColor = CLR_LTBLU2
1660            .AssetDateEnd.BackStyle = acBackStyleNormal
1670            .AssetDateEnd.BorderColor = CLR_LTBLU2
1680          Case False
1690            .AssetDateStart_lbl2.ForeColor = WIN_CLR_DISF
1700            .AssetDateStart_lbl2_dim_hi.Visible = True
1710            .AssetDateStart.BackStyle = acBackStyleTransparent
1720            .AssetDateStart.BorderColor = WIN_CLR_DISR
1730            .AssetDateEnd.BackStyle = acBackStyleTransparent
1740            .AssetDateEnd.BorderColor = WIN_CLR_DISR
1750          End Select
1760        Case "ckgFlds_chkPurchaseDate_AfterUpdate"
1770          Select Case blnShow
              Case True
1780            .PurchaseDateStart_lbl2.ForeColor = CLR_VDKGRY
1790            .PurchaseDateStart_lbl2_dim_hi.Visible = False
1800            .PurchaseDateStart.BackStyle = acBackStyleNormal
1810            .PurchaseDateStart.BorderColor = CLR_LTBLU2
1820            .PurchaseDateEnd.BackStyle = acBackStyleNormal
1830            .PurchaseDateEnd.BorderColor = CLR_LTBLU2
1840          Case False
1850            .PurchaseDateStart_lbl2.ForeColor = WIN_CLR_DISF
1860            .PurchaseDateStart_lbl2_dim_hi.Visible = True
1870            .PurchaseDateStart.BackStyle = acBackStyleTransparent
1880            .PurchaseDateStart.BorderColor = WIN_CLR_DISR
1890            .PurchaseDateEnd.BackStyle = acBackStyleTransparent
1900            .PurchaseDateEnd.BorderColor = WIN_CLR_DISR
1910          End Select
1920        Case "ckgFlds_chkLedgerDescription_AfterUpdate"
1930          Select Case blnShow
              Case True
1940            .ledger_description.BackStyle = acBackStyleNormal
1950            .ledger_description.BorderColor = CLR_LTBLU2
1960          Case False
1970            .ledger_description.BackStyle = acBackStyleTransparent
1980            .ledger_description.BorderColor = WIN_CLR_DISR
1990          End Select
2000        Case "ckgFlds_chkRecurringItem_AfterUpdate"
2010          Select Case blnShow
              Case True
2020            .cmbRecurringItems.BackStyle = acBackStyleNormal
2030            .cmbRecurringItems.BorderColor = CLR_LTBLU2
2040          Case False
2050            .cmbRecurringItems.BackStyle = acBackStyleTransparent
2060            .cmbRecurringItems.BorderColor = WIN_CLR_DISR
2070          End Select
2080        Case "ckgFlds_chkRevCodeDesc_AfterUpdate"
2090          Select Case blnShow
              Case True
2100            .cmbRevenueCodes.BackStyle = acBackStyleNormal
2110            .cmbRevenueCodes.BorderColor = CLR_LTBLU2
2120          Case False
2130            .cmbRevenueCodes.BackStyle = acBackStyleTransparent
2140            .cmbRevenueCodes.BorderColor = WIN_CLR_DISR
2150          End Select
2160        Case "ckgFlds_chkRevCodeTypeDescription_AfterUpdate"
              ' ** Disabling does it all.
2170        Case "ckgFlds_chkTaxCodeDescription_AfterUpdate"
2180          Select Case blnShow
              Case True
2190            .cmbTaxCodes.BackStyle = acBackStyleNormal
2200            .cmbTaxCodes.BorderColor = CLR_LTBLU2
2210          Case False
2220            .cmbTaxCodes.BackStyle = acBackStyleTransparent
2230            .cmbTaxCodes.BorderColor = WIN_CLR_DISR
2240          End Select
2250        Case "ckgFlds_chkTaxCodeTypeDescription_AfterUpdate"
              ' ** Disabling does it all.
2260        Case "ckgFlds_chkLocationName_AfterUpdate"
2270          Select Case blnShow
              Case True
2280            .cmbLocations.BackStyle = acBackStyleNormal
2290            .cmbLocations.BorderColor = CLR_LTBLU2
2300          Case False
2310            .cmbLocations.BackStyle = acBackStyleTransparent
2320            .cmbLocations.BorderColor = WIN_CLR_DISR
2330          End Select
2340        Case "ckgFlds_chkCheckNum_AfterUpdate"
2350          Select Case blnShow
              Case True
2360            .CheckNum.Enabled = True
2370            .CheckNum.BackStyle = acBackStyleNormal
2380            .CheckNum.BorderColor = CLR_LTBLU2
2390          Case False
2400            .CheckNum.Enabled = False
2410            .CheckNum.BackStyle = acBackStyleTransparent
2420            .CheckNum.BorderColor = WIN_CLR_DISR
2430          End Select
2440        Case "ckgFlds_chkJournalUser_AfterUpdate"
2450          Select Case blnShow
              Case True
2460            .cmbUsers.BackStyle = acBackStyleNormal
2470            .cmbUsers.BorderColor = CLR_LTBLU2
2480          Case False
2490            .cmbUsers.BackStyle = acBackStyleTransparent
2500            .cmbUsers.BorderColor = WIN_CLR_DISR
2510          End Select
2520        Case "ckgFlds_chkPosted_AfterUpdate"
2530          Select Case blnShow
              Case True
2540            .PostedDateStart_lbl2.ForeColor = CLR_VDKGRY
2550            .PostedDateStart_lbl2_dim_hi.Visible = False
2560            .PostedDateStart.BackStyle = acBackStyleNormal
2570            .PostedDateStart.BorderColor = CLR_LTBLU2
2580            .PostedDateEnd.BackStyle = acBackStyleNormal
2590            .PostedDateEnd.BorderColor = CLR_LTBLU2
2600          Case False
2610            .PostedDateStart_lbl2.ForeColor = WIN_CLR_DISF
2620            .PostedDateStart_lbl2_dim_hi.Visible = True
2630            .PostedDateStart.BackStyle = acBackStyleTransparent
2640            .PostedDateStart.BorderColor = WIN_CLR_DISR
2650            .PostedDateEnd.BackStyle = acBackStyleTransparent
2660            .PostedDateEnd.BorderColor = WIN_CLR_DISR
2670          End Select
2680        Case "ckgFlds_chkLedgerHidden_AfterUpdate"
2690          Select Case blnShow
              Case True
                '.opgHidden_optInclude_lbl2.ForeColor = CLR_VDKGRY
                '.opgHidden_optInclude_lbl2_dim_hi.Visible = False
                '.opgHidden_optExclude_lbl2.ForeColor = CLR_VDKGRY
                '.opgHidden_optExclude_lbl2_dim_hi.Visible = False
                '.opgHidden_optOnly_lbl2.ForeColor = CLR_VDKGRY
                '.opgHidden_optOnly_lbl2_dim_hi.Visible = False
2700          Case False
                '.opgHidden_optInclude_lbl2.ForeColor = WIN_CLR_DISF
                '.opgHidden_optInclude_lbl2_dim_hi.Visible = True
                '.opgHidden_optExclude_lbl2.ForeColor = WIN_CLR_DISF
                '.opgHidden_optExclude_lbl2_dim_hi.Visible = True
                '.opgHidden_optOnly_lbl2.ForeColor = WIN_CLR_DISF
                '.opgHidden_optOnly_lbl2_dim_hi.Visible = True
2710          End Select
2720        End Select
2730      End With
2740      DoEvents
2750    End With

        ' ** ckgFlds_chkJournalno_AfterUpdate
        ' ** ckgFlds_chkJournalType_AfterUpdate
        ' ** ckgFlds_chkTransDate_AfterUpdate
        ' ** ckgFlds_chkAccountNo_AfterUpdate
        ' ** ckgFlds_chkShortName_AfterUpdate
        ' ** ckgFlds_chkCusip_AfterUpdate
        ' ** ckgFlds_chkAssetDescription_AfterUpdate
        ' ** ckgFlds_chkShareFace_AfterUpdate
        ' ** ckgFlds_chkICash_AfterUpdate
        ' ** ckgFlds_chkPCash_AfterUpdate
        ' ** ckgFlds_chkCost_AfterUpdate
        ' ** ckgFlds_chkCurrID_AfterUpdate
        ' ** ckgFlds_chkAssetDate_AfterUpdate
        ' ** ckgFlds_chkPurchaseDate_AfterUpdate
        ' ** ckgFlds_chkLedgerDescription_AfterUpdate
        ' ** ckgFlds_chkRecurringItem_AfterUpdate
        ' ** ckgFlds_chkRevCodeDesc_AfterUpdate
        ' ** ckgFlds_chkRevCodeTypeDescription_AfterUpdate
        ' ** ckgFlds_chkTaxCodeDescription_AfterUpdate
        ' ** ckgFlds_chkTaxCodeTypeDescription_AfterUpdate
        ' ** ckgFlds_chkLocationName_AfterUpdate
        ' ** ckgFlds_chkCheckNum_AfterUpdate
        ' ** ckgFlds_chkJournalUser_AfterUpdate
        ' ** ckgFlds_chkPosted_AfterUpdate
        ' ** ckgFlds_chkLedgerHidden_AfterUpdate

EXITP:
2760    Exit Sub

ERRH:
2770    Select Case ERR.Number
        Case Else
2780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2790    End Select
2800    Resume EXITP

End Sub

Public Sub ShowFields_Set(lngFrmFlds As Long, arr_varFrmFld As Variant, arr_varFrmFld_ds As Variant, frm As Access.Form)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowFields_Set"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim lngTmp01 As Long
        Dim lngX As Long

2910    With frm

2920      lngTmp01 = .ckgFlds_cmd.Top

          ' ** Coordinate the arrays.
2930      For lngX = 0& To (lngFrmFlds - 1&)
2940        Select Case arr_varFrmFld(FM_FLD_NAM, lngX)
            Case "journalno"
2950          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalno"
2960        Case "journaltype"
2970          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalType"
2980        Case "transdate"
2990          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkTransDate"
3000        Case "accountno"
3010          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkAccountNo"
3020        Case "shortname"
3030          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkShortName"
3040        Case "cusip"
3050          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkCusip"
3060        Case "asset_description"
3070          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkAssetDescription"
3080        Case "shareface"
3090          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkShareFace"
3100        Case "icash"
3110          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkICash"
3120        Case "pcash"
3130          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkPCash"
3140        Case "cost"
3150          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkCost"
3160        Case "curr_id"
3170          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkCurrID"
3180        Case "assetdate"
3190          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkAssetDate"
3200        Case "PurchaseDate"
3210          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkPurchaseDate"
3220        Case "ledger_description"
3230          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkLedgerDescription"
3240        Case "RecurringItem"
3250          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkRecurringItem"
3260        Case "revcode_DESC"
3270          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkRevCodeDesc"
3280        Case "revcode_TYPE_Description"
3290          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkRevCodeTypeDescription"
3300        Case "taxcode_description"
3310          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkTaxCodeDescription"
3320        Case "taxcode_type_description"
3330          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkTaxCodeTypeDescription"
3340        Case "Location_Name"
3350          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkLocationName"
3360        Case "CheckNum"
3370          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkCheckNum"
3380        Case "journal_USER"
3390          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalUser"
3400        Case "posted"
3410          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkPosted"
3420        Case "ledger_HIDDEN"
3430          arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkLedgerHidden"
3440        End Select
3450        arr_varFrmFld_ds(FM_VIEWCHK, lngX) = arr_varFrmFld(FM_VIEWCHK, lngX)
3460        arr_varFrmFld(FM_TOPO, lngX) = .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)).Top
3470        arr_varFrmFld(FM_TOPC, lngX) = lngTmp01
3480        arr_varFrmFld_ds(FM_TOPO, lngX) = .Controls(arr_varFrmFld_ds(FM_VIEWCHK, lngX)).Top
3490        arr_varFrmFld_ds(FM_TOPC, lngX) = lngTmp01
3500      Next  ' ** lngX.

3510      Set frm1 = .frmTransaction_Audit_Sub.Form
3520      Set frm2 = .frmTransaction_Audit_Sub_ds.Form

          ' ** Set field visibility.
3530      For lngX = 0& To (lngFrmFlds - 1&)
3540        If .Controls(arr_varFrmFld(FM_VIEWCHK, lngX)) = False Then
3550          arr_varFrmFld(FM_FLD_VIS, lngX) = CBool(False)
3560          DoEvents
3570        End If
3580      Next  ' ** lngX.

3590    End With

EXITP:
3600    Set frm1 = Nothing
3610    Set frm2 = Nothing
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Public Sub SubFrmGfx_Load(frm As Access.Form)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "SubFrmGfx_Load"

        Dim frmSub As Access.Form, frmGfx As Access.Form, ctl As Access.Control

3710    With frm
3720      Set frmSub = .frmTransaction_Audit_Sub.Form
3730      Set frmGfx = .frmTransaction_Audit_Sub_Graphics.Form
3740      For Each ctl In .Detail.Controls  ' ** These are on frmTransaction_Audit.
3750        With ctl
3760          If .ControlType = acTextBox Then
3770            If Right(.Name, 4) = "_img" Or Right(.Name, 8) = "_img_dis" Then
3780              With frmSub
3790  On Error Resume Next
3800                .Controls(ctl.Name).Value = ctl.Value
3810                If ERR.Number <> 0 Then
3820  On Error GoTo ERRH
                      'Debug.Print "'1 " & ctl.Name
3830                Else
3840  On Error GoTo ERRH
3850                End If
3860              End With
3870            End If
3880          End If
3890        End With  ' ** ctl.
3900      Next  ' ** ctl.
3910      For Each ctl In frmGfx.Detail.Controls
3920        With ctl
3930          If .ControlType = acTextBox Then
3940            If Right(.Name, 4) = "_img" Or Right(.Name, 8) = "_img_dis" Then
3950              With frmSub
3960  On Error Resume Next
3970                .Controls(ctl.Name).Value = ctl.Value
3980                If ERR.Number <> 0 Then
3990  On Error GoTo ERRH
                      'Debug.Print "'2 " & ctl.Name
4000                Else
4010  On Error GoTo ERRH
4020                End If
4030              End With
4040            End If
4050          End If
4060        End With  ' ** ctl.
4070      Next  ' ** ctl.
4080    End With

EXITP:
4090    Set ctl = Nothing
4100    Set frmGfx = Nothing
4110    Set frmSub = Nothing
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Public Sub TransAudit_Resize(blnShorter As Boolean, lngDetail_Height_New As Long, lngDetail_Height_Diff As Long, lngSub_Height As Long, lngSelectBox_Top As Long, lngSelectLbl_Top As Long, lngSelectLegendLbl_Top As Long, lngSelectLbl2_Top As Long, lngNarrowFont_Top As Long, lngShareface_Top As Long, lngSelectAll_Top As Long, lngSelectNone_Top As Long, lngChkBoxLbl_TopOffset As Long, lngCkgFldsVLine_Top As Long, lngFrmFlds As Long, arr_varFrmFld As Variant, frm As Access.Form)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "TransAudit_Resize"

        Dim lngTmp01 As Long
        Dim lngX As Long

4210    If gblnClosing = False Then
4220      With frm

4230        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
4240          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4250        End If

4260        DoCmd.SelectObject acForm, frm.Name, False
4270        DoEvents

4280        Select Case blnShorter
            Case True
4290          .frmTransaction_Audit_Sub.Height = (lngSub_Height - lngDetail_Height_Diff)
4300          .frmTransaction_Audit_Sub_ds.Height = (lngSub_Height - lngDetail_Height_Diff)
4310          .frmTransaction_Audit_Sub_box.Height = (.frmTransaction_Audit_Sub.Height + (6& * lngTpp))
4320          .Nav_hline01.Top = (.frmTransaction_Audit_Sub.Top + .frmTransaction_Audit_Sub.Height)
4330          .Nav_hline02.Top = (.Nav_hline01.Top + lngTpp)
4340          .Nav_hline03.Top = .Nav_hline01.Top
4350          .Nav_vline01.Top = .Nav_hline01.Top
4360          .Nav_vline02.Top = .Nav_hline01.Top
4370          .Nav_vline03.Top = .Nav_hline01.Top
4380          .Nav_vline04.Top = .Nav_hline01.Top
4390          .Nav_box01.Top = (.frmTransaction_Audit_Sub.Top + .frmTransaction_Audit_Sub.Height)
4400          .ShortcutMenu_lbl.Top = ((.frmTransaction_Audit_Sub.Top + .frmTransaction_Audit_Sub.Height) + (6& * lngTpp))
4410          .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
4420          .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
4430          .Detail_hline01.Top = (.ShortcutMenu_lbl.Top + .ShortcutMenu_lbl.Height)
4440          .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
4450          .Detail_vline01.Top = .Detail_hline01.Top
4460          .Detail_vline02.Top = .Detail_hline01.Top
4470          .cmdSelect_box.Top = (lngSelectBox_Top - lngDetail_Height_Diff)
4480          .cmdSelect_lbl.Top = (lngSelectLbl_Top - lngDetail_Height_Diff)
4490          .cmdSelect_lbl3.Top = (.cmdSelect_box.Top - (2& * lngTpp))
4500          .cmdSelect_lbl3_cmd.Top = .cmdSelect_lbl3.Top
4510          .cmdSelect_Legend_lbl.Top = (lngSelectLegendLbl_Top - lngDetail_Height_Diff)
4520          .cmdSelect_Legend_box.Top = (.cmdSelect_Legend_lbl.Top - lngTpp)
4530          .cmdSelect_lbl2.Top = (lngSelectLbl2_Top - lngDetail_Height_Diff)
4540          .chkNarrowFont_box.Top = .cmdSelect_box.Top
4550          .chkNarrowFont.Top = (lngNarrowFont_Top - lngDetail_Height_Diff)
4560          .chkNarrowFont_lbl.Top = (.chkNarrowFont.Top - lngTpp)
4570          .chkIncludeSharefaceTot.Top = (lngShareface_Top - lngDetail_Height_Diff)
4580          .chkIncludeSharefaceTot_lbl.Top = (.chkIncludeSharefaceTot.Top - lngTpp)
4590          .chkPageOf.Top = .chkNarrowFont.Top
4600          .chkPageOf_lbl.Top = .chkNarrowFont_lbl.Top
4610          .chkSaveSizePos.Top = .chkIncludeSharefaceTot.Top
4620          .chkSaveSizePos_lbl.Top = .chkIncludeSharefaceTot_lbl.Top
4630          .chkNarrowFont_vline01.Top = (.cmdSelect_box.Top + lngTpp)
4640          .chkNarrowFont_vline02.Top = .chkNarrowFont_vline01.Top
4650          .cmdSelect_Legend_Tgl_lbl1.Top = (.cmdSelect_Legend_lbl.Top - lngTpp)
4660          .cmdSelect_Legend_Tgl_lbl2.Top = (.cmdSelect_Legend_Tgl_lbl1.Top + (16& * lngTpp))
4670          .cmdSelect_Legend_tgl_on_raised_img.Top = .chkPageOf.Top
4680          .cmdSelect_Legend_tgl_on_raised_focus_img.Top = .cmdSelect_Legend_tgl_on_raised_img.Top
4690          .cmdSelect_Legend_tgl_on.Top = .cmdSelect_Legend_tgl_on_raised_img.Top
4700          .cmdSelect_Legend_tgl_off_raised_img.Top = .chkSaveSizePos_lbl.Top
4710          .cmdSelect_Legend_tgl_off_raised_focus_img.Top = .cmdSelect_Legend_tgl_off_raised_img.Top
4720          .cmdSelect_Legend_tgl_off.Top = .cmdSelect_Legend_tgl_off_raised_img.Top
4730          .cmdSelect_Legend_box.Top = .cmdSelect_Legend_Tgl_lbl1.Top
4740          .cmdSelect_Legend_vline01.Top = (.cmdSelect_Legend_tgl_on_raised_img.Top + lngTpp)
4750          .cmdSelect_Legend_vline02.Top = .cmdSelect_Legend_vline01.Top
4760          .cmdSelectAll.Top = (lngSelectAll_Top - lngDetail_Height_Diff)
4770          .cmdSelectAll_raised_img.Top = .cmdSelectAll.Top
4780          .cmdSelectAll_raised_semifocus_dots_img.Top = .cmdSelectAll.Top
4790          .cmdSelectAll_raised_focus_img.Top = .cmdSelectAll.Top
4800          .cmdSelectAll_raised_focus_dots_img.Top = .cmdSelectAll.Top
4810          .cmdSelectAll_sunken_focus_dots_img.Top = .cmdSelectAll.Top
4820          .cmdSelectAll_raised_img_dis.Top = .cmdSelectAll.Top
4830          .cmdSelectNone.Top = (lngSelectNone_Top - lngDetail_Height_Diff)
4840          .cmdSelectNone_raised_img.Top = .cmdSelectNone.Top
4850          .cmdSelectNone_raised_semifocus_dots_img.Top = .cmdSelectNone.Top
4860          .cmdSelectNone_raised_focus_img.Top = .cmdSelectNone.Top
4870          .cmdSelectNone_raised_focus_dots_img.Top = .cmdSelectNone.Top
4880          .cmdSelectNone_sunken_focus_dots_img.Top = .cmdSelectNone.Top
4890          .cmdSelectNone_raised_img_dis.Top = .cmdSelectNone.Top
4900          lngTmp01 = ((.cmdSelect_box.Top + .cmdSelect_box.Height) + (8& * lngTpp))
4910          .Detail_hline03.Top = lngTmp01
4920          .Detail_hline04.Top = (lngTmp01 + lngTpp)
4930          .Detail_vline03.Top = .Detail_hline03.Top
4940          .Detail_vline04.Top = .Detail_hline03.Top
              ' ** Differences are from the control's original position, not its current position.
4950          .ckgFlds_hline01.Top = (.Detail_hline04.Top + (8& * lngTpp))
4960          .ckgFlds_hline02.Top = (.ckgFlds_hline01.Top + lngTpp)
4970          .ckgFlds_vline01.Top = .ckgFlds_hline02.Top
4980          .ckgFlds_vline02.Top = (.ckgFlds_vline01.Top + lngTpp)
4990          .ckgFlds_vline03.Top = .ckgFlds_vline02.Top
5000          .ckgFlds_vline04.Top = .ckgFlds_vline01.Top
5010          .ckgFlds_lbl.Top = .ckgFlds_hline02.Top
5020          .ckgFlds_box2.Top = (.ckgFlds_lbl.Top + lngTpp)
5030          .ckgFlds_cmd.Top = (.ckgFlds_box2.Top + lngTpp)
5040          .ckgFlds_cmd_raised_img.Top = .ckgFlds_cmd.Top
5050          .ckgFlds_cmd_raised_semifocus_dots_img.Top = .ckgFlds_cmd.Top
5060          .ckgFlds_cmd_raised_focus_img.Top = .ckgFlds_cmd.Top
5070          .ckgFlds_cmd_raised_focus_dots_img.Top = .ckgFlds_cmd.Top
5080          .ckgFlds_cmd_sunken_focus_dots_img.Top = .ckgFlds_cmd.Top
5090          .ckgFlds_cmd_raised_img_dis.Top = .ckgFlds_cmd.Top
5100          If ((.ckgFlds_vline01.Top + .ckgFlds_vline01.Height) + .ckgFlds_box.Height) > .Detail.Height Then
                ' ** There was an error here, although there shouldn't have been.
5110            .Detail.Height = ((.ckgFlds_vline01.Top + .ckgFlds_vline01.Height) + .ckgFlds_box.Height)
5120          End If
5130          .ckgFlds_box.Top = (.ckgFlds_vline01.Top + .ckgFlds_vline01.Height)
5140          .ckgFlds_hline05.Top = .ckgFlds_box.Top
5150          lngTmp01 = ((.ckgFlds_cmd.Top + .ckgFlds_cmd.Height) + lngTpp)
5160          .ckgFlds_hline03.Top = lngTmp01
5170          .ckgFlds_hline04.Top = (.ckgFlds_hline03.Top + lngTpp)
5180          .ckgFlds_vline05.Top = (.ckgFlds_vline01.Top + .ckgFlds_vline01.Height)
5190          .ckgFlds_vline06.Top = (.ckgFlds_vline05.Top + lngTpp)
5200          .ckgFlds_vline07.Top = .ckgFlds_vline06.Top
5210          .ckgFlds_vline08.Top = .ckgFlds_vline05.Top
5220          Select Case .ckgFlds
              Case True
5230            .ckgFlds_vline09.Top = (lngCkgFldsVLine_Top - lngDetail_Height_Diff)
5240            For lngX = 0& To (lngFrmFlds - 1&)
5250              If arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalno" Then
5260                .ckgFlds_chkJournalno.Top = (arr_varFrmFld(FM_TOPO, lngX) - lngDetail_Height_Diff)
5270                .ckgFlds_chkJournalno_lbl.Top = (.ckgFlds_chkJournalno.Top - lngChkBoxLbl_TopOffset)
5280              ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalType" Then
5290                .ckgFlds_chkJournalType.Top = (arr_varFrmFld(FM_TOPO, lngX) - lngDetail_Height_Diff)
5300                .ckgFlds_chkJournalType_lbl.Top = (.ckgFlds_chkJournalType.Top - lngChkBoxLbl_TopOffset)
5310              ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkTransDate" Then
5320                .ckgFlds_chkTransDate.Top = (arr_varFrmFld(FM_TOPO, lngX) - lngDetail_Height_Diff)
5330                .ckgFlds_chkTransDate_lbl.Top = (.ckgFlds_chkTransDate.Top - lngChkBoxLbl_TopOffset)
5340              ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkAccountNo" Then
5350                .ckgFlds_chkAccountNo.Top = (arr_varFrmFld(FM_TOPO, lngX) - lngDetail_Height_Diff)
5360                .ckgFlds_chkAccountNo_lbl.Top = (.ckgFlds_chkAccountNo.Top - lngChkBoxLbl_TopOffset)
5370                Exit For
5380              End If
5390            Next
5400          Case False
                'lngTmp01 = (((lngSelectBox_Top + .cmdSelect_box.Height) + (8& * lngTpp)) + lngTpp)  ' ** Original Detail_hline04.Top.
                'lngTmp01 = ((((lngTmp01 + (8& * lngTpp)) + lngTpp) + lngTpp) + lngTpp)  ' ** Original ckgFlds_cmd.Top.
                'lngTmp01 = (lngTmp01 - .ckgFlds_cmd.Top)  ' ** Difference in .ckgFlds_cmd.Top position.
5410            .ckgFlds_vline09.Top = .ckgFlds_cmd.Top
5420            For lngX = 0& To (lngFrmFlds - 1&)
5430              If arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalno" Then
5440                .ckgFlds_chkJournalno.Top = (arr_varFrmFld(FM_TOPC, lngX) - lngDetail_Height_Diff)
5450                .ckgFlds_chkJournalno_lbl.Top = (.ckgFlds_chkJournalno.Top - lngChkBoxLbl_TopOffset)
5460              ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalType" Then
5470                .ckgFlds_chkJournalType.Top = (arr_varFrmFld(FM_TOPC, lngX) - lngDetail_Height_Diff)
5480                .ckgFlds_chkJournalType_lbl.Top = .ckgFlds_chkJournalType.Top - lngChkBoxLbl_TopOffset
5490              ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkTransDate" Then
5500                .ckgFlds_chkTransDate.Top = (arr_varFrmFld(FM_TOPC, lngX) - lngDetail_Height_Diff)
5510                .ckgFlds_chkTransDate_lbl.Top = .ckgFlds_chkTransDate.Top - lngChkBoxLbl_TopOffset
5520              ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkAccountNo" Then
5530                .ckgFlds_chkAccountNo.Top = (arr_varFrmFld(FM_TOPC, lngX) - lngDetail_Height_Diff)
5540                .ckgFlds_chkAccountNo_lbl.Top = (.ckgFlds_chkAccountNo.Top - lngChkBoxLbl_TopOffset)
5550                Exit For
5560              End If
5570            Next
5580          End Select
5590          .ckgFlds_vline10.Top = .ckgFlds_vline09.Top
5600          .ckgFlds_vline11.Top = .ckgFlds_vline09.Top
5610          .ckgFlds_vline12.Top = .ckgFlds_vline09.Top
5620          .ckgFlds_vline13.Top = .ckgFlds_vline09.Top
5630          .ckgFlds_vline14.Top = .ckgFlds_vline09.Top
5640          .ckgFlds_vline15.Top = .ckgFlds_vline09.Top
5650          .ckgFlds_vline16.Top = .ckgFlds_vline09.Top
5660          .ckgFlds_vline17.Top = .ckgFlds_vline09.Top
5670          .ckgFlds_vline18.Top = .ckgFlds_vline09.Top
5680          .ckgFlds_vline19.Top = .ckgFlds_vline09.Top
5690          .ckgFlds_vline20.Top = .ckgFlds_vline09.Top
              '.ckgFlds_vline21.Top = .ckgFlds_vline09.Top
              '.ckgFlds_vline22.Top = .ckgFlds_vline09.Top
              '.ckgFlds_chkAccountNo.Top = .ckgFlds_chkJournalno.Top
              '.ckgFlds_chkAccountNo_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
5700          .ckgFlds_chkShortName.Top = .ckgFlds_chkJournalno.Top
5710          .ckgFlds_chkShortName_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
5720          .ckgFlds_chkCusip.Top = .ckgFlds_chkJournalType.Top
5730          .ckgFlds_chkCusip_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
5740          .ckgFlds_chkAssetDescription.Top = .ckgFlds_chkTransDate.Top
5750          .ckgFlds_chkAssetDescription_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
5760          .ckgFlds_chkShareFace.Top = .ckgFlds_chkAccountNo.Top
5770          .ckgFlds_chkShareFace_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
5780          .ckgFlds_chkICash.Top = .ckgFlds_chkJournalno.Top
5790          .ckgFlds_chkICash_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
5800          .ckgFlds_chkPCash.Top = .ckgFlds_chkJournalType.Top
5810          .ckgFlds_chkPCash_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
5820          .ckgFlds_chkCost.Top = .ckgFlds_chkTransDate.Top
5830          .ckgFlds_chkCost_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
5840          .ckgFlds_chkCurrID.Top = .ckgFlds_chkAccountNo.Top
5850          .ckgFlds_chkCurrID_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
5860          .ckgFlds_chkAssetDate.Top = .ckgFlds_chkJournalno.Top
5870          .ckgFlds_chkAssetDate_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
5880          .ckgFlds_chkPurchaseDate.Top = .ckgFlds_chkJournalType.Top
5890          .ckgFlds_chkPurchaseDate_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
5900          .ckgFlds_chkLedgerDescription.Top = .ckgFlds_chkTransDate.Top
5910          .ckgFlds_chkLedgerDescription_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
5920          .ckgFlds_chkRecurringItem.Top = .ckgFlds_chkAccountNo.Top
5930          .ckgFlds_chkRecurringItem_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
5940          .ckgFlds_chkRevCodeDesc.Top = .ckgFlds_chkJournalno.Top
5950          .ckgFlds_chkRevCodeDesc_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
5960          .ckgFlds_chkRevCodeTypeDescription.Top = .ckgFlds_chkJournalType.Top
5970          .ckgFlds_chkRevCodeTypeDescription_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
5980          .ckgFlds_chkTaxCodeDescription.Top = .ckgFlds_chkTransDate.Top
5990          .ckgFlds_chkTaxCodeDescription_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
6000          .ckgFlds_chkTaxCodeTypeDescription.Top = .ckgFlds_chkAccountNo.Top
6010          .ckgFlds_chkTaxCodeTypeDescription_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
6020          .ckgFlds_chkLocationName.Top = .ckgFlds_chkJournalno.Top
6030          .ckgFlds_chkLocationName_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
6040          .ckgFlds_chkCheckNum.Top = .ckgFlds_chkJournalType.Top
6050          .ckgFlds_chkCheckNum_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
6060          .ckgFlds_chkJournalUser.Top = .ckgFlds_chkTransDate.Top
6070          .ckgFlds_chkJournalUser_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
6080          .ckgFlds_chkPosted.Top = .ckgFlds_chkAccountNo.Top
6090          .ckgFlds_chkPosted_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
6100          .ckgFlds_chkLedgerHidden.Top = .ckgFlds_chkJournalno.Top
6110          .ckgFlds_chkLedgerHidden_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
6120          .FldCnt.Top = .ckgFlds_chkPosted_lbl.Top
6130          .FldCnt_lbl.Top = .FldCnt.Top
6140          .FocusHolder.Top = .ckgFlds_hline01.Top
6150          .Detail.Height = lngDetail_Height_New
6160        Case False
6170          If lngDetail_Height_New > .Detail.Height Then
6180            .Detail.Height = lngDetail_Height_New
6190          End If
6200  On Error Resume Next
6210          .cmdSelect_box.Top = (lngSelectBox_Top + lngDetail_Height_Diff)
6220          If ERR.Number <> 0 Then
                ' ** Most likely it's closing and shouldn't be here!
                ' ** Something is setting gblClosing = False after cmdClose is clicked!
6230            gblnClosing = True
6240  On Error GoTo ERRH
6250          Else
6260  On Error GoTo ERRH
6270            .cmdSelect_lbl.Top = (lngSelectLbl_Top + lngDetail_Height_Diff)
6280            .cmdSelect_lbl3.Top = (.cmdSelect_box.Top - (2& * lngTpp))
6290            .cmdSelect_lbl3_cmd.Top = .cmdSelect_lbl3.Top
6300            .cmdSelect_Legend_lbl.Top = (lngSelectLegendLbl_Top + lngDetail_Height_Diff)
6310            .cmdSelect_Legend_box.Top = (.cmdSelect_Legend_lbl.Top - lngTpp)
6320            .cmdSelect_lbl2.Top = (lngSelectLbl2_Top + lngDetail_Height_Diff)
6330            .chkNarrowFont_box.Top = .cmdSelect_box.Top
6340            .chkNarrowFont.Top = (lngNarrowFont_Top + lngDetail_Height_Diff)
6350            .chkNarrowFont_lbl.Top = (.chkNarrowFont.Top - lngTpp)
6360            .chkIncludeSharefaceTot.Top = (lngShareface_Top + lngDetail_Height_Diff)
6370            .chkIncludeSharefaceTot_lbl.Top = (.chkIncludeSharefaceTot.Top - lngTpp)
6380            .chkPageOf.Top = .chkNarrowFont.Top
6390            .chkPageOf_lbl.Top = .chkNarrowFont_lbl.Top
6400            .chkSaveSizePos.Top = .chkIncludeSharefaceTot.Top
6410            .chkSaveSizePos_lbl.Top = .chkIncludeSharefaceTot_lbl.Top
6420            .chkNarrowFont_vline01.Top = (.cmdSelect_box.Top + lngTpp)
6430            .chkNarrowFont_vline02.Top = .chkNarrowFont_vline01.Top
6440            .cmdSelect_Legend_Tgl_lbl1.Top = (.cmdSelect_Legend_lbl.Top - lngTpp)
6450            .cmdSelect_Legend_Tgl_lbl2.Top = (.cmdSelect_Legend_Tgl_lbl1.Top + (16& * lngTpp))
6460            .cmdSelect_Legend_tgl_on_raised_img.Top = .chkPageOf.Top
6470            .cmdSelect_Legend_tgl_on_raised_focus_img.Top = .cmdSelect_Legend_tgl_on_raised_img.Top
6480            .cmdSelect_Legend_tgl_on.Top = .cmdSelect_Legend_tgl_on_raised_img.Top
6490            .cmdSelect_Legend_tgl_off_raised_img.Top = .chkSaveSizePos_lbl.Top
6500            .cmdSelect_Legend_tgl_off_raised_focus_img.Top = .cmdSelect_Legend_tgl_off_raised_img.Top
6510            .cmdSelect_Legend_tgl_off.Top = .cmdSelect_Legend_tgl_off_raised_img.Top
6520            .cmdSelect_Legend_box.Top = .cmdSelect_Legend_Tgl_lbl1.Top
6530            .cmdSelect_Legend_vline01.Top = (.cmdSelect_Legend_tgl_on_raised_img.Top + lngTpp)
6540            .cmdSelect_Legend_vline02.Top = .cmdSelect_Legend_vline01.Top
6550            .cmdSelectAll.Top = (lngSelectAll_Top + lngDetail_Height_Diff)
6560            .cmdSelectAll_raised_img.Top = .cmdSelectAll.Top
6570            .cmdSelectAll_raised_semifocus_dots_img.Top = .cmdSelectAll.Top
6580            .cmdSelectAll_raised_focus_img.Top = .cmdSelectAll.Top
6590            .cmdSelectAll_raised_focus_dots_img.Top = .cmdSelectAll.Top
6600            .cmdSelectAll_sunken_focus_dots_img.Top = .cmdSelectAll.Top
6610            .cmdSelectAll_raised_img_dis.Top = .cmdSelectAll.Top
6620            .cmdSelectNone.Top = (lngSelectNone_Top + lngDetail_Height_Diff)
6630            .cmdSelectNone_raised_img.Top = .cmdSelectNone.Top
6640            .cmdSelectNone_raised_semifocus_dots_img.Top = .cmdSelectNone.Top
6650            .cmdSelectNone_raised_focus_img.Top = .cmdSelectNone.Top
6660            .cmdSelectNone_raised_focus_dots_img.Top = .cmdSelectNone.Top
6670            .cmdSelectNone_sunken_focus_dots_img.Top = .cmdSelectNone.Top
6680            .cmdSelectNone_raised_img_dis.Top = .cmdSelectNone.Top
6690            .frmTransaction_Audit_Sub.Height = (lngSub_Height + lngDetail_Height_Diff)
6700            .frmTransaction_Audit_Sub_ds.Height = (lngSub_Height + lngDetail_Height_Diff)
6710            .frmTransaction_Audit_Sub_box.Height = (.frmTransaction_Audit_Sub.Height + (6& * lngTpp))
6720            .Nav_hline01.Top = (.frmTransaction_Audit_Sub.Top + .frmTransaction_Audit_Sub.Height)
6730            .Nav_hline02.Top = (.Nav_hline01.Top + lngTpp)
6740            .Nav_hline03.Top = .Nav_hline01.Top
6750            .Nav_vline01.Top = .Nav_hline01.Top
6760            .Nav_vline02.Top = .Nav_hline01.Top
6770            .Nav_vline03.Top = .Nav_hline01.Top
6780            .Nav_vline04.Top = .Nav_hline01.Top
6790            .Nav_box01.Top = (.frmTransaction_Audit_Sub.Top + .frmTransaction_Audit_Sub.Height)
6800            .ShortcutMenu_lbl.Top = ((.frmTransaction_Audit_Sub.Top + .frmTransaction_Audit_Sub.Height) + (6& * lngTpp))
6810            .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
6820            .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
6830            .Detail_hline01.Top = (.ShortcutMenu_lbl.Top + .ShortcutMenu_lbl.Height)
6840            .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
6850            .Detail_vline01.Top = .Detail_hline01.Top
6860            .Detail_vline02.Top = .Detail_hline01.Top
6870            lngTmp01 = ((.cmdSelect_box.Top + .cmdSelect_box.Height) + (8& * lngTpp))
6880            .Detail_hline03.Top = lngTmp01
6890            .Detail_hline04.Top = (lngTmp01 + lngTpp)
6900            .Detail_vline03.Top = .Detail_hline03.Top
6910            .Detail_vline04.Top = .Detail_hline03.Top
                ' ** Differences are from the control's original position, not its current position.
6920            .ckgFlds_hline01.Top = (.Detail_hline04.Top + (8& * lngTpp))
6930            .ckgFlds_hline02.Top = (.ckgFlds_hline01.Top + lngTpp)
6940            .ckgFlds_vline01.Top = .ckgFlds_hline02.Top
6950            .ckgFlds_vline02.Top = (.ckgFlds_vline01.Top + lngTpp)
6960            .ckgFlds_vline03.Top = .ckgFlds_vline02.Top
6970            .ckgFlds_vline04.Top = .ckgFlds_vline01.Top
6980            .ckgFlds_lbl.Top = .ckgFlds_hline02.Top
6990            .ckgFlds_box2.Top = (.ckgFlds_lbl.Top + lngTpp)
7000            .ckgFlds_cmd.Top = (.ckgFlds_box2.Top + lngTpp)
7010            .ckgFlds_cmd_raised_img.Top = .ckgFlds_cmd.Top
7020            .ckgFlds_cmd_raised_semifocus_dots_img.Top = .ckgFlds_cmd.Top
7030            .ckgFlds_cmd_raised_focus_img.Top = .ckgFlds_cmd.Top
7040            .ckgFlds_cmd_raised_focus_dots_img.Top = .ckgFlds_cmd.Top
7050            .ckgFlds_cmd_sunken_focus_dots_img.Top = .ckgFlds_cmd.Top
7060            .ckgFlds_cmd_raised_img_dis.Top = .ckgFlds_cmd.Top
7070            .ckgFlds_box.Top = (.ckgFlds_vline01.Top + .ckgFlds_vline01.Height)
7080            .ckgFlds_hline05.Top = .ckgFlds_box.Top
7090            lngTmp01 = ((.ckgFlds_cmd.Top + .ckgFlds_cmd.Height) + lngTpp)
7100            .ckgFlds_hline03.Top = lngTmp01
7110            .ckgFlds_hline04.Top = (.ckgFlds_hline03.Top + lngTpp)
7120            .ckgFlds_vline05.Top = (.ckgFlds_vline01.Top + .ckgFlds_vline01.Height)
7130            .ckgFlds_vline06.Top = (.ckgFlds_vline05.Top + lngTpp)
7140            .ckgFlds_vline07.Top = .ckgFlds_vline06.Top
7150            .ckgFlds_vline08.Top = .ckgFlds_vline05.Top
7160            Select Case .ckgFlds
                Case True
7170              .ckgFlds_vline09.Top = (lngCkgFldsVLine_Top + lngDetail_Height_Diff)
7180              For lngX = 0& To (lngFrmFlds - 1&)
7190                If arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalno" Then
7200                  .ckgFlds_chkJournalno.Top = (arr_varFrmFld(FM_TOPO, lngX) + lngDetail_Height_Diff)
7210                  .ckgFlds_chkJournalno_lbl.Top = (.ckgFlds_chkJournalno.Top - lngChkBoxLbl_TopOffset)
7220                ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalType" Then
7230                  .ckgFlds_chkJournalType.Top = (arr_varFrmFld(FM_TOPO, lngX) + lngDetail_Height_Diff)
7240                  .ckgFlds_chkJournalType_lbl.Top = (.ckgFlds_chkJournalType.Top - lngChkBoxLbl_TopOffset)
7250                ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkTransDate" Then
7260                  .ckgFlds_chkTransDate.Top = (arr_varFrmFld(FM_TOPO, lngX) + lngDetail_Height_Diff)
7270                  .ckgFlds_chkTransDate_lbl.Top = (.ckgFlds_chkTransDate.Top - lngChkBoxLbl_TopOffset)
7280                ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkAccountNo" Then
7290                  .ckgFlds_chkAccountNo.Top = (arr_varFrmFld(FM_TOPO, lngX) + lngDetail_Height_Diff)
7300                  .ckgFlds_chkAccountNo_lbl.Top = (.ckgFlds_chkAccountNo.Top - lngChkBoxLbl_TopOffset)
7310                  Exit For
7320                End If
7330              Next
7340            Case False
                  'lngTmp01 = (((lngSelectBox_Top + .cmdSelect_box.Height) + (8& * lngTpp)) + lngTpp)  ' ** Original Detail_hline04.Top.
                  'lngTmp01 = ((((lngTmp01 + (8& * lngTpp)) + lngTpp) + lngTpp) + lngTpp)  ' ** Original ckgFlds_cmd.Top.
                  'lngTmp01 = (.ckgFlds_cmd.Top - lngTmp01)  ' ** Difference in .ckgFlds_cmd.Top position.
7350              .ckgFlds_vline09.Top = .ckgFlds_cmd.Top
7360              For lngX = 0& To (lngFrmFlds - 1&)
7370                If arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalno" Then
7380                  .ckgFlds_chkJournalno.Top = (arr_varFrmFld(FM_TOPC, lngX) - lngDetail_Height_Diff)
7390                  .ckgFlds_chkJournalno_lbl.Top = (.ckgFlds_chkJournalno.Top - lngChkBoxLbl_TopOffset)
7400                ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkJournalType" Then
7410                  .ckgFlds_chkJournalType.Top = (arr_varFrmFld(FM_TOPC, lngX) - lngDetail_Height_Diff)
7420                  .ckgFlds_chkJournalType_lbl.Top = .ckgFlds_chkJournalType.Top - lngChkBoxLbl_TopOffset
7430                ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkTransDate" Then
7440                  .ckgFlds_chkTransDate.Top = (arr_varFrmFld(FM_TOPC, lngX) - lngDetail_Height_Diff)
7450                  .ckgFlds_chkTransDate_lbl.Top = .ckgFlds_chkTransDate.Top - lngChkBoxLbl_TopOffset
7460                ElseIf arr_varFrmFld(FM_VIEWCHK, lngX) = "ckgFlds_chkAccountNo" Then
7470                  .ckgFlds_chkAccountNo.Top = (arr_varFrmFld(FM_TOPC, lngX) - lngDetail_Height_Diff)
7480                  .ckgFlds_chkAccountNo_lbl.Top = (.ckgFlds_chkAccountNo.Top - lngChkBoxLbl_TopOffset)
7490                  Exit For
7500                End If
7510              Next
7520            End Select
7530            .ckgFlds_vline10.Top = .ckgFlds_vline09.Top
7540            .ckgFlds_vline11.Top = .ckgFlds_vline09.Top
7550            .ckgFlds_vline12.Top = .ckgFlds_vline09.Top
7560            .ckgFlds_vline13.Top = .ckgFlds_vline09.Top
7570            .ckgFlds_vline14.Top = .ckgFlds_vline09.Top
7580            .ckgFlds_vline15.Top = .ckgFlds_vline09.Top
7590            .ckgFlds_vline16.Top = .ckgFlds_vline09.Top
7600            .ckgFlds_vline17.Top = .ckgFlds_vline09.Top
7610            .ckgFlds_vline18.Top = .ckgFlds_vline09.Top
7620            .ckgFlds_vline19.Top = .ckgFlds_vline09.Top
7630            .ckgFlds_vline20.Top = .ckgFlds_vline09.Top
                '.ckgFlds_vline21.Top = .ckgFlds_vline09.Top
                '.ckgFlds_vline22.Top = .ckgFlds_vline09.Top
                '.ckgFlds_chkAccountNo.Top = .ckgFlds_chkJournalno.Top
                '.ckgFlds_chkAccountNo_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
7640            .ckgFlds_chkShortName.Top = .ckgFlds_chkJournalno.Top
7650            .ckgFlds_chkShortName_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
7660            .ckgFlds_chkCusip.Top = .ckgFlds_chkJournalType.Top
7670            .ckgFlds_chkCusip_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
7680            .ckgFlds_chkAssetDescription.Top = .ckgFlds_chkTransDate.Top
7690            .ckgFlds_chkAssetDescription_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
7700            .ckgFlds_chkShareFace.Top = .ckgFlds_chkAccountNo.Top
7710            .ckgFlds_chkShareFace_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
7720            .ckgFlds_chkICash.Top = .ckgFlds_chkJournalno.Top
7730            .ckgFlds_chkICash_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
7740            .ckgFlds_chkPCash.Top = .ckgFlds_chkJournalType.Top
7750            .ckgFlds_chkPCash_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
7760            .ckgFlds_chkCost.Top = .ckgFlds_chkTransDate.Top
7770            .ckgFlds_chkCost_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
7780            .ckgFlds_chkCurrID.Top = .ckgFlds_chkAccountNo.Top
7790            .ckgFlds_chkCurrID_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
7800            .ckgFlds_chkAssetDate.Top = .ckgFlds_chkJournalno.Top
7810            .ckgFlds_chkAssetDate_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
7820            .ckgFlds_chkPurchaseDate.Top = .ckgFlds_chkJournalType.Top
7830            .ckgFlds_chkPurchaseDate_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
7840            .ckgFlds_chkLedgerDescription.Top = .ckgFlds_chkTransDate.Top
7850            .ckgFlds_chkLedgerDescription_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
7860            .ckgFlds_chkRecurringItem.Top = .ckgFlds_chkAccountNo.Top
7870            .ckgFlds_chkRecurringItem_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
7880            .ckgFlds_chkRevCodeDesc.Top = .ckgFlds_chkJournalno.Top
7890            .ckgFlds_chkRevCodeDesc_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
7900            .ckgFlds_chkRevCodeTypeDescription.Top = .ckgFlds_chkJournalType.Top
7910            .ckgFlds_chkRevCodeTypeDescription_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
7920            .ckgFlds_chkTaxCodeDescription.Top = .ckgFlds_chkTransDate.Top
7930            .ckgFlds_chkTaxCodeDescription_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
7940            .ckgFlds_chkTaxCodeTypeDescription.Top = .ckgFlds_chkAccountNo.Top
7950            .ckgFlds_chkTaxCodeTypeDescription_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
7960            .ckgFlds_chkLocationName.Top = .ckgFlds_chkJournalno.Top
7970            .ckgFlds_chkLocationName_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
7980            .ckgFlds_chkCheckNum.Top = .ckgFlds_chkJournalType.Top
7990            .ckgFlds_chkCheckNum_lbl.Top = .ckgFlds_chkJournalType_lbl.Top
8000            .ckgFlds_chkJournalUser.Top = .ckgFlds_chkTransDate.Top
8010            .ckgFlds_chkJournalUser_lbl.Top = .ckgFlds_chkTransDate_lbl.Top
8020            .ckgFlds_chkPosted.Top = .ckgFlds_chkAccountNo.Top
8030            .ckgFlds_chkPosted_lbl.Top = .ckgFlds_chkAccountNo_lbl.Top
8040            .ckgFlds_chkLedgerHidden.Top = .ckgFlds_chkJournalno.Top
8050            .ckgFlds_chkLedgerHidden_lbl.Top = .ckgFlds_chkJournalno_lbl.Top
8060            .FldCnt.Top = .ckgFlds_chkPosted_lbl.Top
8070            .FldCnt_lbl.Top = .FldCnt.Top
8080            .FocusHolder.Top = .ckgFlds_hline01.Top
8090            .Detail.Height = lngDetail_Height_New
8100          End If  ' ** Err.Number.
8110        End Select  ' ** blnShorter.

8120      End With
8130    End If  ' ** gblnClosing.

EXITP:
8140    Exit Sub

ERRH:
8150    Select Case ERR.Number
        Case Else
8160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8170    End Select
8180    Resume EXITP

End Sub

Public Sub PrintTgls_Load(frm As Access.Form)
' ** Called by:
' **   frmTransaction_Audit_Sub:
' **     Form_Open()
' **   modTransactionAuditFuncs1, this:
' **     PrintTgls_Set(), below

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "PrintTgls_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, ctl As Access.Control
        Dim lngItems As Long, arr_varItem As Variant
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim lngW As Long, lngE As Long

        ' ** Array: arr_varItem().
        'Const I_DID As Integer = 0
        'Const I_DNAM As Integer = 1
        'Const I_FID  As Integer = 2
        'Const I_FNAM As Integer = 3
        'Const I_CID  As Integer = 4
        Const I_CNAM As Integer = 5
        'Const I_TYP  As Integer = 6
        'Const I_TAB  As Integer = 7
        'Const I_LFT  As Integer = 8
        'Const I_WDT  As Integer = 9
        'Const I_VIS  As Integer = 10

8210    With frm

8220      If lngChks = 0& Or IsEmpty(arr_varChk) = True Then

8230        lngChks = 0&
8240        ReDim arr_varChk(C_ELEMS, 0)

8250        Set dbs = CurrentDb
8260        With dbs
              ' ** tblForm_Control, just 'frmTransaction_Audit_Sub', sorted.
8270          Set qdf = .QueryDefs("qryTransaction_Audit_06")
8280          Set rst = qdf.OpenRecordset
8290          With rst
8300            .MoveLast
8310            lngItems = .RecordCount
8320            .MoveFirst
8330            arr_varItem = .GetRows(lngItems)
                ' *****************************************************
                ' ** Array: arr_varItem()
                ' **
                ' **   Field  Element  Name                Constant
                ' **   =====  =======  ==================  ==========
                ' **     1       0     dbs_id              I_DID
                ' **     2       1     dbs_name            I_DNAM
                ' **     3       2     frm_id              I_FID
                ' **     4       3     frm_name            I_FNAM
                ' **     5       4     ctl_id              I_CID
                ' **     6       5     ctl_name            I_CNAM
                ' **     7       6     ctltype_type        I_TYP
                ' **     8       7     ctlspec_tabindex    I_TAB
                ' **     9       8     ctlspec_left        I_LFT
                ' **    10       9     ctlspec_width       I_WDT
                ' **    11      10     ctlspec_visible     I_VIS
                ' **
                ' *****************************************************
8340            .Close
8350          End With
8360          Set rst = Nothing
8370          Set qdf = Nothing
8380          .Close
8390        End With
8400        Set dbs = Nothing

8410        For lngW = 0& To (lngItems - 1&)
              ' ** 24 Fields.
8420          For Each ctl In .FormHeader.Controls
8430            With ctl
8440              If Right(.Name, 4) = "_chk" Then
8450                If .Visible = False Then
8460                  intPos01 = InStr(.Name, "_chk")
8470                  strTmp01 = Left(.Name, (intPos01 - 1&))
8480                  If strTmp01 = arr_varItem(I_CNAM, lngW) Then
8490                    lngChks = lngChks + 1&
8500                    lngE = lngChks - 1&
8510                    ReDim Preserve arr_varChk(C_ELEMS, lngE)
8520                    arr_varChk(C_FNAM, lngE) = strTmp01
8530                    arr_varChk(C_CHKBX, lngE) = .Name
8540                    arr_varChk(C_INCL, lngE) = .Value
8550                    arr_varChk(C_FOC, lngE) = CBool(False)
8560                    arr_varChk(C_MOUS, lngE) = CBool(False)
8570                    arr_varChk(C_DIS, lngE) = CBool(False)
8580                    strTmp02 = strTmp01 & "_tgl"
8590                    arr_varChk(C_CMD, lngE) = strTmp02
8600                    arr_varChk(C_OFR, lngE) = strTmp02 & "_off_raised_img"
8610                    arr_varChk(C_OFRD, lngE) = strTmp02 & "_off_raised_dots_img"
8620                    arr_varChk(C_OFRF, lngE) = strTmp02 & "_off_raised_focus_img"
8630                    arr_varChk(C_OFRFD, lngE) = strTmp02 & "_off_raised_focus_dots_img"
8640                    arr_varChk(C_OFDIS, lngE) = strTmp02 & "_off_raised_img_dis"
8650                    arr_varChk(C_ONR, lngE) = strTmp02 & "_on_raised_img"
8660                    arr_varChk(C_ONRD, lngE) = strTmp02 & "_on_raised_dots_img"
8670                    arr_varChk(C_ONRF, lngE) = strTmp02 & "_on_raised_focus_img"
8680                    arr_varChk(C_ONRFD, lngE) = strTmp02 & "_on_raised_focus_dots_img"
8690                    arr_varChk(C_ONSD, lngE) = strTmp02 & "_on_sunken_dots_img"
8700                    arr_varChk(C_ONDIS, lngE) = strTmp02 & "_on_raised_img_dis"
                        ' *********************************************************************
                        ' ** Array: arr_varChk()
                        ' **
                        ' **   Field  Element  Name                                Constant
                        ' **   =====  =======  ==================================  ==========
                        ' **     1       0     Field Name                          C_FNAM
                        ' **     2       1     Check Box                           C_CHKBX
                        ' **     3       2     Include T/F                         C_INCL
                        ' **     4       3     Focus T/F                           C_FOC
                        ' **     5       4     MouseDown T/F                       C_MOUS
                        ' **     6       5     Disabled T/F                        C_DIS
                        ' **     7       6     Command Button                      C_CMD
                        ' **     8       7     .._tgl_off_raised_img               C_OFR
                        ' **     9       8     .._tgl_off_raised_dots_img          C_OFRD
                        ' **    10       9     .._tgl_off_raised_focus_img         C_OFRF
                        ' **    11      10     .._tgl_off_raised_focus_dots_img    C_OFRFD
                        ' **    12      11     .._tgl_off_raised_img_dis           C_OFDIS
                        ' **    13      12     .._tgl_on_raised_img                C_ONR
                        ' **    14      13     .._tgl_on_raised_dots_img           C_ONRD
                        ' **    15      14     .._tgl_on_raised_focus_img          C_ONRF
                        ' **    16      15     .._tgl_on_raised_focus_dots_img     C_ONRFD
                        ' **    17      16     .._tgl_on_sunken_dots_img           C_ONSD
                        ' **    18      17     .._tgl_on_raised_img_dis            C_ONDIS
                        ' **
                        ' *********************************************************************
8710                  End If
8720                End If
8730              End If
8740            End With
8750          Next  ' ** ctl.
8760        Next  ' ** lngW.

8770        .ChkArray_Set arr_varChk  ' ** Form Procedure: frmTransaction_Audit_Sub.
8780        DoEvents

            'For lngX = 0& To (lngChks - 1&)
            '  Debug.Print "'" & arr_varChk(C_FNAM, lngX)
            'Next

            'journalno
            'journaltype
            'transdate
            'accountno
            'shortname
            'cusip
            'asset_description
            'shareface
            'icash
            'pcash
            'cost
            'curr_id
            'assetdate
            'PurchaseDate
            'ledger_description
            'RecurringItem
            'revcode_DESC
            'revcode_TYPE_Description
            'taxcode_description
            'taxcode_type_description
            'Location_Name
            'CheckNum
            'journal_USER
            'posted
            'ledger_HIDDEN

8790      End If  ' ** lngChks.

8800    End With

EXITP:
8810    Set ctl = Nothing
8820    Set rst = Nothing
8830    Set qdf = Nothing
8840    Set dbs = Nothing
8850    Exit Sub

ERRH:
8860    Select Case ERR.Number
        Case Else
8870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8880    End Select
8890    Resume EXITP

End Sub

Public Function ChkArray_Get() As Variant

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkArray_Get"

        Dim arr_varRetVal As Variant

8910    arr_varRetVal = arr_varChk

EXITP:
8920    ChkArray_Get = arr_varRetVal
8930    Exit Function

ERRH:
8940    Select Case ERR.Number
        Case Else
8950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8960    End Select
8970    Resume EXITP

End Function

Public Sub PrintTgls_Click(strProc As String, blnFromTgl As Boolean, lngChks As Long, arr_varChk As Variant, frm As Access.Form)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "PrintTgls_Click"

        Dim strFldName As String
        Dim lngX As Long

9010    With frm
9020      strFldName = Left(strProc, (InStr(strProc, "_tgl") - 1))
9030      For lngX = 0& To (lngChks - 1&)
9040        If arr_varChk(C_FNAM, lngX) = strFldName Then
              ' ** Here's where it should be turning on and off the checkbox that preferences uses.
              ' ** Aha!  Because this is called twice (because it seems to
              ' ** be alternating On/Off), what was off gets turned back on!
9050          .Controls(arr_varChk(C_CHKBX, lngX)) = (Not .Controls(arr_varChk(C_CHKBX, lngX)))
9060          arr_varChk(C_INCL, lngX) = .Controls(arr_varChk(C_CHKBX, lngX))
9070          Select Case arr_varChk(C_INCL, lngX)
              Case True
9080            .Controls(arr_varChk(C_ONRFD, lngX)).Visible = True
9090            .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
9100          Case False
9110            .Controls(arr_varChk(C_OFRFD, lngX)).Visible = True
9120            .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
9130          End Select
9140          .Controls(arr_varChk(C_OFR, lngX)).Visible = False
9150          .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
9160          .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
9170          .Controls(arr_varChk(C_OFDIS, lngX)).Visible = False
9180          .Controls(arr_varChk(C_ONR, lngX)).Visible = False
9190          .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
9200          .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
9210          .Controls(arr_varChk(C_ONSD, lngX)).Visible = False
9220          .Controls(arr_varChk(C_ONDIS, lngX)).Visible = False
9230          Exit For
9240        End If
9250      Next
9260      .Print_Chk strProc  ' ** Form Procedure: frmTransaction_Audit_Sub.
9270      If .Controls(strFldName).Enabled = True Then
9280        blnFromTgl = True
9290  On Error Resume Next
9300        .Controls(strFldName).SetFocus
9310  On Error GoTo ERRH
9320      End If
9330    End With

EXITP:
9340    Exit Sub

ERRH:
9350    Select Case ERR.Number
        Case Else
9360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9370    End Select
9380    Resume EXITP

End Sub

Public Sub PrintTgls_Focus(strProc As String, lngChks As Long, arr_varChk As Variant, frm As Access.Form)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "PrintTgls_Focus"

        Dim strFldName As String
        Dim blnGotFocus As Boolean
        Dim lngX As Long

9410    With frm
9420      strFldName = Left(strProc, (InStr(strProc, "_tgl") - 1))
9430      If InStr(strProc, "GotFocus") > 0 Then
9440        blnGotFocus = True
9450      Else
9460        blnGotFocus = False
9470      End If
9480      Select Case blnGotFocus
          Case True
9490        For lngX = 0& To (lngChks - 1&)
9500          If arr_varChk(C_FNAM, lngX) = strFldName Then
9510            arr_varChk(C_FOC, lngX) = CBool(True)
9520            Select Case arr_varChk(C_INCL, lngX)
                Case True
9530              .Controls(arr_varChk(C_ONRD, lngX)).Visible = True
9540              .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
9550            Case False
9560              .Controls(arr_varChk(C_OFRD, lngX)).Visible = True
9570              .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
9580            End Select
9590            .Controls(arr_varChk(C_OFR, lngX)).Visible = False
9600            .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
9610            .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
9620            .Controls(arr_varChk(C_OFDIS, lngX)).Visible = False
9630            .Controls(arr_varChk(C_ONR, lngX)).Visible = False
9640            .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
9650            .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
9660            .Controls(arr_varChk(C_ONSD, lngX)).Visible = False
9670            .Controls(arr_varChk(C_ONDIS, lngX)).Visible = False
9680            Exit For
9690          End If
9700        Next
9710      Case False
9720        For lngX = 0& To (lngChks - 1&)
9730          If arr_varChk(C_FNAM, lngX) = strFldName Then
9740            Select Case arr_varChk(C_INCL, lngX)
                Case True
9750              .Controls(arr_varChk(C_ONR, lngX)).Visible = True
9760              .Controls(arr_varChk(C_OFR, lngX)).Visible = False
9770            Case False
9780              .Controls(arr_varChk(C_OFR, lngX)).Visible = True
9790              .Controls(arr_varChk(C_ONR, lngX)).Visible = False
9800            End Select
9810            .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
9820            .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
9830            .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
9840            .Controls(arr_varChk(C_OFDIS, lngX)).Visible = False
9850            .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
9860            .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
9870            .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
9880            .Controls(arr_varChk(C_ONSD, lngX)).Visible = False
9890            .Controls(arr_varChk(C_ONDIS, lngX)).Visible = False
9900            arr_varChk(C_FOC, lngX) = CBool(False)
9910            Exit For
9920          End If
9930        Next
9940      End Select
9950    End With

EXITP:
9960    Exit Sub

ERRH:
9970    Select Case ERR.Number
        Case Else
9980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9990    End Select
10000   Resume EXITP

End Sub

Public Function PrintTgls_GetArr() As Variant

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintTgls_GetArr"

        Dim arr_varRetVal As Variant

10110   arr_varRetVal = arr_varChk

EXITP:
10120   PrintTgls_GetArr = arr_varRetVal
10130   Exit Function

ERRH:
10140   Select Case ERR.Number
        Case Else
10150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10160   End Select
10170   Resume EXITP

End Function

Public Sub PrintTgls_Mouse(strProc As String, lngChks As Long, arr_varChk As Variant, frm As Access.Form)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintTgls_Mouse"

        Dim strFldName As String
        Dim blnMouseDown As Boolean
        Dim lngX As Long

10210   With frm
10220     strFldName = Left(strProc, (InStr(strProc, "_tgl") - 1))
10230     If InStr(strProc, "MouseDown") > 0 Then
10240       blnMouseDown = True
10250     Else
10260       blnMouseDown = False
10270     End If
10280     Select Case blnMouseDown
          Case True
10290       For lngX = 0& To (lngChks - 1&)
10300         If arr_varChk(C_FNAM, lngX) = strFldName Then
10310           arr_varChk(C_MOUS, lngX) = CBool(True)
10320           .Controls(arr_varChk(C_ONSD, lngX)).Visible = True  ' ** Same for both On and Off.
10330           .Controls(arr_varChk(C_OFR, lngX)).Visible = False
10340           .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
10350           .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
10360           .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
10370           .Controls(arr_varChk(C_OFDIS, lngX)).Visible = False
10380           .Controls(arr_varChk(C_ONR, lngX)).Visible = False
10390           .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
10400           .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
10410           .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
10420           .Controls(arr_varChk(C_ONDIS, lngX)).Visible = False
10430           Exit For
10440         End If
10450       Next
10460     Case False
10470       For lngX = 0& To (lngChks - 1&)
10480         If arr_varChk(C_FNAM, lngX) = strFldName Then
10490           Select Case arr_varChk(C_INCL, lngX)
                Case True
10500             .Controls(arr_varChk(C_ONRFD, lngX)).Visible = True
10510             .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
10520           Case False
10530             .Controls(arr_varChk(C_OFRFD, lngX)).Visible = True
10540             .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
10550           End Select
10560           .Controls(arr_varChk(C_OFR, lngX)).Visible = False
10570           .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
10580           .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
10590           .Controls(arr_varChk(C_OFDIS, lngX)).Visible = False
10600           .Controls(arr_varChk(C_ONR, lngX)).Visible = False
10610           .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
10620           .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
10630           .Controls(arr_varChk(C_ONSD, lngX)).Visible = False
10640           .Controls(arr_varChk(C_ONDIS, lngX)).Visible = False
10650           arr_varChk(C_MOUS, lngX) = CBool(False)
10660           Exit For
10670         End If
10680       Next
10690     End Select
10700   End With

EXITP:
10710   Exit Sub

ERRH:
10720   Select Case ERR.Number
        Case Else
10730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10740   End Select
10750   Resume EXITP

End Sub

Public Sub PrintTgls_Move(strProc As String, lngChks As Long, arr_varChk As Variant, frm As Access.Form)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintTgls_Move"

        Dim strFldName As String
        Dim lngX As Long

10810   With frm
10820     strFldName = Left(strProc, (InStr(strProc, "_tgl") - 1))
10830     For lngX = 0& To (lngChks - 1&)
10840       If arr_varChk(C_FNAM, lngX) = strFldName Then
10850         If arr_varChk(C_MOUS, lngX) = False Then
10860           Select Case arr_varChk(C_INCL, lngX)
                Case True
10870             Select Case arr_varChk(C_FOC, lngX)
                  Case True
10880               .Controls(arr_varChk(C_ONRFD, lngX)).Visible = True
10890               .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
10900             Case False
10910               .Controls(arr_varChk(C_ONRF, lngX)).Visible = True
10920               .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
10930             End Select
10940             .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
10950             .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
10960           Case False
10970             Select Case arr_varChk(C_FOC, lngX)
                  Case True
10980               .Controls(arr_varChk(C_OFRFD, lngX)).Visible = True
10990               .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
11000             Case False
11010               .Controls(arr_varChk(C_OFRF, lngX)).Visible = True
11020               .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
11030             End Select
11040             .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
11050             .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
11060           End Select
11070           .Controls(arr_varChk(C_OFR, lngX)).Visible = False
11080           .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
11090           .Controls(arr_varChk(C_OFDIS, lngX)).Visible = False
11100           .Controls(arr_varChk(C_ONR, lngX)).Visible = False
11110           .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
11120           .Controls(arr_varChk(C_ONSD, lngX)).Visible = False
11130           .Controls(arr_varChk(C_ONDIS, lngX)).Visible = False
11140         End If
11150         Exit For
11160       End If
11170     Next
11180   End With

EXITP:
11190   Exit Sub

ERRH:
11200   Select Case ERR.Number
        Case Else
11210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11220   End Select
11230   Resume EXITP

End Sub

Public Sub PrintTgls_Set(frm As Access.Form)
' ** Called by:
' **   frmTransaction_Audit:
' **     Form_Load()
' **     cmdSelectAll_Click()
' **     cmdSelectNone_Click()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintTgls_Set"

        Dim frm1 As Access.Form
        Dim lngFlds As Long, arr_varFld As Variant
        Dim blnShow As Boolean
        Dim lngX As Long, lngY As Long

11310   Set frm1 = Forms("frmTransaction_Audit").frmTransaction_Audit_Sub.Form
11320   With frm1

11330     blnFromSet = False
11340     If lngChks = 0& Or IsEmpty(arr_varChk) Then
11350       blnFromSet = True
11360       PrintTgls_Load frm1 ' ** Procedure: Above.
11370       blnFromSet = False
11380     End If

11390     lngFlds = 0&
11400     arr_varFld = .FldArray_Get  ' ** Form Function: frmTransaction_Audit_Sub.
11410     lngFlds = UBound(arr_varFld, 2) + 1&

11420     For lngX = 0& To (lngChks - 1&)
11430       arr_varChk(C_INCL, lngX) = .Controls(arr_varChk(C_CHKBX, lngX))
11440       blnShow = True
11450       For lngY = 0& To (lngFlds - 1&)
11460         If arr_varFld(F_FNAM, lngY) = arr_varChk(C_FNAM, lngX) Then
11470           blnShow = .Parent.Controls(arr_varFld(F_CNAM, lngY))
11480           Exit For
11490         End If
11500       Next  ' ** lngY.
11510       If blnShow = True Then
11520         Select Case arr_varChk(C_INCL, lngX)
              Case True
11530           .Controls(arr_varChk(C_ONR, lngX)).Visible = True
11540           .Controls(arr_varChk(C_OFR, lngX)).Visible = False
11550         Case False
11560           .Controls(arr_varChk(C_OFR, lngX)).Visible = True
11570           .Controls(arr_varChk(C_ONR, lngX)).Visible = False
11580         End Select
11590         .Controls(arr_varChk(C_OFRD, lngX)).Visible = False
11600         .Controls(arr_varChk(C_OFRF, lngX)).Visible = False
11610         .Controls(arr_varChk(C_OFRFD, lngX)).Visible = False
11620         .Controls(arr_varChk(C_OFDIS, lngX)).Visible = False
11630         .Controls(arr_varChk(C_ONRD, lngX)).Visible = False
11640         .Controls(arr_varChk(C_ONRF, lngX)).Visible = False
11650         .Controls(arr_varChk(C_ONRFD, lngX)).Visible = False
11660         .Controls(arr_varChk(C_ONSD, lngX)).Visible = False
11670         .Controls(arr_varChk(C_ONDIS, lngX)).Visible = False
11680       End If  ' ** blnShow.
11690     Next  ' ** lngX.

11700   End With

EXITP:
11710   Set frm1 = Nothing
11720   Exit Sub

ERRH:
11730   Select Case ERR.Number
        Case Else
11740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11750   End Select
11760   Resume EXITP

End Sub

Public Sub ProgBar_Width_Trans(frm As Access.Form, dblWidth As Double, intMode As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "ProgBar_Width_Trans"

        Dim strCtlName As String, blnVis As Boolean
        Dim lngX As Long

11810   With frm
11820     Select Case intMode
          Case 1
11830       blnVis = CBool(dblWidth)
11840       For lngX = 1& To 6&
11850         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
11860         .Controls(strCtlName).Visible = blnVis
11870       Next
11880     Case 2
11890       For lngX = 1& To 6&
11900         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
11910         .Controls(strCtlName).Width = dblWidth
11920       Next
11930     End Select
11940   End With

EXITP:
11950   Exit Sub

ERRH:
11960   Select Case ERR.Number
        Case Else
11970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11980   End Select
11990   Resume EXITP

End Sub

Public Function ChecksChecked(frm As Access.Form) As Long

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "ChecksChecked"

        Dim ctl As Access.Control
        Dim strTmp01 As String, strTmp02 As String
        Dim lngRetVal As Long

12010   With frm
12020     lngRetVal = 0&
12030     For Each ctl In .Detail.Controls
12040       With ctl
12050         If Left(.Name, 11) = "ckgFlds_chk" And .ControlType = acCheckBox Then
12060           If .Value = True Then
12070             lngRetVal = lngRetVal + 1&
12080           End If
12090         End If
12100       End With
12110     Next
12120     .FldCnt = lngRetVal  ' ** Doesn't matter whether it's visible or not.
12130     strTmp01 = .cmdSelect_lbl2.Caption
12140     strTmp02 = Mid(strTmp01, InStr(strTmp01, "Select"))
12150     strTmp01 = Left(strTmp01, (InStr(strTmp01, "of") + 1))
12160     strTmp01 = strTmp01 & " " & CStr(lngRetVal) & " " & strTmp02
12170     strTmp02 = Trim(Left(strTmp01, InStr(strTmp01, " ")))
12180     If Val(strTmp02) > lngRetVal Then
12190       strTmp02 = CStr(lngRetVal)
12200       strTmp01 = strTmp02 & Mid(strTmp01, InStr(strTmp01, " "))
12210     End If
12220     .cmdSelect_lbl2.Caption = strTmp01
12230     DoEvents
12240   End With

        ' ** Form opening generated this many hits!
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 3  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 2
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 2  Print_ChkCnt_TA()  modTransactionAuditFuncs1  TOT: 22
        'HERE! 1  ChecksChecked()  modTransactionAuditFuncs1  22 of 22 Selected

EXITP:
12250   Set ctl = Nothing
12260   ChecksChecked = lngRetVal
12270   Exit Function

ERRH:
12280   lngRetVal = 0&
12290   Select Case ERR.Number
        Case Else
12300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12310   End Select
12320   Resume EXITP

End Function

Public Function VisibleCounts_Get(intMode As Long) As Variant
' ** intMode is opgView.
' ** Used by:
' **   rptTransaction_Audit_01:
' **     Report_Open()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "VisibleCounts_Get"

        Dim frm As Access.Form
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim arr_lngRetVal(1) As Long

12410   Set frm = Forms("frmTransaction_Audit")
12420   strTmp01 = frm.VisCnts_Get  ' ** Form Module: frmTransaction_Audit.
12430   intPos01 = InStr(strTmp01, "~")
12440   lngChkCnt = Val(Left(strTmp01, (intPos01 - 1)))
12450   strTmp01 = Mid(strTmp01, (intPos01 + 1))
12460   intPos01 = InStr(strTmp01, "~")
12470   lngVisCnt1 = Val(Left(strTmp01, (intPos01 - 1)))
12480   lngVisCnt2 = Val(Mid(strTmp01, (intPos01 + 1)))

12490   Select Case intMode
        Case 1
12500     arr_lngRetVal(0) = lngChkCnt
12510     arr_lngRetVal(1) = lngVisCnt1
12520   Case 2
12530     arr_lngRetVal(0) = lngChkCnt
12540     arr_lngRetVal(1) = lngVisCnt2
12550   End Select

EXITP:
12560   Set frm = Nothing
12570   VisibleCounts_Get = arr_lngRetVal
12580   Exit Function

ERRH:
12590   arr_lngRetVal(0) = -9&
12600   Select Case ERR.Number
        Case Else
12610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12620   End Select
12630   Resume EXITP

End Function

Public Sub CriteriaWiden(lngFrm_Top As Long, lngFrm_Left As Long, lngFrm_Width As Long, lngFrm_Height As Long, frm As Access.Form, frmCrit As Access.Form)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "CriteriaWiden"

        Dim lngForm_Width_New As Long
        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngLeft2 As Long, lngTop2 As Long, lngWidth2 As Long, lngHeight2 As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

12710   With frm

12720     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
12730       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
12740     End If

          ' ** I want to try to make sure it's centered in the window,
          ' ** even if it's off-center before widening.

          ' ** Variables are fed empty, then populated ByRef.
12750     GetFormDimensions frm, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.

          ' ** Variables are fed empty, then populated ByRef.
12760     GetAppDimensions lngLeft2, lngTop2, lngWidth2, lngHeight2  ' ** Module Procedure: modWindowFunctions.

          ' ** Let's give it some leeway.
12770     If (lngFrm_Width <= (lngWidth + (5& * lngTpp))) And (lngFrm_Width >= (lngWidth - (5& * lngTpp))) Then
            'If lngWidth = lngFrm_Width Then
            ' ** Original Criteria width is frmCrit.FilterBy_vline04.Left

12780       lngTmp01 = (frmCrit.Width - frmCrit.FilterBy_vline04.Left)  ' ** Additional width needed.
12790       lngForm_Width_New = (lngFrm_Width + lngTmp01)  ' ** Preliminary new width.

12800       If lngForm_Width_New > lngWidth Then
12810         lngForm_Width_New = (lngForm_Width_New + (7& * lngTpp))  ' ** Slight adjustment to assure scroll bar not active.
12820         lngTmp01 = (lngWidth2 - lngWidth)  ' ** Difference between form and Access window.
12830         If lngTmp01 > 0 Then
                ' ** Not all users have Access maximized under the forms.
12840           lngTmp02 = (lngTmp01 / 2)  ' ** Left for centered form, before widening.
12850         Else
12860           lngTmp02 = lngLeft  ' ** Just its current Left.
12870         End If
12880         lngTmp01 = (lngForm_Width_New - lngWidth)  ' ** Difference between current and new window.
12890         lngTmp03 = (lngTmp01 / 2)  ' ** Amount to move left to accommodate new width.
12900         lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
12910         lngMonitorNum = 1&: lngTmp04 = 0&
12920         EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
12930         If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
12940         If lngMonitorNum = 1& Then lngTmp04 = lngTop
12950         DoCmd.MoveSize (lngTmp02 - lngTmp03), lngTmp04, lngForm_Width_New, lngHeight  'lngTop
12960         If lngMonitorNum > 1& Then
12970           LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
12980         End If
12990         .Form_Resize  ' ** Form Procedure: frmTransaction_Audit.
13000       End If

13010     Else

13020       lngTmp01 = (lngWidth - lngFrm_Width)  ' ** Excess width to narrow.
13030       lngForm_Width_New = lngFrm_Width      ' ** Original width.

13040       If lngForm_Width_New < lngWidth Then
13050         lngTmp01 = (lngWidth2 - lngWidth)  ' ** Difference between form and Access window.
13060         If lngTmp01 > 0 Then
                ' ** Not all users have Access maximized under the forms.
13070           lngTmp02 = (lngTmp01 / 2)  ' ** Left for centered form, before widening.
13080         Else
13090           lngTmp02 = lngLeft  ' ** Just its current Left.
13100         End If
13110         lngTmp01 = (lngWidth - lngForm_Width_New)  ' ** Difference between current and new window.
13120         lngTmp03 = (lngTmp01 / 2)  ' ** Amount to move right to center new width.
13130         lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
13140         lngMonitorNum = 1&: lngTmp04 = 0&
13150         EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
13160         If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
13170         If lngMonitorNum = 1& Then lngTmp04 = lngTop
13180         DoCmd.MoveSize (lngTmp02 + lngTmp03), lngTmp04, lngForm_Width_New, lngHeight  'lngTop
13190         If lngMonitorNum > 1& Then
13200           LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
13210         End If
13220         .Form_Resize  ' ** Form Procedure: frmTransaction_Audit.
              ' ** Variables are fed empty, then populated ByRef.
13230         GetFormDimensions frm, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.
13240         lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
13250         lngMonitorNum = 1&: lngTmp04 = 0&
13260         EnumMonitors frm  ' ** Module Function: modMonitorFuncs.
13270         If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
              ' ** The first Form_Resize() doesn't appear to trigger.
13280         If lngMonitorNum = 1& Then lngTmp04 = lngTop
13290         DoCmd.MoveSize lngLeft, lngTmp04, lngWidth + lngTpp, lngHeight
13300         If lngMonitorNum > 1& Then
13310           LoadPosition .hwnd, frm.Name  ' ** Module Function: modMonitorFuncs.
13320         End If
13330         .Form_Resize  ' ** Form Procedure: frmTransaction_Audit.
13340       End If

13350     End If

          'FORM: lngLeft: 3225  lngTop: 210  lngWidth: 15150  lngHeight: 10140
          'APP : lngLeft: -120  lngTop: -120  lngWidth: 21840  lngHeight: 13290

13360   End With

EXITP:
13370   Exit Sub

ERRH:
13380   Select Case ERR.Number
        Case Else
13390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13400   End Select
13410   Resume EXITP

End Sub

Public Sub FormResize_TransAud(blnIsOpen As Boolean, blnFromCkgFlds As Boolean, lngFrm_Width As Long, lngSub_Width_Diff As Long, lngSub_Width As Long, lngSub_Width_New As Long, lngForm_Width As Long, lngForm_Width_New As Long, lngSub_Width_Min As Long, lngFrm_Height As Long, lngDetail_Height_Diff As Long, lngDetail_Height As Long, lngDetail_Height_New As Long, lngDetail_Height_Min As Long, lngCkgFldsBox_Height As Long, lngCkgFldsBox_Offset As Long, lngCritSub_Width As Long, lngHLine_Offset As Long, lngClose_Left As Long, lngSizable_Offset As Long, lngCritSub_Max As Long, lngSub_Height As Long, lngSelectBox_Top As Long, lngSelectLbl_Top As Long, lngSelectLegendLbl_Top As Long, lngSelectLbl2_Top As Long, lngNarrowFont_Top As Long, lngShareface_Top As Long, lngSelectAll_Top As Long, lngSelectNone_Top As Long, lngChkBoxLbl_TopOffset As Long, lngCkgFldsVLine_Top As Long, lngFrmFlds As Long, arr_varFrmFld As Variant, frm As Access.Form)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "FormResize_TransAud"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim blnShorter As Boolean, blnNarrower As Boolean
        Dim lngTmp01 As Long

13510   blnShorter = False: blnNarrower = False

13520   With frm
13530     If gblnClosing = False And blnFromCkgFlds = False Then

13540       GetFormDimensions frm, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.

13550       Set frmCrit = .frmTransaction_Audit_Sub_Criteria.Form

13560       Select Case blnIsOpen
            Case True  ' ** It hits this event several times before settling down!
              ' ** No adjustments necessary.
13570         blnIsOpen = False

13580       Case False  ' ** Not on open.

13590         If lngWidth > lngFrm_Width Then
13600           lngSub_Width_Diff = (lngWidth - lngFrm_Width)
13610           lngSub_Width_New = (lngSub_Width + lngSub_Width_Diff)
13620           lngForm_Width_New = (lngForm_Width + lngSub_Width_Diff)
13630         ElseIf lngWidth < lngFrm_Width Then
13640           lngSub_Width_Diff = (lngFrm_Width - lngWidth)
13650           lngSub_Width_New = (lngSub_Width - lngSub_Width_Diff)
13660           lngForm_Width_New = (lngForm_Width - lngSub_Width_Diff)
13670           If lngSub_Width_New < lngSub_Width_Min Then
13680             lngSub_Width_Diff = (lngSub_Width - lngSub_Width_Min)
13690             lngSub_Width_New = lngSub_Width_Min
13700           End If
13710         Else
13720           lngSub_Width_Diff = 0&
13730           lngSub_Width_New = lngSub_Width
13740           lngForm_Width_New = lngForm_Width
13750         End If

13760         If lngSub_Width_New < lngSub_Width Then
13770           blnNarrower = True
13780         Else
13790           blnNarrower = False
13800         End If

              ' ** We need to take into consideration whether the fields are open or closed.
13810         Select Case .ckgFlds
              Case True
13820           If lngHeight > lngFrm_Height Then
13830             lngDetail_Height_Diff = (lngHeight - lngFrm_Height)
13840             lngDetail_Height_New = (lngDetail_Height + lngDetail_Height_Diff)
13850           ElseIf lngHeight < lngFrm_Height Then
13860             lngDetail_Height_Diff = (lngFrm_Height - lngHeight)
13870             lngDetail_Height_New = (lngDetail_Height - lngDetail_Height_Diff)
13880             If lngDetail_Height_New < lngDetail_Height_Min Then
13890               lngDetail_Height_Diff = (lngDetail_Height - lngDetail_Height_Min)
13900               lngDetail_Height_New = lngDetail_Height_Min
13910             End If
13920           Else
13930             lngDetail_Height_Diff = 0&
13940             lngDetail_Height_New = lngDetail_Height
13950           End If
13960         Case False
13970           lngTmp01 = ((.ckgFlds_box.Top + lngCkgFldsBox_Height) + lngCkgFldsBox_Offset)  ' ** Original bottom of Detail.
13980           lngTmp01 = lngTmp01 - (.ckgFlds_hline04.Top + lngCkgFldsBox_Offset)  ' ** Should be difference from original height.
13990           If lngHeight > (lngFrm_Height - lngTmp01) Then
14000             lngDetail_Height_Diff = (lngHeight - (lngFrm_Height - lngTmp01))
14010             lngDetail_Height_New = (lngDetail_Height + lngDetail_Height_Diff)
14020           ElseIf lngHeight < (lngFrm_Height - lngTmp01) Then
14030             lngDetail_Height_Diff = ((lngFrm_Height - lngTmp01) - lngHeight)
14040             lngDetail_Height_New = (lngDetail_Height - lngDetail_Height_Diff)
14050             If lngDetail_Height_New < lngDetail_Height_Min Then
14060               lngDetail_Height_Diff = (lngDetail_Height - lngDetail_Height_Min)
14070               lngDetail_Height_New = lngDetail_Height_Min
14080             End If
14090           Else
14100             lngDetail_Height_Diff = 0&
14110             lngDetail_Height_New = lngDetail_Height
14120           End If
14130         End Select

14140         If lngDetail_Height_New < lngDetail_Height Then
14150           blnShorter = True
14160         Else
14170           blnShorter = False
14180         End If

14190         If lngSub_Width_Diff <> 0& Then
14200           Select Case blnNarrower
                Case True
14210             .frmTransaction_Audit_Sub.Width = lngSub_Width_New
14220             .frmTransaction_Audit_Sub_ds.Width = lngSub_Width_New
14230             .frmTransaction_Audit_Sub_box.Width = (lngSub_Width_New + (2& * lngTpp))
14240             .frmTransaction_Audit_Sub_Criteria.Width = (lngCritSub_Width - lngSub_Width_Diff)
14250             .frmTransaction_Audit_Sub_Criteria_box.Width = ((lngCritSub_Width - lngSub_Width_Diff) + (2& * lngTpp))
14260             .Header_vline01.Left = (lngSub_Width_New + lngHLine_Offset)
14270             .Header_vline02.Left = .Header_vline01.Left
14280             .Detail_vline01.Left = .Header_vline01.Left
14290             .Detail_vline02.Left = .Header_vline01.Left
14300             .Detail_vline03.Left = .Header_vline01.Left
14310             .Detail_vline03.Left = .Header_vline01.Left
14320             .Footer_vline01.Left = .Header_vline01.Left
14330             .Footer_vline02.Left = .Header_vline01.Left
14340             .Header_hline01.Width = .Header_vline01.Left
14350             .Header_hline02.Width = .Header_vline01.Left
14360             .Detail_hline01.Width = .Header_vline01.Left
14370             .Detail_hline02.Width = .Header_vline01.Left
14380             .Detail_hline03.Width = .Header_vline01.Left
14390             .Detail_hline04.Width = .Header_vline01.Left
14400             .Footer_hline01.Width = .Header_vline01.Left
14410             .Footer_hline02.Width = .Header_vline01.Left
14420             .Nav_box01.Width = .frmTransaction_Audit_Sub.Width
14430             .cmdClose.Left = (lngClose_Left - lngSub_Width_Diff)
14440             .Sizable_lbl1.Left = (.Header_vline01.Left - lngSizable_Offset)
14450             .Sizable_lbl2.Left = .Sizable_lbl1.Left
14460             .Nav_hline03.Width = .frmTransaction_Audit_Sub.Width
14470             .Width = lngForm_Width_New
14480           Case False
14490             .Width = lngForm_Width_New
14500             .frmTransaction_Audit_Sub.Width = lngSub_Width_New
14510             .frmTransaction_Audit_Sub_ds.Width = lngSub_Width_New
14520             .frmTransaction_Audit_Sub_box.Width = (lngSub_Width_New + (2& * lngTpp))
14530             lngTmp01 = (lngCritSub_Width + lngSub_Width_Diff)
14540             If lngTmp01 > lngCritSub_Max Then lngTmp01 = lngCritSub_Max
14550             .frmTransaction_Audit_Sub_Criteria.Width = lngTmp01
14560             .frmTransaction_Audit_Sub_Criteria_box.Width = (lngTmp01 + (2& * lngTpp))
14570             .Header_vline01.Left = (lngSub_Width_New + lngHLine_Offset)
14580             .Header_vline02.Left = .Header_vline01.Left
14590             .Detail_vline01.Left = .Header_vline01.Left
14600             .Detail_vline02.Left = .Header_vline01.Left
14610             .Detail_vline03.Left = .Header_vline01.Left
14620             .Detail_vline04.Left = .Header_vline01.Left
14630             .Footer_vline01.Left = .Header_vline01.Left
14640             .Footer_vline02.Left = .Header_vline01.Left
14650             .Header_hline01.Width = .Header_vline01.Left
14660             .Header_hline02.Width = .Header_vline01.Left
14670             .Detail_hline01.Width = .Header_vline01.Left
14680             .Detail_hline02.Width = .Header_vline01.Left
14690             .Detail_hline03.Width = .Header_vline01.Left
14700             .Detail_hline04.Width = .Header_vline01.Left
14710             .Footer_hline01.Width = .Header_vline01.Left
14720             .Footer_hline02.Width = .Header_vline01.Left
14730             .Nav_box01.Width = .frmTransaction_Audit_Sub.Width
14740             .cmdClose.Left = (lngClose_Left + lngSub_Width_Diff)
14750             .Sizable_lbl1.Left = (.Header_vline01.Left - lngSizable_Offset)
14760             .Sizable_lbl2.Left = .Sizable_lbl1.Left
14770             .Nav_hline03.Width = .frmTransaction_Audit_Sub.Width
14780           End Select  ' ** blnNarrower.
14790         End If  ' ** lngSub_Width_Diff.

14800         If lngDetail_Height_Diff <> 0& Then
14810           TransAudit_Resize blnShorter, lngDetail_Height_New, lngDetail_Height_Diff, lngSub_Height, _
                  lngSelectBox_Top, lngSelectLbl_Top, lngSelectLegendLbl_Top, lngSelectLbl2_Top, lngNarrowFont_Top, _
                  lngShareface_Top, lngSelectAll_Top, lngSelectNone_Top, lngChkBoxLbl_TopOffset, lngCkgFldsVLine_Top, _
                  lngFrmFlds, arr_varFrmFld, frm  ' ** Procedure: Below.
14820         End If  ' ** lngDetail_Height_Diff.

14830       End Select  ' ** blnIsOpen.

14840       lngLeft = 0&: lngTop = 0&: lngWidth = 0&: lngHeight = 0&
14850       GetFormDimensions frm, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.
14860       lngTmp01 = (frmCrit.Width - frmCrit.FilterBy_vline04.Left)  ' ** Additional width needed.
14870       lngForm_Width_New = ((lngFrm_Width + lngTmp01) + (7& * lngTpp))  ' ** Preliminary new width.
14880       If lngForm_Width_New > lngWidth Then
14890         .cmdWidenToCriteria.Enabled = True
14900         .cmdWidenToCriteria_raised_img.Visible = True
14910         .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = False
14920         .cmdWidenToCriteria_raised_focus_img.Visible = False
14930         .cmdWidenToCriteria_raised_focus_dots_img.Visible = False
14940         .cmdWidenToCriteria_sunken_focus_dots_img.Visible = False
14950         .cmdWidenToCriteria_raised_img_dis.Visible = False
14960       Else
14970 On Error Resume Next
14980         If ERR.Number <> 0 Then
14990 On Error GoTo ERRH
15000           .FocusHolder.SetFocus
15010         Else
15020 On Error GoTo ERRH
15030         End If
15040         .cmdWidenToCriteria_raised_img_dis.Visible = True
15050         .cmdWidenToCriteria_raised_img.Visible = False
15060         .cmdWidenToCriteria_raised_semifocus_dots_img.Visible = False
15070         .cmdWidenToCriteria_raised_focus_img.Visible = False
15080         .cmdWidenToCriteria_raised_focus_dots_img.Visible = False
15090         .cmdWidenToCriteria_sunken_focus_dots_img.Visible = False
15100       End If

15110     End If  ' ** gblnClosing.
15120   End With

EXITP:
15130   Set frmCrit = Nothing
15140   Exit Sub

ERRH:
15150   Select Case ERR.Number
        Case Else
15160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15170   End Select
15180   Resume EXITP

End Sub

Public Function DoReport_TA(dblFilterRecs As Double, lngFrmFlds As Long, lngFrmFlds_ds As Long, arr_varTmp01 As Variant, arr_varTmp02 As Variant, frm As Access.Form) As Boolean

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport_TA"

        Dim ctl As Access.Control
        Dim blnAllowIfNoFilter As Boolean
        Dim arr_varFrmFldX() As Variant, arr_varFrmFld_dsX() As Variant
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

15210   blnRetVal = True

15220   With frm

15230     ReDim arr_varFrmFldX(FM_ELEMS, 0)
15240     If lngFrmFlds > 0& Then
15250       lngE = lngFrmFlds - 1&
15260       ReDim arr_varFrmFldX(FM_ELEMS, lngE)
15270       For lngX = 0& To (lngFrmFlds - 1&)
15280         For lngY = 0& To FM_ELEMS
15290           arr_varFrmFldX(lngY, lngX) = arr_varTmp01(lngY, lngX)
15300         Next
15310       Next
15320     End If
15330     ReDim arr_varFrmFld_dsX(FM_ELEMS, 0)
15340     If lngFrmFlds_ds > 0& Then
15350       lngE = lngFrmFlds_ds - 1&
15360       ReDim arr_varFrmFld_dsX(FM_ELEMS, lngE)
15370       For lngX = 0& To (lngFrmFlds_ds - 1&)
15380         For lngY = 0& To FM_ELEMS
15390           arr_varFrmFld_dsX(lngY, lngX) = arr_varTmp02(lngY, lngX)
15400         Next
15410       Next
15420     End If

          ' ** See if I can let it print if there's no filter,
          ' ** just the entire, unfiltered list.
15430     If dblFilterRecs = 0# Then
15440       blnAllowIfNoFilter = True
15450       Select Case .opgView
            Case .opgView_optForm.OptionValue
15460         For Each ctl In .frmTransaction_Audit_Sub.Form.Section("Detail").Controls
15470           With ctl
15480             If .Visible = True Then
15490               If .ControlType = acTextBox Then
15500                 If .BackColor = CLR_GRY2 Then
15510                   blnAllowIfNoFilter = False
15520                   Exit For
15530                 End If
15540               End If
15550             End If
15560           End With
15570         Next
15580       Case .opgView_optDatasheet.OptionValue
15590         For Each ctl In .frmTransaction_Audit_Sub_ds.Form.Section("Detail").Controls
15600           With ctl
15610             If .Visible = True Then
15620               If .ControlType = acTextBox Then
15630                 If .BackColor = CLR_GRY2 Then
15640                   blnAllowIfNoFilter = False
15650                   Exit For
15660                 End If
15670               End If
15680             End If
15690           End With
15700         Next
15710       End Select
15720     End If

15730     If lngFrmFlds = 0& Or lngFrmFlds_ds = 0& Then  ' ** I don't really know why this gets lost!

15740       lngFrmFlds = 0&
15750       ReDim arr_varFrmFldX(FM_ELEMS, 0)

15760       lngFrmFlds_ds = 0&
15770       ReDim arr_varFrmFld_dsX(FM_ELEMS, 0)

            ' ****************************************************
            ' ** Array:  arr_varFrmFldX(),  arr_varFrmFld_dsX()
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

15780       For lngX = 1& To 2&
15790         Select Case lngX
              Case 1&

                ' ** Get a list of the form data fields.
15800           For Each ctl In .frmTransaction_Audit_Sub.Form.Section("Detail").Controls
15810             With ctl
15820               Select Case .Name
                    Case "FocusHolder", "FocusHolder2", "assetno", "revcode_ID", "revcode_TYPE", "taxcode", "taxcode_type"
                      ' ** Skip these.
15830               Case Else
15840                 Select Case .ControlType
                      Case acTextBox, acCheckBox, acComboBox, acOptionButton, acCommandButton
15850                   lngFrmFlds = lngFrmFlds + 1&
15860                   lngE = lngFrmFlds - 1&
15870                   ReDim Preserve arr_varFrmFldX(FM_ELEMS, lngE)
15880                   arr_varFrmFldX(FM_FLD_NAM, lngE) = .Name
15890                   arr_varFrmFldX(FM_FLD_TAB, lngE) = .TabIndex
15900                   arr_varFrmFldX(FM_FLD_VIS, lngE) = .Visible
15910                   arr_varFrmFldX(FM_CHK_NAM, lngE) = vbNullString
15920                   arr_varFrmFldX(FM_CHK_VAL, lngE) = CBool(False)
15930                 Case Else
                        ' ** acLine.
15940                 End Select
15950               End Select
15960             End With
15970           Next

                ' ** Get a list of the form data field checkboxes.
15980           lngChkCnt = 0&
15990           For Each ctl In .frmTransaction_Audit_Sub.Form.Section("FormHeader").Controls
16000             With ctl
16010               If Right(.Name, 4) = "_chk" Then
                      ' ** For non-visible ID fields, checkbox elements will remain vbNullString and False.
16020                 lngChkCnt = lngChkCnt + 1&
16030                 strTmp01 = Left(.Name, (Len(.Name) - 4))
16040                 For lngY = 0& To (lngFrmFlds - 1&)
16050                   If arr_varFrmFldX(FM_FLD_NAM, lngY) = strTmp01 Then
16060                     arr_varFrmFldX(FM_CHK_NAM, lngY) = .Name
16070                     arr_varFrmFldX(FM_CHK_VAL, lngY) = CBool(.Value)
16080                     Exit For
16090                   End If
16100                 Next
16110               End If
16120             End With
16130           Next

16140         Case 2&

                ' ** Get a list of the form data fields.
16150           For Each ctl In .frmTransaction_Audit_Sub_ds.Form.Section("Detail").Controls
16160             With ctl
16170               Select Case .Name
                    Case "FocusHolder", "FocusHolder2", "assetno", "revcode_ID", "revcode_TYPE", "taxcode", "taxcode_type"
                      ' ** Skip these.
16180               Case Else
16190                 Select Case .ControlType
                      Case acTextBox, acCheckBox, acComboBox, acOptionButton, acCommandButton
16200                   lngFrmFlds_ds = lngFrmFlds_ds + 1&
16210                   lngE = lngFrmFlds_ds - 1&
16220                   ReDim Preserve arr_varFrmFld_dsX(FM_ELEMS, lngE)
16230                   arr_varFrmFld_dsX(FM_FLD_NAM, lngE) = .Name
16240                   arr_varFrmFld_dsX(FM_FLD_TAB, lngE) = .TabIndex
16250                   arr_varFrmFld_dsX(FM_FLD_VIS, lngE) = .Visible
16260                   arr_varFrmFld_dsX(FM_CHK_NAM, lngE) = vbNullString
16270                   arr_varFrmFld_dsX(FM_CHK_VAL, lngE) = CBool(False)
16280                 Case Else
                        ' ** acLine.
16290                 End Select
16300               End Select
16310             End With
16320           Next

                ' ** Get a list of the form data field checkboxes.
16330           lngChkCnt = 0&
16340           For Each ctl In .frmTransaction_Audit_Sub_ds.Form.Section("FormHeader").Controls
16350             With ctl
16360               If Right(.Name, 4) = "_chk" Then
                      ' ** For non-visible ID fields, checkbox elements will remain vbNullString and False.
16370                 lngChkCnt = lngChkCnt + 1&
16380                 strTmp01 = Left(.Name, (Len(.Name) - 4))
16390                 For lngY = 0& To (lngFrmFlds_ds - 1&)
16400                   If arr_varFrmFld_dsX(FM_FLD_NAM, lngY) = strTmp01 Then
16410                     arr_varFrmFld_dsX(FM_CHK_NAM, lngY) = .Name
16420                     arr_varFrmFld_dsX(FM_CHK_VAL, lngY) = CBool(.Value)
16430                     Exit For
16440                   End If
16450                 Next
16460               End If
16470             End With
16480           Next

16490         End Select

              ' ** Sort the list, left to right on form.
16500         .FormFields_Sort lngX  ' ** Form Procedure: frmTransaction_Audit.

16510       Next  ' ** lngX

16520     End If  ' ** lngFrmFlds.

          ' ** Make sure there are records returned.
16530     If dblFilterRecs > 0# Or blnAllowIfNoFilter = True Then

16540       For lngX = 1& To 2&
16550         Select Case lngX
              Case 1&

16560           varTmp00 = .frmTransaction_Audit_Sub.Form.ledger_description_max
16570           If IsNull(varTmp00) = False Then
16580             .ledger_description_max = varTmp00
16590           End If
16600           varTmp00 = .frmTransaction_Audit_Sub.Form.RecurringItem_max
16610           If IsNull(varTmp00) = False Then
16620             .RecurringItem_max = varTmp00
16630           End If
16640           varTmp00 = .frmTransaction_Audit_Sub.Form.ledger_HIDDEN_min
16650           If IsNull(varTmp00) = False Then
16660             .ledger_HIDDEN_min = varTmp00
16670           End If
16680           varTmp00 = .frmTransaction_Audit_Sub.Form.ForExCnt_sum
16690           If IsNull(varTmp00) = False Then
16700             .ForExCnt_sum = varTmp00
16710           End If

                ' ** Make sure there are fields checked.
16720           lngVisCnt1 = 0&
16730           For lngY = 0& To (lngFrmFlds - 1&)
16740             If arr_varFrmFldX(FM_CHK_NAM, lngY) <> vbNullString Then
16750               Set ctl = .frmTransaction_Audit_Sub.Form.Controls(arr_varFrmFldX(FM_CHK_NAM, lngY))
16760               With ctl
16770                 arr_varFrmFldX(FM_CHK_VAL, lngY) = CBool(.Value)
16780                 If arr_varFrmFldX(FM_CHK_VAL, lngY) = True Then
16790                   lngVisCnt1 = lngVisCnt1 + 1&
16800                 End If
16810               End With
16820             End If
16830           Next

16840         Case 2&

16850 On Error Resume Next
16860           varTmp00 = .frmTransaction_Audit_Sub_ds.Form.ledger_description_max
16870 On Error GoTo ERRH
16880           If IsNull(varTmp00) = False Then
16890             .ledger_description_max = varTmp00
16900           End If
16910 On Error Resume Next
16920           varTmp00 = .frmTransaction_Audit_Sub_ds.Form.RecurringItem_max
16930 On Error GoTo ERRH
16940           If IsNull(varTmp00) = False Then
16950             .RecurringItem_max = varTmp00
16960           End If
16970 On Error Resume Next
16980           varTmp00 = .frmTransaction_Audit_Sub_ds.Form.ledger_HIDDEN_min
16990 On Error GoTo ERRH
17000           If IsNull(varTmp00) = False Then
17010             .ledger_HIDDEN_min = varTmp00
17020           End If
17030 On Error Resume Next
17040           varTmp00 = .frmTransaction_Audit_Sub_ds.Form.ForExCnt_sum
17050 On Error GoTo ERRH
17060           If IsNull(varTmp00) = False Then
17070             .ForExCnt_sum = varTmp00
17080           End If

                ' ** Make sure there are fields checked.
17090           lngVisCnt2 = 0&
17100           For lngY = 0& To (lngFrmFlds_ds - 1&)
17110             If arr_varFrmFld_dsX(FM_CHK_NAM, lngY) <> vbNullString Then
17120               Set ctl = .frmTransaction_Audit_Sub_ds.Form.Controls(arr_varFrmFld_dsX(FM_CHK_NAM, lngY))
17130               With ctl
17140                 arr_varFrmFld_dsX(FM_CHK_VAL, lngY) = CBool(.Value)
17150                 If arr_varFrmFld_dsX(FM_CHK_VAL, lngY) = True Then
17160                   lngVisCnt2 = lngVisCnt2 + 1&
17170                 End If
17180               End With
17190             End If
17200           Next

17210         End Select

17220       Next  ' ** lngX

17230       Select Case .opgView
            Case .opgView_optForm.OptionValue
17240         If lngVisCnt1 > 0& Then
                ' ** OK to go.
17250         Else
17260           blnRetVal = False
17270           Beep
17280           MsgBox "There are no fields checked to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
17290         End If
17300       Case .opgView_optDatasheet.OptionValue
17310         If lngVisCnt2 > 0& Then
                ' ** OK to go.
17320         Else
17330           blnRetVal = False
17340           Beep
17350           MsgBox "There are no fields checked to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
17360         End If
17370       End Select

17380     Else
17390       blnRetVal = False
17400       Beep
17410       MsgBox "There are no records to print.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
17420     End If

17430     arr_varTmp01 = arr_varFrmFldX
17440     arr_varTmp02 = arr_varFrmFld_dsX

17450   End With

EXITP:
17460   Set ctl = Nothing
17470   DoReport_TA = blnRetVal
17480   Exit Function

ERRH:
17490   blnRetVal = False
17500   Select Case ERR.Number
        Case Else
17510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17520   End Select
17530   Resume EXITP

End Function

Public Sub Calendar_Handler_TA(strProc As String, lngTmp01 As Long, arr_varTmp02 As Variant, clsMonthClass As clsMonthCal, frm As Access.Form, Optional varAble As Variant)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_TA"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnAble As Boolean
        Dim strCtl As String, strMode As String
        Dim intPos01 As Integer, intNum As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, strTmp05 As String, strTmp06 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCal().
        Const C_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const C_CNAM  As Integer = 0
        Const C_FOCUS As Integer = 1
        Const C_DOWN  As Integer = 2
        Const C_ABLE  As Integer = 3
        Const C_FLD   As Integer = 4

17610   With frm
17620     intPos01 = InStr(strProc, "_")
17630     If intPos01 > 0 Then
17640       strCtl = Left(strProc, (intPos01 - 1))
17650       strMode = Mid(strProc, (intPos01 + 1))
17660       If strCtl <> "Form" Then
17670         intNum = Val(Right(strCtl, 1))
17680         strTmp01 = strCtl & "_raised_img"                 ' ** .cmdCalendar1_raised_img
17690         strTmp02 = strCtl & "_raised_semifocus_dots_img"  ' ** .cmdCalendar1_raised_semifocus_dots_img
17700         strTmp03 = strCtl & "_raised_focus_img"           ' ** .cmdCalendar1_raised_focus_img
17710         strTmp04 = strCtl & "_raised_focus_dots_img"      ' ** .cmdCalendar1_raised_focus_dots_img
17720         strTmp05 = strCtl & "_sunken_focus_dots_img"      ' ** .cmdCalendar1_sunken_focus_dots_img
17730         strTmp06 = strCtl & "_raised_img_dis"             ' ** .cmdCalendar1_raised_img_dis
17740       End If
17750       Select Case strMode
            Case "Load"
              ' ** Load the calendar array.
17760         If lngCals = 0& Or IsEmpty(arr_varCal) = True Then
17770           lngCals = 0&
17780           ReDim arr_varCal(C_ELEMS, 0)
                ' ** This will be a 1-based array.
17790           For lngX = 1& To 8&
17800             lngCals = lngCals + 1&
17810             lngE = lngCals
17820             ReDim Preserve arr_varCal(C_ELEMS, lngE)
17830             arr_varCal(C_CNAM, lngE) = "cmdCalendar" & CStr(lngX)
17840             arr_varCal(C_FOCUS, lngE) = CBool(False)
17850             arr_varCal(C_DOWN, lngE) = CBool(False)
17860             arr_varCal(C_ABLE, lngE) = CBool(True)
17870             Select Case lngX
                  Case 1&
17880               strTmp01 = "TransDateStart"
17890             Case 2&
17900               strTmp01 = "TransDateEnd"
17910             Case 3&
17920               strTmp01 = "AssetDateStart"
17930             Case 4&
17940               strTmp01 = "AssetDateEnd"
17950             Case 5&
17960               strTmp01 = "PurchaseDateStart"
17970             Case 6&
17980               strTmp01 = "PurchaseDateEnd"
17990             Case 7&
18000               strTmp01 = "PostedDateStart"
18010             Case 8&
18020               strTmp01 = "PostedDateEnd"
18030             End Select
18040             arr_varCal(C_FLD, lngE) = strTmp01
18050           Next
18060         End If
18070       Case "Click"
              ' ** cmdCalendar1_Click.
18080         datStartDate = Date
18090         datEndDate = 0
18100         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
18110         If blnRetVal = True Then
18120           .Controls(arr_varCal(C_FLD, intNum)) = datStartDate
18130         Else
18140           .Controls(arr_varCal(C_FLD, intNum)) = CDate(Format(Date, "mm/dd/yyyy"))
18150         End If
18160       Case "GotFocus"
              ' ** cmdCalendar1_GotFocus.
18170         arr_varCal(C_FOCUS, intNum) = CBool(True)
18180         .Controls(strTmp02).Visible = True  ' ** _raised_semifocus_dots_img.
18190         .Controls(strTmp01).Visible = False
18200         .Controls(strTmp03).Visible = False
18210         .Controls(strTmp04).Visible = False
18220         .Controls(strTmp05).Visible = False
18230         .Controls(strTmp06).Visible = False
18240       Case "LostFocus"
              ' ** cmdCalendar1_LostFocus.
18250         .Controls(strTmp01).Visible = True  ' ** _raised_img.
18260         .Controls(strTmp02).Visible = False
18270         .Controls(strTmp03).Visible = False
18280         .Controls(strTmp04).Visible = False
18290         .Controls(strTmp05).Visible = False
18300         .Controls(strTmp06).Visible = False
18310         arr_varCal(C_FOCUS, intNum) = CBool(False)
18320       Case "MouseMove"
              ' ** cmdCalendar1_MouseMove.
18330         If arr_varCal(C_DOWN, intNum) = False Then
18340           Select Case arr_varCal(C_FOCUS, intNum)
                Case True
18350             .Controls(strTmp04).Visible = True  ' ** _raised_focus_dots_img.
18360             .Controls(strTmp03).Visible = False
18370           Case False
18380             .Controls(strTmp03).Visible = True  ' ** _raised_focus_img.
18390             .Controls(strTmp04).Visible = False
18400           End Select
18410           .Controls(strTmp01).Visible = False
18420           .Controls(strTmp02).Visible = False
18430           .Controls(strTmp05).Visible = False
18440           .Controls(strTmp06).Visible = False
18450           Select Case intNum
                Case 3
18460             If .cmdCalendar5_raised_focus_dots_img.Visible = True Or .cmdCalendar5_raised_focus_img.Visible = True Then
18470               Select Case arr_varCal(C_FOCUS, 5)
                    Case True
18480                 .cmdCalendar5_raised_semifocus_dots_img.Visible = True
18490                 .cmdCalendar5_raised_img.Visible = False
18500               Case False
18510                 .cmdCalendar5_raised_img.Visible = True
18520                 .cmdCalendar5_raised_semifocus_dots_img.Visible = False
18530               End Select
18540               .cmdCalendar5_raised_focus_img.Visible = False
18550               .cmdCalendar5_raised_focus_dots_img.Visible = False
18560               .cmdCalendar5_sunken_focus_dots_img.Visible = False
18570               .cmdCalendar5_raised_img_dis.Visible = False
18580             End If
18590           Case 4
18600             If .cmdCalendar6_raised_focus_dots_img.Visible = True Or .cmdCalendar6_raised_focus_img.Visible = True Then
18610               Select Case arr_varCal(C_FOCUS, 6)
                    Case True
18620                 .cmdCalendar6_raised_semifocus_dots_img.Visible = True
18630                 .cmdCalendar6_raised_img.Visible = False
18640               Case False
18650                 .cmdCalendar6_raised_img.Visible = True
18660                 .cmdCalendar6_raised_semifocus_dots_img.Visible = False
18670               End Select
18680               .cmdCalendar6_raised_focus_img.Visible = False
18690               .cmdCalendar6_raised_focus_dots_img.Visible = False
18700               .cmdCalendar6_sunken_focus_dots_img.Visible = False
18710               .cmdCalendar6_raised_img_dis.Visible = False
18720             End If
18730           Case 5
18740             If .cmdCalendar3_raised_focus_dots_img.Visible = True Or .cmdCalendar3_raised_focus_img.Visible = True Then
18750               Select Case arr_varCal(C_FOCUS, 3)
                    Case True
18760                 .cmdCalendar3_raised_semifocus_dots_img.Visible = True
18770                 .cmdCalendar3_raised_img.Visible = False
18780               Case False
18790                 .cmdCalendar3_raised_img.Visible = True
18800                 .cmdCalendar3_raised_semifocus_dots_img.Visible = False
18810               End Select
18820               .cmdCalendar3_raised_focus_img.Visible = False
18830               .cmdCalendar3_raised_focus_dots_img.Visible = False
18840               .cmdCalendar3_sunken_focus_dots_img.Visible = False
18850               .cmdCalendar3_raised_img_dis.Visible = False
18860             End If
18870           Case 6
18880             If .cmdCalendar4_raised_focus_dots_img.Visible = True Or .cmdCalendar4_raised_focus_img.Visible = True Then
18890               Select Case arr_varCal(C_FOCUS, 4)
                    Case True
18900                 .cmdCalendar4_raised_semifocus_dots_img.Visible = True
18910                 .cmdCalendar4_raised_img.Visible = False
18920               Case False
18930                 .cmdCalendar4_raised_img.Visible = True
18940                 .cmdCalendar4_raised_semifocus_dots_img.Visible = False
18950               End Select
18960               .cmdCalendar4_raised_focus_img.Visible = False
18970               .cmdCalendar4_raised_focus_dots_img.Visible = False
18980               .cmdCalendar4_sunken_focus_dots_img.Visible = False
18990               .cmdCalendar4_raised_img_dis.Visible = False
19000             End If
19010           End Select
19020         End If
19030       Case "MouseDown"
              ' ** cmdCalendar1_MouseDown.
19040         arr_varCal(C_DOWN, intNum) = CBool(True)
19050         .Controls(strTmp05).Visible = True  ' ** _sunken_focus_dots_img.
19060         .Controls(strTmp01).Visible = False
19070         .Controls(strTmp02).Visible = False
19080         .Controls(strTmp03).Visible = False
19090         .Controls(strTmp04).Visible = False
19100         .Controls(strTmp06).Visible = False
19110       Case "MouseUp"
              ' ** cmdCalendar1_MouseUp.
19120         .Controls(strTmp04).Visible = True  ' ** _raised_focus_dots_img.
19130         .Controls(strTmp01).Visible = False
19140         .Controls(strTmp02).Visible = False
19150         .Controls(strTmp03).Visible = False
19160         .Controls(strTmp05).Visible = False
19170         .Controls(strTmp06).Visible = False
19180         arr_varCal(C_DOWN, intNum) = CBool(False)
19190       Case "Disable"
              ' ** Disable button.
19200         Select Case IsMissing(varAble)
              Case True
19210           blnAble = True
19220         Case False
19230           blnAble = varAble
19240         End Select
19250         arr_varCal(C_ABLE, intNum) = blnAble
19260         Select Case blnAble
              Case True
19270           .Controls(strTmp01).Visible = True  ' ** _raised_img.
19280           .Controls(strTmp06).Visible = False
19290         Case False
19300           .Controls(strTmp06).Visible = True  ' ** _raised_img_dis.
19310           .Controls(strTmp01).Visible = False
19320         End Select
19330         .Controls(strTmp02).Visible = False
19340         .Controls(strTmp03).Visible = False
19350         .Controls(strTmp04).Visible = False
19360         .Controls(strTmp05).Visible = False
19370       Case Else
19380         Beep
19390         MsgBox "Unknown calling procedure.", vbInformation + vbOKOnly, "Calendar_Handler"
19400       End Select
19410     End If
19420     lngTmp01 = lngCals
19430     arr_varTmp02 = arr_varCal
19440   End With

EXITP:
19450   Exit Sub

ERRH:
19460   Select Case ERR.Number
        Case Else
19470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19480   End Select
19490   Resume EXITP

End Sub

Public Sub ClearAll_TA(strFilter01 As String, dblFilterRecs As Double, frm As Access.Form, frmCrit As Access.Form)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "ClearAll_TA"

        Dim frm1 As Access.Form, frm2 As Access.Form
        Dim ctls As Object, ctl1 As Access.Control, ctls2 As Object, ctl2 As Access.Control
        Dim lngForeColor As Long
        Dim strPrintChk As String
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long

19510   With frm

19520     DoCmd.Hourglass True
19530     DoEvents

19540     If CLR_DISABLED_FG = 0& Or CLR_DISABLED_BG = 0& Then
19550       CLR_DISABLED_FG = CLR_DKGRY
19560       CLR_DISABLED_BG = CLR_LTTEAL
19570     End If

19580     Set frm1 = .frmTransaction_Audit_Sub.Form
19590     Set frm2 = .frmTransaction_Audit_Sub_ds.Form

19600     arr_varFilt = Empty
19610     arr_varFilt = FilterRecs_GetArr(1)  ' ** Function: Above.
19620     If IsEmpty(arr_varFilt) = True Then
19630       FilterRecs_Load  ' ** Procedure: Above.
19640       DoEvents
19650     End If
19660     lngFilts = UBound(arr_varFilt, 2) + 1&

19670     If IsNothing(frmCrit) = True Then  ' ** modUtilities.
19680       Set frmCrit = .frmTransaction_Audit_Sub_Criteria.Form
19690     End If

          ' ******************************************************
          ' ** Array: arr_varFilt()
          ' **
          ' **   Field  Element  Name                 Constant
          ' **   =====  =======  ===================  ==========
          ' **     1       0     taf_index            F_NAM
          ' **     2       1     vbdec_name1          F_NAM
          ' **     3       2     vbdec_value1         F_CONST
          ' **     4       3     ctl_name1            F_CTL
          ' **     5       4     ctl_name_lbl1        F_CLBL
          ' **     6       5     fld_ctl_name1        F_FLD
          ' **     7       6     fld_ctl_name_lbl1    F_FLBL
          ' **     8       7     ctl_name2            F_CTL2
          ' **     9       8     ctl_name_lbl2        F_CLBL2
          ' **    10       9     fld_ctl_name2        F_FLD2
          ' **    11      10     fld_ctl_name_lbl2    F_FLBL2
          ' **    12      11     ctl_name_lbl3        F_CLBL3
          ' **    13      12     fld_ctl_name3        F_FLD3
          ' **    14      13     fld_ctl_name_lbl3    F_FLBL3
          ' **
          ' ******************************************************
19700     Set ctls = frmCrit.Controls
          ' ** We should only need to do this once.
19710     frmCrit.FocusHolder.SetFocus
19720     For lngY = 0& To (lngFilts - 1&)
19730       Set ctl1 = ctls(arr_varFilt(F_CTL, lngY))
19740       With ctl1
19750         Select Case .ControlType
              Case acComboBox, acTextBox
19760           .Value = Null
19770           If .Name = "cmbJournalType1" Then
19780             With frmCrit
19790               .cmbJournalType2 = Null
19800               .cmbJournalType2.Enabled = False
19810               .cmbJournalType2.BorderColor = WIN_CLR_DISR
19820               .cmbJournalType2.BackStyle = acBackStyleTransparent
19830               .cmbJournalType3 = Null
19840               .cmbJournalType3.Enabled = False
19850               .cmbJournalType3.BorderColor = WIN_CLR_DISR
19860               .cmbJournalType3.BackStyle = acBackStyleTransparent
19870             End With
19880           End If
19890         Case acCheckBox
19900           Select Case .Name
                Case "chkRevcodeType_Income", "chkRevcodeType_Expense"
19910             .Value = True  ' ** Default to include both.
19920           End Select
19930         Case acOptionGroup
19940           If .Name = "opgHidden" Then
19950             .Value = 1  ' ** Default to Include.
19960           End If
19970         End Select
19980         If arr_varFilt(F_CLBL, lngY) <> vbNullString Then
19990           Set ctl2 = .Controls(arr_varFilt(F_CLBL, lngY))
20000           With ctl2
20010             .ForeColor = CLR_VDKGRY
20020             If .Name = "cmbJournalType1_lbl" Then
20030               With frmCrit
20040                 .cmbJournalType_lbl.ForeColor = CLR_VDKGRY
20050                 .cmbJournalType2_lbl.ForeColor = CLR_VDKGRY
20060                 .cmbJournalType3_lbl.ForeColor = CLR_VDKGRY
20070               End With
20080             End If
20090           End With
20100         End If
20110         If arr_varFilt(F_CLBL3, lngY) <> vbNullString Then
20120           Set ctl2 = ctls(arr_varFilt(F_CLBL3, lngY))
20130           With ctl2
20140             .ForeColor = CLR_VDKGRY
20150           End With
20160         End If
20170       End With
20180       Set ctl1 = Nothing: Set ctl2 = Nothing
20190     Next  ' ** lngY.
20200     DoEvents

20210     If gblnMessage = False Then
20220       frmCrit.opgAccountSource_AfterUpdate  ' ** Procedure: Below.
20230       frmCrit.opgAssetSource_AfterUpdate  ' ** Procedure: Below.
20240       frmCrit.chkRevcodeType_Income_AfterUpdate  ' ** Procedure: Below.
20250       frmCrit.chkRevcodeType_Income_AfterUpdate  ' ** Procedure: Below.
20260     End If
20270     DoEvents
20280     strFilter01 = vbNullString
20290     dblFilterRecs = 0#
20300     .CurrentFilter1 = vbNullString
20310     .CurrentFilter2 = vbNullString
20320     frm1.FilterRecs_Set vbNullString, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
20330     frm1.Filter = vbNullString
20340     frm1.FilterOn = False
20350     frm2.FilterRecs_Set vbNullString, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
20360     frm2.Filter = vbNullString
20370     frm2.FilterOn = False
20380     frm2.Requery

20390     strTmp01 = frm2.RecordSource
20400     frm2.RecordSource = strTmp01
20410     For lngX = 1& To 2&
20420       Select Case lngX
            Case 1&
20430         Set ctls = frm1.Section("Detail").Controls
20440       Case 2&
20450         Set ctls = frm2.Section("Detail").Controls
20460       End Select
20470       For Each ctl1 In ctls
20480         With ctl1
20490           If .Visible = True Then
20500             If .ControlType = acTextBox Then
20510               .BackColor = CLR_WHT
20520             End If
20530           End If
20540         End With
20550       Next
20560     Next  ' ** lngX.
20570     DoEvents

20580     For lngX = 1& To 1&  ' ** We really only need to do this once, since the datasheet doesn't show FormHeader.
20590       Select Case lngX
            Case 1&
20600         Set ctls2 = frm1.Section("FormHeader").Controls
20610       Case 2&
20620         Set ctls2 = frm2.Section("FormHeader").Controls
20630       End Select
20640       For Each ctl1 In ctls2
20650         With ctl1
20660           If .ControlType = acLabel Then
20670             If .Name <> "Sort_lbl" Then
20680               strPrintChk = Left(.Name, (InStr(.Name, "_lbl") - 1))
20690               strPrintChk = strPrintChk & "_chk"
20700               If ControlExists(strPrintChk, ctls2) = True Then  ' ** Module Function: modFileUtilities.
20710                 Set ctl2 = ctls2(strPrintChk)
20720                 Select Case ctl2
                      Case True
20730                   lngForeColor = CLR_DKGRY2
20740                 Case False
20750                   lngForeColor = WIN_CLR_DISF
20760                 End Select
20770                 .ForeColor = lngForeColor
20780                 Select Case .Name
                      Case "journalno_lbl", "journaltype_lbl", "transdate_lbl", "accountno_lbl"
20790                   ctls2(.Name & "2").ForeColor = lngForeColor
20800                 End Select
20810               Else
20820                 lngForeColor = CLR_DKGRY2
20830                 .ForeColor = lngForeColor
20840               End If
20850             End If
20860           End If
20870         End With
20880       Next
20890     Next  ' ** lngX.
20900     DoEvents

20910     .FilterRecs = "All"
20920     .FilterRecs.ForeColor = CLR_DISABLED_FG
20930     .FilterRecs.BackColor = CLR_DISABLED_BG
20940     .FilterRecs_lbl.ForeColor = CLR_VDKGRY
20950     .frmTransaction_Audit_Sub_Criteria.SetFocus
20960     frmCrit.FocusHolder.SetFocus
20970     If frmCrit.journalno.Enabled = True Then
20980       frmCrit.journalno.SetFocus
20990     End If

21000     DoCmd.Hourglass False

21010   End With

EXITP:
21020   Set ctl1 = Nothing
21030   Set ctls = Nothing
21040   Set ctl2 = Nothing
21050   Set ctls2 = Nothing
21060   Set frm1 = Nothing
21070   Set frm2 = Nothing
21080   Exit Sub

ERRH:
21090   DoCmd.Hourglass False
21100   Select Case ERR.Number
        Case Else
21110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21120   End Select
21130   Resume EXITP

End Sub

Public Sub Preview_TA(strFilter01 As String, dblFilterRecs As Double, lngTmp01 As Long, arr_varTmp02 As Variant, lngTmp03 As Long, arr_varTmp04 As Variant, blnIsMaximized As Boolean, frm As Access.Form)

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "Preview_TA"

        Dim strDocName As String
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngE As Long

21210   With frm

21220     If DoReport_TA(dblFilterRecs, lngTmp01, lngTmp03, arr_varTmp02, arr_varTmp04, frm) = True Then  ' ** Function: Above.

21230       lngFrmFlds = lngTmp01
21240       lngE = (lngFrmFlds - 1&)
21250       ReDim arr_varFrmFld(FM_ELEMS, lngE)
21260       For lngX = 0& To (lngFrmFlds - 1&)
21270         For lngY = 0& To FM_ELEMS
21280           arr_varFrmFld(lngY, lngX) = arr_varTmp02(lngY, lngX)
21290         Next
21300       Next
21310       lngFrmFlds_ds = lngTmp03
21320       lngE = (lngFrmFlds_ds - 1&)
21330       ReDim arr_varFrmFld_ds(FM_ELEMS, lngE)
21340       For lngX = 0& To (lngFrmFlds_ds - 1&)
21350         For lngY = 0& To FM_ELEMS
21360           arr_varFrmFld_ds(lngY, lngX) = arr_varTmp04(lngY, lngX)
21370         Next
21380       Next

21390       blnIsMaximized = IsMaximized(frm)  ' ** Module Function: modWindowFunctions.

21400       varTmp00 = .frmTransaction_Audit_Sub.Form.ledger_description_max
21410       If IsNull(varTmp00) = False Then
21420         .ledger_description_max = varTmp00
21430       End If
21440       varTmp00 = .frmTransaction_Audit_Sub.Form.RecurringItem_max
21450       If IsNull(varTmp00) = False Then
21460         .RecurringItem_max = varTmp00
21470       End If
21480       varTmp00 = .frmTransaction_Audit_Sub.Form.ledger_HIDDEN_min
21490       If IsNull(varTmp00) = False Then
21500         .ledger_HIDDEN_min = varTmp00
21510       End If
21520       varTmp00 = .frmTransaction_Audit_Sub.Form.ForExCnt_sum
21530       If IsNull(varTmp00) = False Then
21540         .ForExCnt_sum = varTmp00
21550       End If

21560       strDocName = "rptTransaction_Audit_01"
21570       DoCmd.OpenReport strDocName, acViewPreview, , strFilter01, , frm.Name & "~" & CStr(acViewPreview)

21580       DoCmd.Maximize
21590       DoCmd.RunCommand acCmdFitToWindow
            'DoCmd.RunCommand acCmdZoom100

21600     End If

21610   End With

EXITP:
21620   Exit Sub

ERRH:
21630   frm.Visible = True
21640   Select Case ERR.Number
        Case Else
21650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21660   End Select
21670   Resume EXITP

End Sub

Public Sub Print_TA(strFilter01 As String, dblFilterRecs As Double, lngTmp01 As Long, arr_varTmp02 As Variant, lngTmp03 As Long, arr_varTmp04 As Variant, frm As Access.Form)

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "Print_TA"

        Dim strDocName As String
        Dim lngX As Long, lngY As Long, lngE As Long

21710   If DoReport_TA(dblFilterRecs, lngTmp01, lngTmp03, arr_varTmp02, arr_varTmp04, frm) = True Then  ' ** Function: Above.

21720     lngFrmFlds = lngTmp01
21730     lngE = (lngFrmFlds - 1&)
21740     ReDim arr_varFrmFld(FM_ELEMS, lngE)
21750     For lngX = 0& To (lngFrmFlds - 1&)
21760       For lngY = 0& To FM_ELEMS
21770         arr_varFrmFld(lngY, lngX) = arr_varTmp02(lngY, lngX)
21780       Next
21790     Next
21800     lngFrmFlds_ds = lngTmp03
21810     lngE = (lngFrmFlds_ds - 1&)
21820     ReDim arr_varFrmFld_ds(FM_ELEMS, lngE)
21830     For lngX = 0& To (lngFrmFlds_ds - 1&)
21840       For lngY = 0& To FM_ELEMS
21850         arr_varFrmFld_ds(lngY, lngX) = arr_varTmp04(lngY, lngX)
21860       Next
21870     Next

21880     strDocName = "rptTransaction_Audit_01"
          '##GTR_Ref: rptTransaction_Audit_01
21890     DoCmd.OpenReport strDocName, acViewNormal, , strFilter01, , frm.Name & "~" & CStr(acViewNormal)

21900   End If

EXITP:
21910   Exit Sub

ERRH:
21920   Select Case ERR.Number
        Case Else
21930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21940   End Select
21950   Resume EXITP

End Sub

Public Sub Word_TA(strFilter01 As String, dblFilterRecs As Double, lngTmp01 As Long, arr_varTmp02 As Variant, lngTmp03 As Long, arr_varTmp04 As Variant, frm As Access.Form)

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "Word_TA"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim lngX As Long, lngY As Long, lngE As Long

22010   With frm

22020     If DoReport_TA(dblFilterRecs, lngTmp01, lngTmp03, arr_varTmp02, arr_varTmp04, frm) = True Then  ' ** Function: Above.

22030       lngFrmFlds = lngTmp01
22040       lngE = (lngFrmFlds - 1&)
22050       ReDim arr_varFrmFld(FM_ELEMS, lngE)
22060       For lngX = 0& To (lngFrmFlds - 1&)
22070         For lngY = 0& To FM_ELEMS
22080           arr_varFrmFld(lngY, lngX) = arr_varTmp02(lngY, lngX)
22090         Next
22100       Next
22110       lngFrmFlds_ds = lngTmp03
22120       lngE = (lngFrmFlds_ds - 1&)
22130       ReDim arr_varFrmFld_ds(FM_ELEMS, lngE)
22140       For lngX = 0& To (lngFrmFlds_ds - 1&)
22150         For lngY = 0& To FM_ELEMS
22160           arr_varFrmFld_ds(lngY, lngX) = arr_varTmp04(lngY, lngX)
22170         Next
22180       Next

22190       If IsNull(.UserReportPath) = True Then
22200         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
22210       Else
22220         strRptPath = .UserReportPath
22230       End If
22240       strRptCap = "rptMasterBalance_" & Format(Date, "yyyymmdd")

22250       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

22260       If strRptPathFile <> vbNullString Then
22270         gstrCrtRpt_Ordinal = strFilter01
22280         strRpt = "rptTransaction_Audit_01"
22290         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
22300         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
22310       End If

22320     End If
22330   End With

EXITP:
22340   Exit Sub

ERRH:
22350   Select Case ERR.Number
        Case Else
22360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22370   End Select
22380   Resume EXITP

End Sub

Public Sub Excel_TA(strFilter01 As String, dblFilterRecs As Double, lngTmp01 As Long, arr_varTmp02 As Variant, lngTmp03 As Long, arr_varTmp04 As Variant, frm As Access.Form)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "Excel_TA"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, frm1 As Access.Form
        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim varSQL As Variant, strSortNow1 As String
        Dim blnContinue As Boolean, blnFound As Boolean
        Dim intPos01 As Integer, intPos02 As Integer
        Dim varTmp00 As Variant, varTmp01 As Variant, varTmp02 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

22410   With frm

22420     DoCmd.Hourglass True
22430     DoEvents

22440     If DoReport_TA(dblFilterRecs, lngTmp01, lngTmp03, arr_varTmp02, arr_varTmp04, frm) = True Then  ' ** Function: Above.

22450       lngFrmFlds = lngTmp01
22460       lngE = (lngFrmFlds - 1&)
22470       ReDim arr_varFrmFld(FM_ELEMS, lngE)
22480       For lngX = 0& To (lngFrmFlds - 1&)
22490         For lngY = 0& To FM_ELEMS
22500           arr_varFrmFld(lngY, lngX) = arr_varTmp02(lngY, lngX)
22510         Next
22520       Next
22530       lngFrmFlds_ds = lngTmp03
22540       lngE = (lngFrmFlds_ds - 1&)
22550       ReDim arr_varFrmFld_ds(FM_ELEMS, lngE)
22560       For lngX = 0& To (lngFrmFlds_ds - 1&)
22570         For lngY = 0& To FM_ELEMS
22580           arr_varFrmFld_ds(lngY, lngX) = arr_varTmp04(lngY, lngX)
22590         Next
22600       Next

22610       blnContinue = True

22620       If IsNull(.UserReportPath) = True Then
22630         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
22640       Else
22650         strRptPath = .UserReportPath
22660       End If

            ' ** Only use the one from form, not the datasheet.
22670       Set frm1 = .frmTransaction_Audit_Sub.Form

22680       varSQL = vbNullString
22690       varTmp00 = vbNullString: varTmp01 = vbNullString: varTmp02 = vbNullString

22700       Set dbs = CurrentDb
22710       With dbs
              ' ** qryTransaction_Audit_20 (Ledger, with add'l fields.), For Export.
22720         Set qdf = .QueryDefs("qryTransaction_Audit_21_bak")
22730         varSQL = qdf.SQL
22740         varSQL = Trim(varSQL)
22750         Set qdf = Nothing
22760         .Close
22770       End With  ' ** dbs.
22780       Set dbs = Nothing

22790       If Right(varSQL, 1) = ";" Then varSQL = Left(varSQL, (Len(varSQL) - 1))

22800       intPos01 = InStr(varSQL, " ")  ' ** Strip off the SELECT.
22810       varTmp00 = Trim(Mid(varSQL, (intPos01 + 1)))  ' ** Starts with list of fields.
22820       intPos01 = InStr(varTmp00, "FROM ")
22830       varTmp00 = Trim(Left(varTmp00, (intPos01 - 1)))  ' ** Strip off the FROM clause.
22840       If Right(varTmp00, 2) = vbCrLf Then varTmp00 = Left(varTmp00, (Len(varTmp00) - 2))  ' ** Strip of CrLf.

            ' ****************************************************
            ' ** Array: arr_varFrmFld()
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

            ' ** SELECT ledger.journalno AS [Journal Num], ledger.journaltype AS [Journal Type],
            ' **   ledger.transdate AS [Posting Date], ledger.accountno AS [Account Num], account.shortname AS Name,
            ' **   masterasset.cusip AS CUSIP, masterasset.description AS Asset, ledger.shareface AS [Share/Face],
            ' **   ledger.icash AS [Income Cash], ledger.pcash AS [Principal Cash], ledger.cost AS Cost,
            ' **   IIf(IsNull([ledger].[assetdate])=True,Null,Format([assetdate],'mm/dd/yyyy hh:nn:ss')) AS [Trade Date],
            ' **   IIf(IsNull([ledger].[PurchaseDate])=True,Null,Format([ledger].[PurchaseDate],'mm/dd/yyyy hh:nn:ss')) AS [Original Trade Date],
            ' **   ledger.description AS Comments, ledger.RecurringItem AS [Recurring Item], m_REVCODE.revcode_DESC AS [Inc/Exp Codes],
            ' **   m_REVCODE_TYPE.revcode_TYPE_Description AS [Inc/Exp], TaxCode.taxcode_description AS [Tax Codes],
            ' **   TaxCode_Type.taxcode_type_description AS [Inc/Ded], ledger.journal_USER AS [User],
            ' **   IIf(IsNull([ledger].[posted])=True,Null,Format([ledger].[posted],'mm/dd/yyyy hh:nn:ss')) AS [Date Posted],
            ' **   ledger.ledger_HIDDEN AS Hidden
            ' ** FROM m_REVCODE_TYPE INNER JOIN (account INNER JOIN ((TaxCode_Type INNER JOIN TaxCode ON
            ' **   TaxCode_Type.taxcode_type = TaxCode.taxcode_type) INNER JOIN (m_REVCODE INNER JOIN
            ' **   (ledger LEFT JOIN masterasset ON ledger.assetno = masterasset.assetno) ON
            ' **   m_REVCODE.revcode_ID = ledger.revcode_ID) ON TaxCode.taxcode = ledger.taxcode) ON
            ' **   account.accountno = ledger.accountno) ON m_REVCODE_TYPE.revcode_TYPE = m_REVCODE.revcode_TYPE;

22850       varTmp02 = vbNullString
22860       For lngX = 1& To 1&  ' ** Only use the form, not the datasheet.
22870         Select Case lngX
              Case 1&
22880           For lngY = 1& To lngFrmFlds
22890             If Left(varTmp00, 4) = "IIf(" Then
                    ' ** IIf() fields: Trade Date, Original Trade Date, Date Posted.
22900               intPos02 = InStr(varTmp00, " AS ")
22910               intPos01 = InStr(intPos02, varTmp00, ",")  ' ** Comma after 'AS' fields.
22920               varTmp01 = Trim(Left(varTmp00, intPos02))
22930               If InStr(varTmp01, "assetdate") > 0 Then         ' ** Trade Date.
22940                 varTmp01 = "assetdate"
22950               ElseIf InStr(varTmp01, "PurchaseDate") > 0 Then  ' ** Original Trade Date.
22960                 varTmp01 = "PurchaseDate"
22970               ElseIf InStr(varTmp01, "posted") > 0 Then        ' ** Date Posted.
22980                 varTmp01 = "posted"
22990               Else
                      ' ** Oh, that's enough!
23000               End If
23010             Else
23020               intPos01 = InStr(varTmp00, ",")  ' ** ledger_HIDDEN won't have a comma!
23030               intPos02 = InStr(varTmp00, " AS ")  ' ** All fields are renamed, so have 'AS'.
23040               varTmp01 = Trim(Left(varTmp00, intPos02))  ' ** ledger.journalno
23050               intPos02 = InStr(varTmp01, ".")
23060               varTmp01 = Mid(varTmp01, (intPos02 + 1))  ' ** journalno
23070             End If
23080             blnFound = False
23090             For lngZ = 0& To (lngFrmFlds - 1&)
23100               If varTmp01 = "description" Then
23110                 If Left(varTmp00, 11) = "masterasset" Then
                        ' ** masterasset.description = asset_description
23120                   varTmp01 = "asset_description"
23130                 ElseIf Left(varTmp00, 6) = "ledger" Then
                        ' ** ledger.description = ledger_description
23140                   varTmp01 = "ledger_description"
23150                 Else
23160                   blnContinue = False
23170                 End If
23180               End If
23190               If blnContinue = True Then
23200                 If arr_varFrmFld(FM_FLD_NAM, lngZ) = varTmp01 Then
23210                   blnFound = True
23220                   If arr_varFrmFld(FM_CHK_VAL, lngZ) = True Then
23230                     If intPos01 > 0 Then
                            ' ** Put this field in, from start to comma.
23240                       varTmp02 = varTmp02 & Left(varTmp00, intPos01)
23250                       varTmp02 = varTmp02 & " "  ' ** Add a space.
23260                     Else
                            ' ** Last field, should be ledger_HIDDEN.
23270                       varTmp02 = varTmp02 & varTmp00  ' ** CrLf should have already been stripped off.
23280                     End If
23290                   End If
23300                   Exit For
23310                 End If
23320               End If  ' ** blnContinue.
23330             Next  ' ** lngZ: lngFrmFlds.
23340             If blnFound = False And blnContinue = True Then
23350               blnContinue = False
23360             End If
23370             If blnContinue = False Then
23380               Exit For
23390             End If
23400             If intPos01 > 0 Then
23410               varTmp00 = Trim(Mid(varTmp00, (intPos01 + 1)))  ' ** Strip off this processed field.
23420             End If
23430           Next  ' ** lngY: lngFrmFlds.
23440         Case 2&
                ' ** Not using DataSheet.
23450         End Select
23460       Next  ' ** lngX

23470       If blnContinue = True Then

23480         varTmp02 = Trim(varTmp02)
23490         If Right(varTmp02, 1) = "," Then varTmp02 = Left(varTmp02, (Len(varTmp02) - 1))
23500         varTmp00 = "SELECT " & varTmp02 & vbCrLf & Mid(varSQL, InStr(varSQL, "FROM "))
23510         varTmp00 = Trim(varTmp00)
23520         strSortNow1 = frm1.SortNow_Get  ' ** Form Function: frmTransaction_Audit_Sub.
23530         If strSortNow1 <> vbNullString Then
23540           varTmp00 = varTmp00 & vbCrLf & "ORDER BY " & strSortNow1
23550         End If
23560         intPos01 = InStr(varTmp00, ";")
23570         If intPos01 > 0 Then varTmp00 = Left(varTmp00, (intPos01 - 1))
23580         varSQL = varTmp00
23590         varTmp00 = vbNullString: varTmp01 = vbNullString: varTmp02 = vbNullString

23600         Set dbs = CurrentDb
23610         Set qdf = dbs.QueryDefs("qryTransaction_Audit_30")  ' ** This errors with other syntax!
23620         If strFilter01 <> vbNullString Then
23630           varSQL = varSQL & vbCrLf & "WHERE " & strFilter01 & ";"
                ' ** Fields that need to be qualified:
                ' **   accountno
                ' **   assetno
                ' **   revcode_ID
                ' **   taxcode
23640           intPos01 = InStr(varSQL, "FROM ")
23650           intPos02 = InStr(intPos01, varSQL, "[accountno]")
23660           If intPos02 > 0 Then
23670             varSQL = Left(varSQL, (intPos02 - 1)) & "[ledger]." & Mid(varSQL, intPos02)
23680           End If
23690           intPos02 = InStr(intPos01, varSQL, "[assetno]")
23700           If intPos02 > 0 Then
23710             varSQL = Left(varSQL, (intPos02 - 1)) & "[ledger]." & Mid(varSQL, intPos02)
23720           End If
23730           intPos02 = InStr(intPos01, varSQL, "[revcode_ID]")
23740           If intPos02 > 0 Then
23750             varSQL = Left(varSQL, (intPos02 - 1)) & "[ledger]." & Mid(varSQL, intPos02)
23760           End If
23770           intPos02 = InStr(intPos01, varSQL, "[taxcode]")
23780           If intPos02 > 0 Then
23790             varSQL = Left(varSQL, (intPos02 - 1)) & "[ledger]." & Mid(varSQL, intPos02)
23800           End If
23810           intPos01 = InStr(varSQL, "Hiddenledger")
23820           If intPos01 > 0 Then
                  ' ** , ledger.ledger_HIDDEN AS Hiddenledger.ledger_HIDDEN AS Hiddenledger.ledger_HIDDEN AS Hiddenledger.ledger_HIDDEN AS Hidden
23830             varTmp00 = Left(varSQL, (intPos01 + 6))
23840             intPos01 = InStr(varSQL, "FROM ")
23850             varTmp00 = varTmp00 & vbCrLf & Mid(varSQL, intPos01)
23860             varSQL = varTmp00
23870           End If
23880 On Error Resume Next
23890           qdf.SQL = varSQL
23900           If ERR.Number <> 0 Then
23910             blnContinue = False
23920 On Error GoTo ERRH
23930           Else
23940 On Error GoTo ERRH
23950           End If
23960         Else
23970           intPos01 = InStr(varSQL, "Hiddenledger")
23980           If intPos01 > 0 Then
                  ' ** , ledger.ledger_HIDDEN AS Hiddenledger.ledger_HIDDEN AS Hiddenledger.ledger_HIDDEN AS Hiddenledger.ledger_HIDDEN AS Hidden
23990             varTmp00 = Left(varSQL, (intPos01 + 6))
24000             intPos01 = InStr(varSQL, "FROM ")
24010             varTmp00 = varTmp00 & vbCrLf & Mid(varSQL, intPos01)
24020             varSQL = varTmp00
24030           End If
24040           varSQL = varSQL & ";"
24050           qdf.SQL = varSQL
24060         End If
24070         Set qdf = Nothing

24080       End If  ' ** blnContinue.

24090       dbs.Close

24100       If blnContinue = True Then

24110         strRptCap = "rptTrxAudit_" & Format(Date, "yyyymmdd")

24120         DoCmd.Hourglass False
24130         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

24140         If strRptPathFile <> vbNullString Then
24150           DoCmd.Hourglass True
24160           DoEvents
24170           strQry = "qryTransaction_Audit_30"
24180           DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
24190           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24200         End If

24210       Else
24220         DoCmd.Hourglass False
24230         Beep
24240         MsgBox "A problem occurred preparing your selection for export to Excel.", vbInformation + vbOKOnly, "Export Failed"
24250       End If

24260     End If

24270     DoCmd.Hourglass False

24280   End With

EXITP:
24290   Set frm1 = Nothing
24300   Set qdf = Nothing
24310   Set dbs = Nothing
24320   Exit Sub

ERRH:
24330   DoCmd.Hourglass False
24340   Select Case ERR.Number
        Case Else
24350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24360   End Select
24370   Resume EXITP

End Sub

Public Sub NextFocus(strProc As String, blnNext As Boolean, lngFlds As Long, arr_varFld As Variant, frm As Access.Form)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "NextFocus"

        Dim strCtlName As String
        Dim lngRecsCur As Long
        Dim blnFound As Boolean, blnFound2 As Boolean
        Dim lngX As Long

24410   With frm
24420     If Right(strProc, 8) = "_KeyDown" Then
24430       strCtlName = Left(strProc, (Len(strProc) - 8))
24440       blnFound = False: blnFound2 = False
24450       Select Case blnNext
            Case True
24460         For lngX = 0& To (lngFlds - 1&)
24470           If arr_varFld(F_FNAM, lngX) = strCtlName Then
24480             blnFound = True
24490           ElseIf blnFound = True And arr_varFld(F_VIS, lngX) = True Then
24500             blnFound2 = True
24510             .Controls(arr_varFld(F_FNAM, lngX)).SetFocus
24520             Exit For
24530           End If
24540         Next
24550         If blnFound2 = False Then
                ' ** Start over from the beginning.
24560           For lngX = 0& To (lngFlds - 1&)
24570             If arr_varFld(F_VIS, lngX) = True And arr_varFld(F_FNAM, lngX) <> strCtlName Then
24580               blnFound2 = True
24590               lngRecsCur = .RecCnt  ' ** Form Function: frmTransaction_Audit_Sub.
24600               If .CurrentRecord < lngRecsCur Then
24610                 .MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmTransaction_Audit_Sub.
24620                 .Controls(arr_varFld(F_FNAM, lngX)).SetFocus
24630               Else
24640                 .FocusHolder.SetFocus
24650               End If
24660               Exit For
24670             End If
24680           Next
24690           If blnFound2 = False Then
24700             .FocusHolder.SetFocus
24710             blnFound2 = True
24720           End If
24730         End If
24740       Case False
24750         For lngX = (lngFlds - 1&) To 0& Step -1&
24760           If arr_varFld(F_FNAM, lngX) = strCtlName Then
24770             blnFound = True
24780           ElseIf blnFound = True And arr_varFld(F_VIS, lngX) = True Then
24790             blnFound2 = True
24800             .Controls(arr_varFld(F_FNAM, lngX)).SetFocus
24810             Exit For
24820           End If
24830         Next
24840         If blnFound2 = False Then
                ' ** Start over again from the end.
24850           For lngX = (lngFlds - 1&) To 0& Step -1&
24860             If arr_varFld(F_VIS, lngX) = True And arr_varFld(F_FNAM, lngX) <> strCtlName Then
24870               blnFound2 = True
24880               If .CurrentRecord > 1 Then
24890                 .MoveRec acCmdRecordsGoToPrevious
24900                 .Controls(arr_varFld(F_FNAM, lngX)).SetFocus
24910               Else
24920                 .FocusHolder2.SetFocus
24930               End If
24940               Exit For
24950             End If
24960           Next
24970           If blnFound2 = False Then
24980             .FocusHolder2.SetFocus
24990             blnFound2 = True
25000           End If
25010         End If
25020       End Select
25030     End If
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

Public Sub Print_Chk_TA(strProc As String, intMode As Integer, frm As Access.Form)

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "Print_Chk_TA"

        Dim strCtl As String, strLabel As String
        Dim lngForeColor As Long
        Dim lngRetVal As Long

25110   With frm

25120     Select Case intMode
          Case 1  ' ** frmTransaction_Audit_Sub.

25130       If InStr(strProc, "_AfterUpdate") = 0 Then
25140         strCtl = Left(strProc, (Len(strProc) - Len("_tgl_Click")))
25150         strCtl = strCtl & "_chk"
25160       Else
25170         strCtl = Left(strProc, (Len(strProc) - Len("_AfterUpdate")))
25180       End If
25190       strLabel = Left(strCtl, (Len(strCtl) - 4)) & "_lbl"

25200       Select Case .Controls(strCtl)
            Case True
25210         If .Controls(strLabel).Tag = "Filter" Then
25220           lngForeColor = CLR_BLU
25230         Else
25240           lngForeColor = CLR_DKGRY2
25250         End If
25260         .Controls(strLabel).ForeColor = lngForeColor
25270         .Controls(strLabel & "_line").BorderColor = CLR_DKGRY
25280         Select Case strLabel
              Case "journalno_lbl", "journaltype_lbl", "transdate_lbl", "accountno_lbl", "CheckNum_lbl"  ' ** 2-Label headers.
25290           .Controls(strLabel & "2").ForeColor = lngForeColor
25300         End Select
25310       Case False
25320         If .Controls(strLabel).Tag = "Filter" Then
25330           lngForeColor = CLR_IEI_BLU_DIS
25340         Else
25350           lngForeColor = WIN_CLR_DISF
25360         End If
25370         .Controls(strLabel).ForeColor = lngForeColor
25380         .Controls(strLabel & "_line").BorderColor = CLR_GRY4
25390         Select Case strLabel
              Case "journalno_lbl", "journaltype_lbl", "transdate_lbl", "accountno_lbl", "CheckNum_lbl"  ' ** 2-Label headers.
25400           .Controls(strLabel & "2").ForeColor = lngForeColor
25410         End Select
25420       End Select
25430       lngRetVal = Print_ChkCnt_TA(intMode, frm)  ' ** Function: Below.

25440     Case 2  ' ** frmTransaction_Audit_Sub_ds.

25450       lngRetVal = Print_ChkCnt_TA(intMode, frm)  ' ** Function: Below.

25460     End Select

25470   End With

EXITP:
25480   Exit Sub

ERRH:
25490   Select Case ERR.Number
        Case Else
25500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25510   End Select
25520   Resume EXITP

End Sub

Public Function Print_ChkCnt_TA(intMode As Integer, frm As Access.Form) As Long

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "Print_ChkCnt_TA"

        Dim frm1 As Access.Form, ctl As Access.Control
        Dim lngChksTot As Long
        Dim strTmp01 As String
        Dim lngRetVal As Long

25610   lngRetVal = 0&

25620   With frm

25630     Select Case intMode
          Case 1  ' ** frmTransaction_Audit_Sub.

25640       lngChksTot = 0&
            ' ** Make sure this matches .._Sub_ds.
25650       Set frm1 = Forms(.Parent.Name).Controls(frm.Name & "_ds").Form
25660       For Each ctl In .FormHeader.Controls
25670         With ctl
                ' ** Only count those visible!
25680           If Right(.Name, 4) = "_chk" Then
25690             strTmp01 = Left(.Name, (Len(.Name) - 4))
25700             If frm.Controls(strTmp01).Visible = True Then
25710               lngChksTot = lngChksTot + 1&
25720             End If
25730             Select Case .Value
                  Case True
25740               If frm.Controls(strTmp01).Visible = True Then
25750                 lngRetVal = lngRetVal + 1&
25760               End If
25770               frm1.Controls(.Name) = True
25780             Case False
25790               frm1.Controls(.Name) = False
25800             End Select
25810           End If
25820         End With
25830       Next
25840       If lngRetVal = 0 Then
25850         .Parent.cmdSelect_lbl2.ForeColor = CLR_DKRED
25860         .Parent.cmdSelect_lbl2.FontBold = True
25870       ElseIf lngRetVal < lngChksTot Then
25880         .Parent.cmdSelect_lbl2.ForeColor = CLR_DKBLU
25890         .Parent.cmdSelect_lbl2.FontBold = False
25900       Else
25910         .Parent.cmdSelect_lbl2.ForeColor = CLR_VDKGRY
25920         .Parent.cmdSelect_lbl2.FontBold = False
25930       End If
25940       .Parent.cmdSelect_lbl2.Caption = CStr(lngRetVal) & " of " & CStr(lngChksTot) & " Selected"
25950       DoEvents

25960     Case 2  ' ** frmTransaction_Audit_Sub_ds.

25970       lngChksTot = 0&
25980       For Each ctl In .FormHeader.Controls
25990         With ctl
                ' ** Only count those visible!
26000           If Right(.Name, 4) = "_chk" Then
26010             strTmp01 = Left(.Name, (Len(.Name) - 4))
26020             strTmp01 = frm.SwapNames(strTmp01)  ' ** Form Function: frmTransaction_Audit_Sub_ds.
26030             If frm.Controls(strTmp01).Visible = True Then
26040               lngChksTot = lngChksTot + 1&
26050             End If
26060             If .Value = True Then
26070               If frm.Controls(strTmp01).Visible = True Then
26080                 lngRetVal = lngRetVal + 1&
26090               End If
26100             End If
26110           End If
26120         End With
26130       Next
26140       If lngRetVal = 0 Then
26150         .Parent.cmdSelect_lbl2.ForeColor = CLR_DKRED
26160         .Parent.cmdSelect_lbl2.FontBold = True
26170       ElseIf lngRetVal < lngChksTot Then
26180         .Parent.cmdSelect_lbl2.ForeColor = CLR_DKBLU
26190         .Parent.cmdSelect_lbl2.FontBold = False
26200       Else
26210         .Parent.cmdSelect_lbl2.ForeColor = CLR_VDKGRY
26220         .Parent.cmdSelect_lbl2.FontBold = False
26230       End If
26240       .Parent.cmdSelect_lbl2.Caption = CStr(lngRetVal) & " of " & CStr(lngChksTot) & " Selected"
26250       DoEvents

26260     End Select

26270   End With

EXITP:
26280   Set ctl = Nothing
26290   Set frm1 = Nothing
26300   Print_ChkCnt_TA = lngRetVal
26310   Exit Function

ERRH:
26320   lngRetVal = 0&
26330   Select Case ERR.Number
        Case Else
26340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26350   End Select
26360   Resume EXITP

End Function

Public Sub FilterRecs_Clr_TA(strProc As String, blnOn As Boolean, frm As Access.Form)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Clr_TA"

        Dim strCtl As String, strCtlLabel As String, strCtl2 As String, strCtlLabel2 As String, strCtlLabel3 As String
        Dim strFld As String, strFldLabelA As String, strFldLabelB As String
        Dim strFld2 As String, strFldLabel2 As String, strFld3 As String, strFldLabel3 As String
        Dim intPos01 As Integer
        Dim lngX As Long

26410   strCtl = Left(strProc, (Len(strProc) - Len("_AfterUpdate")))  ' ** PostedDateStart_AfterUpdate -> PostedDateStart
26420   strCtlLabel = strCtl & "_lbl"                                  ' ** PostedDateStart -> PostedDateStart_lbl
26430   strFld = vbNullString: strFldLabelA = vbNullString: strFldLabelB = vbNullString
26440   strFld2 = vbNullString: strFldLabel2 = vbNullString: strFld3 = vbNullString: strFldLabel3 = vbNullString
26450   strCtl2 = vbNullString: strCtlLabel2 = vbNullString: strCtlLabel3 = vbNullString

26460   arr_varFilt = Empty
26470   arr_varFilt = FilterRecs_GetArr(1)  ' ** Function: Above.
26480   If IsEmpty(arr_varFilt) = True Then
26490     FilterRecs_Load  ' ** Procedure: Above.
26500     DoEvents
26510   End If
26520   If IsEmpty(arr_varFilt) = False Then
          'If IsNull(arr_varFilt) = False Then

26530     lngFilts = (UBound(arr_varFilt, 2) + 1&)

          'If IsNothing(frmCrit) = True Then  ' ** modUtilities.
26540     Set frmCrit = Forms("frmTransaction_Audit").frmTransaction_Audit_Sub_Criteria.Form
          'End If

          ' ******************************************************
          ' ** Array: arr_varFilt()
          ' **
          ' **   Field  Element  Name                 Constant
          ' **   =====  =======  ===================  ==========
          ' **     1       0     taf_index            F_NAM
          ' **     2       1     vbdec_name1          F_NAM
          ' **     3       2     vbdec_value1         F_CONST
          ' **     4       3     ctl_name1            F_CTL
          ' **     5       4     ctl_name_lbl1        F_CLBL
          ' **     6       5     fld_ctl_name1        F_FLD
          ' **     7       6     fld_ctl_name_lbl1    F_FLBL
          ' **     8       7     ctl_name2            F_CTL2
          ' **     9       8     ctl_name_lbl2        F_CLBL2
          ' **    10       9     fld_ctl_name2        F_FLD2
          ' **    11      10     fld_ctl_name_lbl2    F_FLBL2
          ' **    12      11     ctl_name_lbl3        F_CLBL3
          ' **    13      12     fld_ctl_name3        F_FLD3
          ' **    14      13     fld_ctl_name_lbl3    F_FLBL3
          ' **
          ' ******************************************************

26550     For lngX = 0& To (lngFilts - 1&)
26560       If (arr_varFilt(F_CTL, lngX) = strCtl) Or (arr_varFilt(F_CTL, lngX) = "cmbJournalType1" And _
                (strCtl = "cmbJournalType2" Or strCtl = "cmbJournalType3")) Then
26570         strFld = arr_varFilt(F_FLD, lngX)
26580         strFldLabelA = arr_varFilt(F_FLBL, lngX)  ' ** Not displayed: [revcode_ID], [assetno].
26590         intPos01 = InStr(strFldLabelA, "~")
26600         If intPos01 > 0 Then
26610           strFldLabelB = Mid(strFldLabelA, (intPos01 + 1))
26620           strFldLabelA = Left(strFldLabelA, (intPos01 - 1))
26630         End If
26640         If arr_varFilt(F_CTL2, lngX) <> vbNullString Then
                ' ** Two criteria controls affect one field.
                ' ** For example: {TransDateStart, TransDateEnd}, {AssetDateStart, AssetDateEnd}, {PostedDateStart, PostedDateEnd}
26650           strCtl2 = arr_varFilt(F_CTL2, lngX)
26660           If arr_varFilt(F_CLBL2, lngX) <> vbNullString Then
26670             strCtlLabel2 = arr_varFilt(F_CLBL2, lngX)
26680           End If
26690           If arr_varFilt(F_CLBL3, lngX) <> vbNullString Then
26700             strCtlLabel3 = arr_varFilt(F_CLBL3, lngX)
26710           End If
26720         ElseIf arr_varFilt(F_CLBL2, lngX) <> vbNullString Then
26730           strCtlLabel2 = arr_varFilt(F_CLBL2, lngX)
26740         End If
26750         If arr_varFilt(F_FLD2, lngX) <> vbNullString Then
                ' ** Two fields affected by one criteria control.
                ' ** For example: {accountno, shortname}, {assetno, asset_description}, {assetdate, PurchaseDate}, {revcode_ID, revcode_DESC}
26760           strFld2 = arr_varFilt(F_FLD2, lngX)
26770           If arr_varFilt(F_FLBL2, lngX) <> vbNullString Then
26780             strFldLabel2 = arr_varFilt(F_FLBL2, lngX)
26790           End If
26800           If arr_varFilt(F_FLD3, lngX) <> vbNullString Then
                  ' ** For example: {assetno, asset_description, cusip}
26810             strFld3 = arr_varFilt(F_FLD3, lngX)
26820             If arr_varFilt(F_FLBL3, lngX) <> vbNullString Then
26830               strFldLabel3 = arr_varFilt(F_FLBL3, lngX)
26840             End If
26850           End If
26860         End If
26870         Exit For
26880       End If
26890     Next

26900     Set frmCrit = Forms("frmTransaction_Audit").frmTransaction_Audit_Sub_Criteria.Form

26910     If strFld <> vbNullString Then
26920       With frm
26930         Select Case blnOn
              Case True
                ' ** Set parent label colors.
26940           frmCrit.Controls(strCtlLabel).ForeColor = CLR_BLU
26950           If strCtlLabel2 <> vbNullString Then
26960             frmCrit.Controls(strCtlLabel2).ForeColor = CLR_BLU
26970           End If
26980           If strCtlLabel3 <> vbNullString Then
26990             frmCrit.Controls(strCtlLabel3).ForeColor = CLR_BLU
27000           End If
                ' ** Set field label colors.
27010           If strFldLabelA <> vbNullString Then
27020             .Controls(strFldLabelA).ForeColor = CLR_BLU
27030             .Controls(strFldLabelA).Tag = "Filter"
27040           End If
27050           If strFldLabelB <> vbNullString Then
27060             .Controls(strFldLabelB).ForeColor = CLR_BLU
27070             .Controls(strFldLabelB).Tag = "Filter"
27080           End If
27090           If strFldLabel2 <> vbNullString Then
27100             .Controls(strFldLabel2).ForeColor = CLR_BLU
27110             .Controls(strFldLabel2).Tag = "Filter"
27120           End If
27130           If strFldLabel3 <> vbNullString Then
27140             .Controls(strFldLabel3).ForeColor = CLR_BLU
27150             .Controls(strFldLabel3).Tag = "Filter"
27160           End If
27170         Case False
                ' ** Set parent label colors.
27180           frmCrit.Controls(strCtlLabel).ForeColor = CLR_VDKGRY
27190           If strCtlLabel2 <> vbNullString Then
27200             frmCrit.Controls(strCtlLabel2).ForeColor = CLR_VDKGRY
27210           End If
27220           If strCtlLabel3 <> vbNullString Then
27230             frmCrit.Controls(strCtlLabel3).ForeColor = CLR_VDKGRY
27240           End If
                ' ** Set field label colors.
27250           If strCtl <> "cmbJournaltype2" And strCtl <> "cmbJournaltype3" Then
27260             If strFldLabelA <> vbNullString Then
27270               .Controls(strFldLabelA).ForeColor = CLR_DKGRY2
27280               .Controls(strFldLabelA).Tag = vbNullString
27290             End If
27300             If strFldLabelB <> vbNullString Then
27310               .Controls(strFldLabelB).ForeColor = CLR_DKGRY2
27320               .Controls(strFldLabelB).Tag = vbNullString
27330             End If
27340             If strFldLabel2 <> vbNullString Then
27350               .Controls(strFldLabel2).ForeColor = CLR_DKGRY2
27360               .Controls(strFldLabel2).Tag = vbNullString
27370             End If
27380             If strFldLabel3 <> vbNullString Then
27390               .Controls(strFldLabel3).ForeColor = CLR_DKGRY2
27400               .Controls(strFldLabel3).Tag = vbNullString
27410             End If
27420           End If
27430         End Select
27440       End With
27450     End If

27460   End If

EXITP:
27470   Exit Sub

ERRH:
27480   Select Case ERR.Number
        Case Else
27490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27500   End Select
27510   Resume EXITP

End Sub

Public Sub RstAllSet_TA(rstAll1 As DAO.Recordset, rstAll2 As DAO.Recordset, strQryName As String, Optional varMode As Variant)

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "RstAllSet_TA"

        Dim qdfAll As DAO.QueryDef
        Dim intMode As Integer

27610   Select Case IsMissing(varMode)
        Case True
27620     intMode = 1
27630   Case False
27640     intMode = varMode
27650   End Select

27660   Select Case intMode
        Case 1
27670     Set qdfAll = CurrentDb.QueryDefs(strQryName)
27680     Set rstAll1 = qdfAll.OpenRecordset
27690     If rstAll1.BOF = True And rstAll1.EOF = True Then
            ' ** Unlikely.
27700     Else
27710       rstAll1.MoveFirst
27720     End If
27730   Case 2
27740     Set qdfAll = CurrentDb.QueryDefs(strQryName)
27750     Set rstAll2 = qdfAll.OpenRecordset
27760     If rstAll2.BOF = True And rstAll2.EOF = True Then
            ' ** Unlikely.
27770     Else
27780       rstAll2.MoveFirst
27790     End If
27800   End Select

EXITP:
27810   Set qdfAll = Nothing
27820   Exit Sub

ERRH:
27830   Select Case ERR.Number
        Case Else
27840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27850   End Select
27860   Resume EXITP

End Sub

Public Sub PostedDateStartAfterUpdate_TA(strFilter01 As String, strFilter02 As String, dblFilterRecs As Double, rstAll1 As DAO.Recordset, frmCall As Access.Form)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "PostedDateStartAfterUpdate_TA"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

27910   With frmCall
27920     DoCmd.Hourglass True
27930     DoEvents
27940     Set frm = .Parent
27950     If IsNull(.PostedDateStart) = False Then
27960       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
27970         strFilter01 = POSTED_START & .PostedDateStart & "#"
27980       Else
              ' ** There are clauses present.

27990         .FilterRec_GetArr  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.

28000         intPos01 = InStr(strFilter01, POSTED_START)
28010         If intPos01 = 0& Then
                ' ** This clause isn't present.
28020           intPos03 = 0&: intPos04 = 0&
28030           For lngX = (lngFilts - 1&) To 0& Step -1&
28040             If arr_varFilt(F_CONST, lngX) = POSTED_START Then
28050               intPos04 = -1
28060             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
28070               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
28080               If intPos03 > 0 Then
28090                 intPos04 = 0&
28100                 Exit For
28110               End If
28120             End If
28130           Next
28140           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
28150             strFilter01 = POSTED_START & .PostedDateStart & "#" & ANDF & strFilter01
28160           Else
                  ' ** There's a clause before this one.
28170             intPos02 = InStr(intPos03, strFilter01, ANDF)
28180             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
28190               strFilter01 = strFilter01 & ANDF & POSTED_START & .PostedDateStart & "#"
28200             Else
                    ' ** Add this clause to the middle of the filter.
28210               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & POSTED_START & .PostedDateStart & "#" & Mid(strFilter01, intPos02)
28220             End If
28230           End If
28240         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
28250           intPos02 = InStr((intPos01 + Len(POSTED_START) + 1), strFilter01, "#")  ' ** Find the closing paren.
28260           If intPos02 > 0 Then
28270             strFilter01 = Left(strFilter01, (intPos01 - 1)) & POSTED_START & .PostedDateStart & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
28280           Else
28290             strFilter01 = Left(strFilter01, (intPos01 - 1)) & POSTED_START & .PostedDateStart & "#"
28300           End If
28310         End If
28320       End If
28330       strFilter02 = strFilter01
28340       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
28350       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
28360       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
28370       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28380       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
28390       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28400     Else
28410       strFilter02 = strFilter01
28420       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
28430       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem POSTED_START  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
28440       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
28450       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
28460       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28470       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem POSTED_START  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28480       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
28490       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
28500     End If
28510     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          '.PostedDateEnd.SetFocus
28520     DoCmd.Hourglass False
28530     DoEvents
28540   End With

EXITP:
28550   Set frm = Nothing
28560   Exit Sub

ERRH:
28570   DoCmd.Hourglass False
28580   Select Case ERR.Number
        Case Else
28590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28600   End Select
28610   Resume EXITP

End Sub

Public Sub PostedDateEndAfterUpdate_TA(strFilter01 As String, strFilter02 As String, dblFilterRecs As Double, rstAll1 As DAO.Recordset, frmCall As Access.Form)

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "PostedDateEndAfterUpdate_TA"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

28710   With frmCall
28720     DoCmd.Hourglass True
28730     DoEvents
28740     Set frm = .Parent
28750     If IsNull(.PostedDateEnd) = False Then
28760       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
28770         strFilter01 = POSTED_END & .PostedDateEnd & "#"
28780       Else
              ' ** There are clauses present.

28790         .FilterRec_GetArr  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.

28800         intPos01 = InStr(strFilter01, POSTED_END)
28810         If intPos01 = 0& Then
                ' ** This clause isn't present.
28820           intPos03 = 0&: intPos04 = 0&
28830           For lngX = (lngFilts - 1&) To 0& Step -1&
28840             If arr_varFilt(F_CONST, lngX) = POSTED_END Then
28850               intPos04 = -1
28860             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
28870               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
28880               If intPos03 > 0 Then
28890                 intPos04 = 0&
28900                 Exit For
28910               End If
28920             End If
28930           Next
28940           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
28950             strFilter01 = POSTED_END & .PostedDateEnd & "#" & ANDF & strFilter01
28960           Else
                  ' ** There's a clause before this one.
28970             intPos02 = InStr(intPos03, strFilter01, ANDF)
28980             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
28990               strFilter01 = strFilter01 & ANDF & POSTED_END & .PostedDateEnd & "#"
29000             Else
                    ' ** Add this clause to the middle of the filter.
29010               strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & POSTED_END & .PostedDateEnd & "#" & Mid(strFilter01, intPos02)
29020             End If
29030           End If
29040         Else
                ' ** Replace this clause, whether or not it's the last one.
                ' ** Note: If intPos01 = 1, then the Left() function is OK with
                ' ** returning the left 0 characters (as long as it doesn't go below 0).
29050           intPos02 = InStr((intPos01 + Len(POSTED_END) + 1), strFilter01, "#")  ' ** Find the closing paren.
29060           If intPos02 > 0 Then
29070             strFilter01 = Left(strFilter01, (intPos01 - 1)) & POSTED_END & .PostedDateEnd & Mid(strFilter01, intPos02)
                  ' ** Left() ends with ' And ', and there's an ' And ' right after the closing paren.
29080           Else
29090             strFilter01 = Left(strFilter01, (intPos01 - 1)) & POSTED_END & .PostedDateEnd & "#"
29100           End If
29110         End If
29120       End If
29130       strFilter02 = strFilter01
29140       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
29150       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
29160       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
29170       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
29180       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
29190       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
29200     Else
29210       strFilter02 = strFilter01
29220       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
29230       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem POSTED_END  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
29240       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
29250       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
29260       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
29270       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem POSTED_END  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
29280       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
29290       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
29300     End If
29310     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'If .opgHidden.Enabled = True Then
          '  .opgHidden.SetFocus
          'End If
29320     DoCmd.Hourglass False
29330     DoEvents
29340   End With

EXITP:
29350   Set frm = Nothing
29360   Exit Sub

ERRH:
29370   DoCmd.Hourglass False
29380   Select Case ERR.Number
        Case Else
29390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29400   End Select
29410   Resume EXITP

End Sub

Public Sub opgHiddenAfterUpdate_TA(strFilter01 As String, strFilter02 As String, dblFilterRecs As Double, rstAll1 As DAO.Recordset, frmCall As Access.Form)

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgHiddenAfterUpdate_TA"

        Dim frm As Access.Form
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim lngX As Long

29510   With frmCall
29520     DoCmd.Hourglass True
29530     DoEvents
29540     Set frm = .Parent
29550     Select Case .opgHidden
          Case .opgHidden_optInclude.OptionValue
29560       .opgHidden_optInclude_lbl.FontBold = True
29570       .opgHidden_optExclude_lbl.FontBold = False
29580       .opgHidden_optOnly_lbl.FontBold = False
29590     Case .opgHidden_optExclude.OptionValue
29600       .opgHidden_optInclude_lbl.FontBold = False
29610       .opgHidden_optExclude_lbl.FontBold = True
29620       .opgHidden_optOnly_lbl.FontBold = False
29630     Case .opgHidden_optOnly.OptionValue
29640       .opgHidden_optInclude_lbl.FontBold = False
29650       .opgHidden_optExclude_lbl.FontBold = False
29660       .opgHidden_optOnly_lbl.FontBold = True
29670     End Select
29680     If .opgHidden <> .opgHidden_optInclude.OptionValue Then
29690       If strFilter01 = vbNullString Or strFilter02 = vbNullString Then
              ' ** This will be the only clause.
29700         Select Case .opgHidden
              Case .opgHidden_optOnly.OptionValue
29710           strFilter01 = HIDDEN_TRX1
29720         Case .opgHidden_optExclude.OptionValue
29730           strFilter01 = HIDDEN_TRX2
29740         End Select
29750       Else
              ' ** There are clauses present.

29760         .FilterRec_GetArr  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.

29770         intPos01 = InStr(strFilter01, HIDDEN_TRX1)
29780         intPos02 = InStr(strFilter01, HIDDEN_TRX2)
29790         If intPos01 = 0& And intPos02 = 0& Then
                ' ** This clause isn't present.
29800           intPos03 = 0&: intPos04 = 0&
29810           For lngX = (lngFilts - 1&) To 0& Step -1&
29820             If arr_varFilt(F_CONST, lngX) = HIDDEN_TRX1 Then
29830               intPos04 = -1
29840             ElseIf intPos04 = -1 Then
                    ' ** Look for the next previous clause present in strFilter01.
29850               intPos03 = InStr(strFilter01, arr_varFilt(F_CONST, lngX))
29860               If intPos03 > 0 Then
29870                 intPos04 = 0&
29880                 Exit For
29890               End If
29900             End If
29910           Next
29920           If intPos03 = 0& Then
                  ' ** Add this clause at the start of the filter.
29930             Select Case .opgHidden
                  Case .opgHidden_optOnly.OptionValue
29940               strFilter01 = HIDDEN_TRX1 & ANDF & strFilter01
29950             Case .opgHidden_optExclude.OptionValue
29960               strFilter01 = HIDDEN_TRX2 & ANDF & strFilter01
29970             End Select
29980           Else
                  ' ** There's a clause before this one.
29990             intPos02 = InStr(intPos03, strFilter01, ANDF)
30000             If intPos02 = 0 Then
                    ' ** Add this clause to the end of the filter.
30010               Select Case .opgHidden
                    Case .opgHidden_optOnly.OptionValue
30020                 strFilter01 = strFilter01 & ANDF & HIDDEN_TRX1
30030               Case .opgHidden_optExclude.OptionValue
30040                 strFilter01 = strFilter01 & ANDF & HIDDEN_TRX2
30050               End Select
30060             Else
                    ' ** Add this clause to the middle of the filter.
30070               Select Case .opgHidden
                    Case .opgHidden_optOnly.OptionValue
30080                 strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & HIDDEN_TRX1 & Mid(strFilter01, intPos02)
30090               Case .opgHidden_optExclude.OptionValue
30100                 strFilter01 = Left(strFilter01, (intPos02 - 1)) & ANDF & HIDDEN_TRX2 & Mid(strFilter01, intPos02)
30110               End Select
30120             End If
30130           End If
30140         Else
30150           If intPos01 > 0 Then
                  ' ** strFilter01 has HIDDEN_TRX1.
30160             Select Case .opgHidden
                  Case .opgHidden_optOnly.OptionValue
                    ' ** Do nothing.
30170             Case .opgHidden_optExclude.OptionValue
30180               strFilter01 = StringReplace(strFilter01, HIDDEN_TRX1, HIDDEN_TRX2)  ' ** Module Function: modStringFuncs.
30190             End Select
30200           ElseIf intPos02 > 0 Then
                  ' ** strFilter01 has HIDDEN_TRX2.
30210             Select Case .opgHidden
                  Case .opgHidden_optOnly.OptionValue
30220               strFilter01 = StringReplace(strFilter01, HIDDEN_TRX2, HIDDEN_TRX1)  ' ** Module Function: modStringFuncs.
30230             Case .opgHidden_optExclude.OptionValue
                    ' ** Do Nothing.
30240             End Select
30250           End If
30260         End If
30270       End If
30280       strFilter02 = strFilter01
30290       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
30300       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
30310       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
30320       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30330       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
30340       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, True  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30350     Else
30360       strFilter02 = strFilter01
30370       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub.
30380       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Rem HIDDEN_TRX1  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
30390       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub.
30400       frm.frmTransaction_Audit_Sub.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub.
30410       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Set strFilter02, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30420       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Rem HIDDEN_TRX1  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30430       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Cnt rstAll1  ' ** Form Function: frmTransaction_Audit_Sub_ds.
30440       frm.frmTransaction_Audit_Sub_ds.Form.FilterRecs_Clr THIS_PROC, False  ' ** ' ** Form Procedure: frmTransaction_Audit_Sub_ds.
30450     End If
30460     frm.FilterRecs_Set strFilter01, dblFilterRecs  ' ** Forms Procedure: frmTransaction_Audit.
          'DoCmd.SelectObject acForm, .Parent.Name, False
          'Select Case frm.opgView
          'Case frm.opgView_optForm.OptionValue
          '  frm.frmTransaction_Audit_Sub.SetFocus
          'Case frm.opgView_optDatasheet.OptionValue
          '  frm.frmTransaction_Audit_Sub_ds.SetFocus
          'End Select
30470     DoCmd.Hourglass False
30480     DoEvents
30490   End With

EXITP:
30500   Set frm = Nothing
30510   Exit Sub

ERRH:
30520   DoCmd.Hourglass False
30530   Select Case ERR.Number
        Case Else
30540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30550   End Select
30560   Resume EXITP

End Sub

Public Sub ShowFields_Sub_TA(strProc As String, intMode As Integer, lngFields As Long, arr_varField As Variant, lngFldSep As Long, blnFromTgl As Boolean, lngSortLbl_Width As Long, frm As Access.Form)
' ** All the original calls still go to the subform,
' ** which then calls this with the additional parameters.

30600 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowFields_Sub_TA"

        Dim ctl1 As Access.Control, ctl2 As Access.Control
        Dim strCtlName As String, strFieldName As String, strLabel As String
        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim lngSortLbl_Adj As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, blnTmp03 As Boolean
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

30610   With frm

30620     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
30630       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
30640     End If

30650     Select Case intMode
          Case 1  ' ** Load field array.

            ' ** This will pass the array back to the subform via lngFields and arr_varField.
            ' ** When this proc is called for the other modes, it will ignore the fed vars,
            ' ** and just work with the vars that were defined here.

30660       lngFlds = 0&
30670       ReDim arr_varFld(FLD_ELEMS, 0)

30680       If lngChks = 0 Or IsEmpty(arr_varChk) = True Then
30690         arr_varChk = ChkArray_Get ' ** Function: Above.
30700         If IsEmpty(arr_varChk) = True Then
30710           PrintTgls_Load frm ' ** Procedure: Above.
                ' ** This includes a ChkArray_Set(), sending it back to here.
30720         Else
30730           lngChks = UBound(arr_varChk, 2)
30740           If lngChks > 0& Then
30750             lngChks = lngChks + 1&
30760           End If
30770         End If
              ' ** Either PrintTgls_Load() sent the array back here via ChkArray_Set(),
              ' ** or we picked it up from there via ChkArray_Get().
30780         If lngChks = 0& Then
30790           Stop
30800         End If
30810       End If

30820       For lngX = 0& To (lngChks - 1&)
30830         strCtlName = vbNullString: strLabel = vbNullString: lngSortLbl_Adj = 0&
30840         lngFlds = lngFlds + 1&
30850         lngE = lngFlds - 1&
30860         ReDim Preserve arr_varFld(FLD_ELEMS, lngE)
30870         Select Case arr_varChk(C_FNAM, lngX)
              Case "journalno"
30880           strCtlName = "ckgFlds_chkJournalno"
30890           strLabel = "journalno_lbl2"
30900           lngSortLbl_Adj = 6&
30910         Case "journaltype"
30920           strCtlName = "ckgFlds_chkJournalType"
30930           strLabel = "journaltype_lbl2"
30940         Case "transdate"
30950           strCtlName = "ckgFlds_chkTransDate"
30960           strLabel = "transdate_lbl2"
30970           lngSortLbl_Adj = 2&
30980         Case "accountno"
30990           strCtlName = "ckgFlds_chkAccountNo"
31000           strLabel = "accountno_lbl2"
31010         Case "shortname"
31020           strCtlName = "ckgFlds_chkShortName"
31030         Case "cusip"
31040           strCtlName = "ckgFlds_chkCusip"
31050         Case "asset_description"
31060           strCtlName = "ckgFlds_chkAssetDescription"
31070         Case "shareface"
31080           strCtlName = "ckgFlds_chkShareFace"
31090           lngSortLbl_Adj = 8&
31100         Case "icash"
31110           strCtlName = "ckgFlds_chkICash"
31120           lngSortLbl_Adj = 1&
31130         Case "pcash"
31140           strCtlName = "ckgFlds_chkPCash"
31150           lngSortLbl_Adj = 4&
31160         Case "cost"
31170           strCtlName = "ckgFlds_chkCost"
31180         Case "curr_id"
31190           strCtlName = "ckgFlds_chkCurrID"
31200           lngSortLbl_Adj = 6&
31210         Case "assetdate"
31220           strCtlName = "ckgFlds_chkAssetDate"
31230         Case "PurchaseDate"
31240           strCtlName = "ckgFlds_chkPurchaseDate"
31250           lngSortLbl_Adj = 7&
31260         Case "ledger_description"
31270           strCtlName = "ckgFlds_chkLedgerDescription"
31280         Case "RecurringItem"
31290           strCtlName = "ckgFlds_chkRecurringItem"
31300         Case "revcode_DESC"
31310           strCtlName = "ckgFlds_chkRevCodeDesc"
31320         Case "revcode_TYPE_Description"
31330           strCtlName = "ckgFlds_chkRevCodeTypeDescription"
31340           lngSortLbl_Adj = 3&
31350         Case "taxcode_description"
31360           strCtlName = "ckgFlds_chkTaxCodeDescription"
31370         Case "taxcode_type_description"
31380           strCtlName = "ckgFlds_chkTaxCodeTypeDescription"
31390         Case "Location_Name"
31400           strCtlName = "ckgFlds_chkLocationName"
31410         Case "CheckNum"
31420           strCtlName = "ckgFlds_chkCheckNum"
31430           strLabel = "CheckNum_lbl2"
31440           lngSortLbl_Adj = 3&
31450         Case "journal_USER"
31460           strCtlName = "ckgFlds_chkJournalUser"
31470         Case "posted"
31480           strCtlName = "ckgFlds_chkPosted"
31490         Case "ledger_HIDDEN"
31500           strCtlName = "ckgFlds_chkLedgerHidden"
31510           lngSortLbl_Adj = 5&
31520         End Select
31530         arr_varFld(F_CNAM, lngE) = strCtlName
31540         arr_varFld(F_FNAM, lngE) = arr_varChk(C_FNAM, lngX)
31550         arr_varFld(F_LFT, lngE) = .Controls(arr_varChk(C_FNAM, lngX)).Left
31560         arr_varFld(F_WDT, lngE) = .Controls(arr_varChk(C_FNAM, lngX)).Width
31570         arr_varFld(F_LBL1, lngE) = arr_varChk(C_FNAM, lngX) & "_lbl"
31580         If strLabel <> vbNullString Then
31590           arr_varFld(F_LBL2, lngE) = strLabel
31600         Else
31610           arr_varFld(F_LBL2, lngE) = Null
31620         End If
31630         arr_varFld(F_LBL_LFT, lngE) = .Controls(arr_varFld(F_LBL1, lngE)).Left
31640         arr_varFld(F_LBL_WDT, lngE) = .Controls(arr_varFld(F_LBL1, lngE)).Width
31650         arr_varFld(F_LIN, lngE) = arr_varChk(C_FNAM, lngX) & "_lbl_line"
31660         arr_varFld(F_LIN_LFT, lngE) = .Controls(arr_varFld(F_LIN, lngE)).Left
31670         arr_varFld(F_LIN_WDT, lngE) = .Controls(arr_varFld(F_LIN, lngE)).Width
31680         arr_varFld(F_SRT_ADJ, lngE) = lngSortLbl_Adj
31690         arr_varFld(F_VIS, lngE) = CBool(.Parent.Controls(strCtlName))
31700         arr_varFld(F_CHK_ELEM, lngE) = lngX
31710       Next  ' ** lngX.

            ' ** These will have to be dealt with case-by-case.
            'icash_str
            'icash_str_bg
            'icash_box
            'pcash_str
            'pcash_str_bg
            'pcash_box
            'cost_str
            'cost_str_bg
            'cost_box
            'curr_id_box
            'curr_id_bg

            ' ** Here's where it passes them back to the subform.
31720       lngFields = lngFlds
31730       arr_varField = arr_varFld

            'For lngX = 0& To (lngFlds - 1&)
            '  Debug.Print "'" & arr_varFld(F_FNAM, lngX)
            'Next

31740     Case 2  ' ** Show field.

31750       strCtlName = Left(strProc, (CharPos(strProc, 2, "_") - 1))  ' ** Module Function: modStringFuncs.

            'NOW, DO FIELDS APPEAR IN THE ORDER THEY'RE
            'CHOSEN, OR ALWAYS IN THE STARTING ORDER?
            'HOW ABOUT WE START WITH OPENING ORDER,
            'AND ADD ANY-ORDER OPTION LATER!

31760       Do Until lngFlds > 0&
31770         If lngFlds = 0& Then
31780           .ShowFields_Sub strProc, 1  ' ** Form Procedure: frmTransaction_Audit_Sub.  ' ** Recursive.
31790           DoEvents
31800         End If
31810       Loop

31820       lngZ = -1&
31830       For lngX = 0& To (lngFlds - 1&)
31840         If arr_varFld(F_CNAM, lngX) = strCtlName Then
31850           lngZ = lngX
31860           Exit For
31870         End If
31880       Next  ' ** lngX

31890       strFieldName = arr_varFld(F_FNAM, lngZ)
31900       .Controls(arr_varFld(F_FNAM, lngZ)).Visible = True
31910       .Controls(arr_varFld(F_LBL1, lngZ)).Visible = True
31920       If IsNull(arr_varFld(F_LBL2, lngZ)) = False Then
31930         .Controls(arr_varFld(F_LBL2, lngZ)).Visible = True
31940       End If
31950       .Controls(arr_varFld(F_LIN, lngZ)).Visible = True
31960       arr_varFld(F_VIS, lngZ) = CBool(True)

            ' ** Special cases.
31970       Select Case strFieldName
            Case "icash"
31980         .icash_box.Visible = True
31990         .icash_str_bg.Visible = True
32000         .icash_str.Visible = True
32010       Case "pcash"
32020         .pcash_box.Visible = True
32030         .pcash_str_bg.Visible = True
32040         .pcash_str.Visible = True
32050       Case "cost"
32060         .cost_box.Visible = True
32070         .cost_str_bg.Visible = True
32080         .cost_str.Visible = True
32090       Case "curr_id"
32100         .curr_id_box.Visible = True
32110         .Curr_id_bg.Visible = True
32120       Case Else
              ' ** Nothing else.
32130       End Select

            ' ** It may not reappear in the same place that it left.
32140       lngTmp01 = .journalno.Left
32150       For lngX = 0& To (lngFlds - 1&)
32160         If arr_varFld(F_VIS, lngX) = True Then
32170           If arr_varFld(F_FNAM, lngX) = strFieldName Then
32180             If strFieldName = "ledger_HIDDEN" Then
32190               lngTmp02 = (arr_varFld(F_LFT, lngX) - arr_varFld(F_LIN_LFT, lngX))
32200               If .Controls(arr_varFld(F_FNAM, lngX)).Left <> (lngTmp01 + lngTmp02) Then
32210                 .Controls(arr_varFld(F_FNAM, lngX)).Left = (lngTmp01 + lngTmp02)
32220                 lngTmp02 = arr_varFld(F_LIN_LFT, lngX) - arr_varFld(F_LBL_LFT, lngX)
32230                 .Controls(arr_varFld(F_LBL1, lngX)).Left = (lngTmp01 - lngTmp02)
32240                 .Controls(arr_varFld(F_LIN, lngX)).Left = lngTmp01
32250                 lngE = arr_varFld(F_CHK_ELEM, lngX)
32260                 For lngY = C_CMD To C_ONDIS
32270                   .Controls(arr_varChk(lngY, lngE)).Left = lngTmp01
32280                 Next  ' ** lngY.
32290               End If
32300             Else
32310               If .Controls(arr_varFld(F_FNAM, lngX)).Left <> lngTmp01 Then
32320                 .Controls(arr_varFld(F_FNAM, lngX)).Left = lngTmp01
32330                 If arr_varFld(F_LBL_LFT, lngX) <> arr_varFld(F_LFT, lngX) Then
                        ' ** An offset label.
32340                   lngTmp02 = (arr_varFld(F_LIN_LFT, lngX) - arr_varFld(F_LBL_LFT, lngX))
32350                   .Controls(arr_varFld(F_LBL1, lngX)).Left = (lngTmp01 - lngTmp02)
32360                 Else
32370                   .Controls(arr_varFld(F_LBL1, lngX)).Left = lngTmp01
32380                 End If
32390                 If IsNull(arr_varFld(F_LBL2, lngX)) = False Then
32400                   .Controls(arr_varFld(F_LBL2, lngX)).Left = .Controls(arr_varFld(F_LBL1, lngX)).Left
32410                 End If
32420                 .Controls(arr_varFld(F_LIN, lngX)).Left = lngTmp01
32430                 lngE = arr_varFld(F_CHK_ELEM, lngX)
32440                 For lngY = C_CMD To C_ONDIS
32450                   .Controls(arr_varChk(lngY, lngE)).Left = lngTmp01
32460                 Next  ' ** lngY.
32470               End If
                    ' ** Special cases.
32480               Select Case arr_varFld(F_FNAM, lngX)
                    Case "icash"
32490                 .icash_box.Left = .ICash.Left
32500                 .icash_str_bg.Left = (.ICash.Left - lngTpp)
32510                 .icash_str.Left = .ICash.Left
32520               Case "pcash"
32530                 .pcash_box.Left = .PCash.Left
32540                 .pcash_str_bg.Left = (.PCash.Left - lngTpp)
32550                 .pcash_str.Left = .PCash.Left
32560               Case "cost"
32570                 .cost_box.Left = .Cost.Left
32580                 .cost_str_bg.Left = (.Cost.Left - lngTpp)
32590                 .cost_str.Left = .Cost.Left
32600               Case "curr_id"
32610                 .curr_id_box.Left = .curr_id.Left
32620                 .Curr_id_bg.Left = (.curr_id.Left - lngTpp)
32630               Case Else
                      ' ** Nothing else.
32640               End Select
32650             End If
32660             Exit For
32670           Else
32680             lngTmp01 = lngTmp01 + arr_varFld(F_WDT, lngX) + lngFldSep
32690           End If
32700         End If
32710       Next  ' ** lngX.

            ' ** This comes in alternating True/False!
32720       lngE = arr_varFld(F_CHK_ELEM, lngZ)
32730       .Controls(arr_varChk(C_CMD, lngE)).Visible = True
32740       blnTmp03 = .Controls(arr_varFld(F_FNAM, lngZ) & "_chk")
32750       PrintTgls_Click arr_varFld(F_FNAM, lngZ) & "_tgl_Click", blnFromTgl, lngChks, arr_varChk, frm  ' ** Procedure: Above.
32760       DoEvents
32770       If .Controls(arr_varFld(F_FNAM, lngZ) & "_chk") <> blnTmp03 Then
32780         PrintTgls_Click arr_varFld(F_FNAM, lngZ) & "_tgl_Click", blnFromTgl, lngChks, arr_varChk, frm  ' ** Procedure: Above.
32790       End If

            ' ** Move everyone aside.
32800       lngTmp01 = .journalno.Left
32810       blnSortHere = False: blnResort = False
32820       For lngX = 0& To (lngFlds - 1&)
32830         blnSortHere = False
32840         If arr_varFld(F_FNAM, lngX) = strFieldName Then
                ' ** This is the one just added.
32850           If strFieldName = "ledger_HIDDEN" Then
32860             lngTmp01 = (lngTmp01 + arr_varFld(F_LIN_WDT, lngX) + lngFldSep)
32870           Else
32880             lngTmp01 = (lngTmp01 + arr_varFld(F_WDT, lngX) + lngFldSep)
32890           End If
32900         ElseIf arr_varFld(F_VIS, lngX) = True Then
32910           If .Sort_line.Left = .Controls(arr_varFld(F_LIN, lngX)).Left Then blnSortHere = True
32920           If arr_varFld(F_FNAM, lngX) = "ledger_HIDDEN" Then
32930             lngTmp02 = (arr_varFld(F_LFT, lngX) - arr_varFld(F_LIN_LFT, lngX))
32940             If .Controls(arr_varFld(F_FNAM, lngX)).Left <> (lngTmp01 + lngTmp02) Then
32950               .Controls(arr_varFld(F_FNAM, lngX)).Left = (lngTmp01 + lngTmp02)
32960               lngTmp02 = (arr_varFld(F_LIN_LFT, lngX) - arr_varFld(F_LBL_LFT, lngX))
32970               .Controls(arr_varFld(F_LBL1, lngX)).Left = (lngTmp01 - lngTmp02)
32980               .Controls(arr_varFld(F_LIN, lngX)).Left = lngTmp01
32990               lngE = arr_varFld(F_CHK_ELEM, lngX)
33000               For lngY = C_CMD To C_ONDIS
33010                 .Controls(arr_varChk(lngY, lngE)).Left = lngTmp01
33020               Next  ' ** lngY.
33030             End If
33040             lngTmp01 = (lngTmp01 + arr_varFld(F_LIN_WDT, lngX) + lngFldSep)
33050           Else
33060             If .Controls(arr_varFld(F_FNAM, lngX)).Left <> lngTmp01 Then
33070               .Controls(arr_varFld(F_FNAM, lngX)).Left = lngTmp01
33080               If arr_varFld(F_LFT, lngX) <> arr_varFld(F_LBL_LFT, lngX) Then
33090                 lngTmp02 = (arr_varFld(F_LIN_LFT, lngX) - arr_varFld(F_LBL_LFT, lngX))
33100                 .Controls(arr_varFld(F_LBL1, lngX)).Left = (lngTmp01 - lngTmp02)
33110               Else
33120                 .Controls(arr_varFld(F_LBL1, lngX)).Left = lngTmp01
33130               End If
33140               If IsNull(arr_varFld(F_LBL2, lngX)) = False Then
33150                 .Controls(arr_varFld(F_LBL2, lngX)).Left = .Controls(arr_varFld(F_LBL1, lngX)).Left
33160               End If
33170               .Controls(arr_varFld(F_LIN, lngX)).Left = lngTmp01
33180               lngE = arr_varFld(F_CHK_ELEM, lngX)
33190               For lngY = C_CMD To C_ONDIS
33200                 .Controls(arr_varChk(lngY, lngE)).Left = lngTmp01
33210               Next  ' ** lngY.
33220             End If
                  ' ** Special cases.
33230             Select Case arr_varFld(F_FNAM, lngX)
                  Case "icash"
33240               .icash_box.Left = .ICash.Left
33250               .icash_str_bg.Left = (.ICash.Left - lngTpp)
33260               .icash_str.Left = .ICash.Left
33270             Case "pcash"
33280               .pcash_box.Left = .PCash.Left
33290               .pcash_str_bg.Left = (.PCash.Left - lngTpp)
33300               .pcash_str.Left = .PCash.Left
33310             Case "cost"
33320               .cost_box.Left = .Cost.Left
33330               .cost_str_bg.Left = (.Cost.Left - lngTpp)
33340               .cost_str.Left = .Cost.Left
33350             Case "curr_id"
33360               .curr_id_box.Left = .curr_id.Left
33370               .Curr_id_bg.Left = (.curr_id.Left - lngTpp)
33380             Case Else
                    ' ** Nothing else.
33390             End Select
33400             lngTmp01 = (lngTmp01 + arr_varFld(F_WDT, lngX) + lngFldSep)
33410           End If
33420           If blnSortHere = True Then
33430             .Sort_line.Left = .Controls(arr_varFld(F_LIN, lngX)).Left
33440             .Sort_lbl.Left = (((.Controls(arr_varFld(F_LBL1, lngX)).Left + arr_varFld(F_LBL_WDT, lngX)) - _
                    lngSortLbl_Width) + (arr_varFld(F_SRT_ADJ, lngX) * lngTpp))
33450           End If
33460         Else
                ' ** These remain hidden.
33470         End If
33480       Next  ' ** lngX.

            ' ** Check toggle button visibility.
            'For lngX = 0& To (lngFlds - 1&)
            '  If arr_varFld(F_VIS, lngX) = False Then
            '    lngE = arr_varFld(F_CHK_ELEM, lngX)
            '    For lngY = C_CMD To C_ONDIS
            '      .Controls(arr_varChk(lngY, lngE)).Visible = False
            '    Next  ' ** lngY.
            '  End If
            'Next  ' ** lngX

33490     Case 3  ' ** Hide field.

33500       .FocusHolder.SetFocus

33510       strCtlName = Left(strProc, (CharPos(strProc, 2, "_") - 1))  ' ** Module Function: modStringFuncs.

33520       lngZ = -1&
33530       For lngX = 0& To (lngFlds - 1&)
33540         If arr_varFld(F_CNAM, lngX) = strCtlName Then
33550           lngZ = lngX
33560           Exit For
33570         End If
33580       Next  ' ** lngX

33590       blnSortHere = False: blnResort = False
33600       If .Sort_line.Left = .Controls(arr_varFld(F_LIN, lngZ)).Left Then
33610         blnSortHere = True
33620       End If

33630       strFieldName = arr_varFld(F_FNAM, lngZ)
33640       .Controls(arr_varFld(F_FNAM, lngZ)).Visible = False
33650       .Controls(arr_varFld(F_LBL1, lngZ)).Visible = False
33660       If IsNull(arr_varFld(F_LBL2, lngZ)) = False Then
33670         .Controls(arr_varFld(F_LBL2, lngZ)).Visible = False
33680       End If
33690       .Controls(arr_varFld(F_LIN, lngZ)).Visible = False
33700       arr_varFld(F_VIS, lngZ) = CBool(False)

            ' ** Special cases.
33710       Select Case strFieldName
            Case "icash"
33720         .icash_box.Visible = False
33730         .icash_str_bg.Visible = False
33740         .icash_str.Visible = False
33750       Case "pcash"
33760         .pcash_box.Visible = False
33770         .pcash_str_bg.Visible = False
33780         .pcash_str.Visible = False
33790       Case "cost"
33800         .cost_box.Visible = False
33810         .cost_str_bg.Visible = False
33820         .cost_str.Visible = False
33830       Case "curr_id"
33840         .curr_id_box.Visible = False
33850         .Curr_id_bg.Visible = False
33860       Case Else
              ' ** Nothing else.
33870       End Select

33880       If blnSortHere = True Then
33890         .Sort_lbl.Visible = False
33900         .Sort_line.Visible = False
33910         blnResort = True
33920       End If

33930       lngE = arr_varFld(F_CHK_ELEM, lngZ)
33940       For lngX = C_CMD To C_ONDIS
33950         .Controls(arr_varChk(lngX, lngE)).Visible = False
33960       Next  ' ** lngX.

            ' ** Now close up the holes.
            ' ** Field separation is always 45 Twips, 3 pixels: lngFldSep.
33970       lngTmp01 = .journalno.Left
33980       For lngX = 0& To (lngFlds - 1&)
33990         blnSortHere = False
34000         If arr_varFld(F_FNAM, lngX) = strFieldName Then
                ' ** This one's now gone.
34010         ElseIf arr_varFld(F_VIS, lngX) = True Then
34020           If arr_varFld(F_FNAM, lngX) = "ledger_HIDDEN" Then
                  ' ** Label is offset.
34030             If .Sort_line.Left = .Controls(arr_varFld(F_LIN, lngX)).Left Then blnSortHere = True
34040             If .Controls(arr_varFld(F_LIN, lngX)).Left <> lngTmp01 Then
                    ' ** Move this one to its new left.
34050               lngTmp02 = (arr_varFld(F_LFT, lngX) - arr_varFld(F_LIN_LFT, lngX))
34060               .Controls(arr_varFld(F_FNAM, lngX)).Left = (lngTmp01 + lngTmp02)
34070               lngTmp02 = (arr_varFld(F_LIN_LFT, lngX) - arr_varFld(F_LBL_LFT, lngX))
34080               .Controls(arr_varFld(F_LBL1, lngX)).Left = (lngTmp01 - lngTmp02)
34090               .Controls(arr_varFld(F_LIN, lngX)).Left = lngTmp01
34100               lngE = arr_varFld(F_CHK_ELEM, lngX)
34110               For lngY = C_CMD To C_ONDIS
34120                 .Controls(arr_varChk(lngY, lngE)).Left = lngTmp01
34130               Next  ' ** lngY.
34140             End If
34150             lngTmp01 = lngTmp01 + arr_varFld(F_LIN_WDT, lngX) + lngFldSep
34160             If blnSortHere = True Then
34170               .Sort_line.Left = .Controls(arr_varFld(F_LIN, lngX)).Left
34180               .Sort_lbl.Left = (((.Controls(arr_varFld(F_LBL1, lngX)).Left + arr_varFld(F_LBL_WDT, lngX)) - _
                      lngSortLbl_Width) + (arr_varFld(F_SRT_ADJ, lngX) * lngTpp))
34190             End If
34200           Else
34210             If .Sort_line.Left = .Controls(arr_varFld(F_LIN, lngX)).Left Then blnSortHere = True
34220             If .Controls(arr_varFld(F_FNAM, lngX)).Left <> lngTmp01 Then
                    ' ** Move this one to its new left.
34230               .Controls(arr_varFld(F_FNAM, lngX)).Left = lngTmp01
34240               If arr_varFld(F_LBL_LFT, lngX) <> arr_varFld(F_LFT, lngX) Then
34250                 lngTmp02 = (arr_varFld(F_LIN_LFT, lngX) - arr_varFld(F_LBL_LFT, lngX))
34260                 .Controls(arr_varFld(F_LBL1, lngX)).Left = (lngTmp01 - lngTmp02)
34270               Else
34280                 .Controls(arr_varFld(F_LBL1, lngX)).Left = lngTmp01
34290               End If
34300               If IsNull(arr_varFld(F_LBL2, lngX)) = False Then
34310                 .Controls(arr_varFld(F_LBL2, lngX)).Left = .Controls(arr_varFld(F_LBL1, lngX)).Left
34320               End If
34330               .Controls(arr_varFld(F_LIN, lngX)).Left = lngTmp01
34340               lngE = arr_varFld(F_CHK_ELEM, lngX)
34350               For lngY = C_CMD To C_ONDIS
34360                 .Controls(arr_varChk(lngY, lngE)).Left = lngTmp01
34370               Next  ' ** lngY.
34380             End If
34390             lngTmp01 = lngTmp01 + arr_varFld(F_WDT, lngX) + lngFldSep
34400             If blnSortHere = True Then
34410               .Sort_line.Left = .Controls(arr_varFld(F_LIN, lngX)).Left
34420               .Sort_lbl.Left = (((.Controls(arr_varFld(F_LBL1, lngX)).Left + arr_varFld(F_LBL_WDT, lngX)) - _
                      lngSortLbl_Width) + (arr_varFld(F_SRT_ADJ, lngX) * lngTpp))
34430             End If
                  ' ** Special cases.
34440             Select Case arr_varFld(F_FNAM, lngX)
                  Case "icash"
34450               .icash_box.Left = .ICash.Left
34460               .icash_str_bg.Left = (.ICash.Left - lngTpp)
34470               .icash_str.Left = .ICash.Left
34480             Case "pcash"
34490               .pcash_box.Left = .PCash.Left
34500               .pcash_str_bg.Left = (.PCash.Left - lngTpp)
34510               .pcash_str.Left = .PCash.Left
34520             Case "cost"
34530               .cost_box.Left = .Cost.Left
34540               .cost_str_bg.Left = (.Cost.Left - lngTpp)
34550               .cost_str.Left = .Cost.Left
34560             Case "curr_id"
34570               .curr_id_box.Left = .curr_id.Left
34580               .Curr_id_bg.Left = (.curr_id.Left - lngTpp)
34590             Case Else
                    ' ** Nothing else.
34600             End Select
34610           End If

34620         Else
                ' ** This one was already off.
34630         End If
34640       Next  ' ** lngX.

34650       If blnResort = True Then
              ' ** If journalno is turned off, go to transdate, otherwise goto journalno.
34660         Select Case .journalno.Visible
              Case True
34670           .SortNow "Form_Load"  ' ** Form Procedure: frmTransaction_Audit_Sub.
34680         Case False
34690           Select Case .transdate.Visible
                Case True
34700             .SortNow "transdate_lbl_DblClick"  ' ** Form Procedure: frmTransaction_Audit_Sub.
34710           Case False
34720             Select Case .accountno.Visible
                  Case True
34730               .SortNow "accountno_lbl_DblClick"  ' ** Form Procedure: frmTransaction_Audit_Sub.
34740             Case False
34750               Select Case .journaltype.Visible
                    Case True
34760                 .SortNow "journaltype_lbl_DblClick"  ' ** Form Procedure: frmTransaction_Audit_Sub.
34770               Case False
34780                 Select Case .posted.Visible
                      Case True
34790                   .SortNow "posted_lbl_DblClick"  ' ** Form Procedure: frmTransaction_Audit_Sub.
34800                 Case False
34810                   For lngX = 0& To (lngFlds - 1&)
34820                     If arr_varFld(F_VIS, lngX) = True Then
                            ' ** And the lucky winner is...
34830                       .SortNow arr_varFld(F_FNAM, lngX) & "_lbl_DblClick"  ' ** Form Procedure: frmTransaction_Audit_Sub.
34840                     End If
34850                   Next  ' ** lngX.
34860                 End Select
34870               End Select
34880             End Select
34890           End Select
34900         End Select
34910       End If
34920       DoEvents

            ' ** Check toggle button visibility.
            'For lngX = 0& To (lngFlds - 1&)
            '  If arr_varFld(F_VIS, lngX) = False Then
            '    lngE = arr_varFld(F_CHK_ELEM, lngX)
            '    For lngY = C_CMD To C_ONDIS
            '      .Controls(arr_varChk(lngY, lngE)).Visible = False
            '    Next  ' ** lngY.
            '  End If
            'Next  ' ** lngX

34930     End Select  ' ** intMode.

34940   End With

        ' ** ckgFlds_chkJournalno_AfterUpdate
        ' ** ckgFlds_chkJournalType_AfterUpdate
        ' ** ckgFlds_chkTransDate_AfterUpdate
        ' ** ckgFlds_chkAccountNo_AfterUpdate
        ' ** ckgFlds_chkShortName_AfterUpdate
        ' ** ckgFlds_chkCusip_AfterUpdate
        ' ** ckgFlds_chkAssetDescription_AfterUpdate
        ' ** ckgFlds_chkShareFace_AfterUpdate
        ' ** ckgFlds_chkICash_AfterUpdate
        ' ** ckgFlds_chkPCash_AfterUpdate
        ' ** ckgFlds_chkCost_AfterUpdate
        ' ** ckgFlds_chkAssetDate_AfterUpdate
        ' ** ckgFlds_chkPurchaseDate_AfterUpdate
        ' ** ckgFlds_chkLedgerDescription_AfterUpdate
        ' ** ckgFlds_chkRecurringItem_AfterUpdate
        ' ** ckgFlds_chkRevCodeDesc_AfterUpdate
        ' ** ckgFlds_chkRevCodeTypeDescription_AfterUpdate
        ' ** ckgFlds_chkTaxCodeDescription_AfterUpdate
        ' ** ckgFlds_chkTaxCodeTypeDescription_AfterUpdate
        ' ** ckgFlds_chkJournalUser_AfterUpdate
        ' ** ckgFlds_chkPosted_AfterUpdate
        ' ** ckgFlds_chkLedgerHidden_AfterUpdate
        ' ** ckgFlds_chkCheckNum_AfterUpdate
        ' ** ckgFlds_chkCurrID_AfterUpdate

        ' **      Field                             Twips  Pixels
        ' **      ================================  =====  ======
        ' ** 1.   journalno.Width                     780     52
        ' ** 2.   journaltype.Width                  1020     68
        ' ** 3.   transdate.Width                     900     60
        ' ** 4.   accountno.Width                    1200     80
        ' ** 5.   shortname.Width                    2160    144
        ' ** 6.   cusip.Width                        1080     72
        ' ** 7.   asset_description.Width            2100    140
        ' ** 8.   shareface.Width                    1080     72
        ' ** 9.   icash.Width                        1440     96
        ' ** 10.  pcash.Width                        1440     96
        ' ** 11.  cost.Width                         1440     96
        ' ** 12.  curr_id                             840     56
        ' ** 13.  assetdate.Width                    1575    105
        ' ** 14.  PurchaseDate.Width                 1575    105
        ' ** 15.  ledger_description.Width           1980    132
        ' ** 16.  RecurringItem.Width                1980    132
        ' ** 17.  revcode_DESC.Width                 1920    128
        ' ** 18.  revcode_TYPE_Description.Width      810     54
        ' ** 19.  taxcode_description.Width          1920    128
        ' ** 20.  taxcode_type_description.Width      915     61
        ' ** 21.  Location_Name.Width                1920    128
        ' ** 22.  CheckNum                            780     52
        ' ** 23.  journal_USER.Width                 1020     68
        ' ** 24.  posted.Width                       1575    105
        ' ** 25.  ledger_HIDDEN_lbl.Width             720     48

        ' *********************************************************************
        ' ** Array: arr_varChk()
        ' **
        ' **   Field  Element  Name                                Constant
        ' **   =====  =======  ==================================  ==========
        ' **     1       0     Field Name                          C_FNAM
        ' **     2       1     Check Box                           C_CHKBX
        ' **     3       2     Include T/F                         C_INCL
        ' **     4       3     Focus T/F                           C_FOC
        ' **     5       4     MouseDown T/F                       C_MOUS
        ' **     6       5     Disabled T/F                        C_DIS
        ' **     7       6     Command Button                      C_CMD
        ' **     8       7     .._tgl_off_raised_img               C_OFR
        ' **     9       8     .._tgl_off_raised_dots_img          C_OFRD
        ' **    10       9     .._tgl_off_raised_focus_img         C_OFRF
        ' **    11      10     .._tgl_off_raised_focus_dots_img    C_OFRFD
        ' **    12      11     .._tgl_off_raised_img_dis           C_OFDIS
        ' **    13      12     .._tgl_on_raised_img                C_ONR
        ' **    14      13     .._tgl_on_raised_dots_img           C_ONRD
        ' **    15      14     .._tgl_on_raised_focus_img          C_ONRF
        ' **    16      15     .._tgl_on_raised_focus_dots_img     C_ONRFD
        ' **    17      16     .._tgl_on_sunken_dots_img           C_ONSD
        ' **    18      17     .._tgl_on_raised_img_dis            C_ONDIS
        ' **
        ' *********************************************************************

        'journalno
        'journaltype
        'transdate
        'accountno
        'shortname
        'cusip
        'asset_description
        'shareface
        'icash
        'pcash
        'cost
        'curr_id
        'assetdate
        'PurchaseDate
        'ledger_description
        'RecurringItem
        'revcode_DESC
        'revcode_TYPE_Description
        'taxcode_description
        'taxcode_type_description
        'CheckNum
        'journal_USER
        'posted
        'ledger_HIDDEN

EXITP:
34950   Set ctl1 = Nothing
34960   Set ctl2 = Nothing
34970   Exit Sub

ERRH:
34980   Select Case ERR.Number
        Case Else
34990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35000   End Select
35010   Resume EXITP

End Sub

Public Function FilterRecs_Cnt_TA(strFilter As String, dblFilterRecs As Double, rstAll As DAO.Recordset, frm As Access.Form) As Boolean
' ** All the original calls still go to the subform,
' ** which then calls this with the additional parameters.

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Cnt_TA"

        Dim rst As DAO.Recordset, ctl As Access.Control
        Dim blnRetVal As Boolean

35110   blnRetVal = True

35120   If strFilter <> vbNullString Then
35130     With frm
35140       DoCmd.Hourglass True
35150       rstAll.Filter = strFilter
35160       .Parent.CurrentFilter1 = strFilter
35170       Set rst = rstAll.OpenRecordset
35180       If rst.BOF = True And rst.EOF = True Then
35190         blnRetVal = False
35200       Else
35210         rst.MoveLast
35220         dblFilterRecs = rst.RecordCount
35230         .Parent.FilterRecs_Set strFilter, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit.
35240       End If
35250       rst.Close
35260       rstAll.Filter = vbNullString
35270       .Parent.CurrentFilter1 = vbNullString
35280       If blnRetVal = False Then
35290         .Parent.FilterRecs = "None"
35300         .Parent.FilterRecs.ForeColor = CLR_DKRED
35310         .Parent.FilterRecs_lbl.ForeColor = CLR_DKRED
35320         For Each ctl In .Section("Detail").Controls
35330           With ctl
35340             If .Visible = True Then
35350               Select Case .ControlType
                    Case acTextBox, acComboBox
35360                 .BackColor = CLR_GRY2
35370               End Select
35380             End If
35390           End With
35400         Next
35410         .Parent.FilterRecsZero  ' ** Form Procedure: frmTransaction_Audit.
35420       Else
35430         .Filter = strFilter
35440         .Parent.CurrentFilter1 = strFilter
35450         .FilterOn = True
35460         .Parent.FilterRecs = Format(dblFilterRecs, "#,##0")
35470         .Parent.FilterRecs.ForeColor = CLR_BLU
35480         .Parent.FilterRecs_lbl.ForeColor = CLR_BLU
35490         For Each ctl In .Section("Detail").Controls
35500           With ctl
35510             If .Visible = True Then
35520               Select Case .ControlType
                    Case acTextBox, acComboBox
35530                 .BackColor = CLR_WHT
35540               End Select
35550             End If
35560           End With
35570         Next
35580       End If
35590       DoCmd.Hourglass False
35600     End With
35610   Else
35620     With frm
35630       If .Filter <> vbNullString Then
35640         .Filter = vbNullString
35650         .Parent.CurrentFilter1 = vbNullString
35660       End If
35670       If .FilterOn = True Then
35680         .FilterOn = False
35690       End If
35700       .Parent.FilterRecs = "All"
35710       .Parent.FilterRecs.ForeColor = CLR_DISABLED_FG
35720       .Parent.FilterRecs.BackColor = CLR_DISABLED_BG
35730       .Parent.FilterRecs_lbl.ForeColor = CLR_VDKGRY
35740       For Each ctl In .Section("Detail").Controls
35750         With ctl
35760           If .Visible = True Then
35770             Select Case .ControlType
                  Case acTextBox, acComboBox
35780               .BackColor = CLR_WHT
35790             End Select
35800           End If
35810         End With
35820       Next
35830     End With
35840   End If

EXITP:
35850   Set ctl = Nothing
35860   Set rst = Nothing
35870   FilterRecs_Cnt_TA = blnRetVal
35880   Exit Function

ERRH:
35890   blnRetVal = False
35900   DoCmd.Hourglass False
35910   Select Case ERR.Number
        Case Else
35920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35930   End Select
35940   Resume EXITP

End Function

Public Sub FilterRecs_Rem_TA(strClause As String, strFilter As String, dblFilterRecs As Double, frm As Access.Form)
' ** Remove a clause from the filter string.
' ** All the original calls still go to the subform,
' ** which then calls this with the additional parameters.

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "FilterRecs_Rem_TA"

        Dim blnIsJType As Boolean, blnIsCheckNum As Boolean, blnIsHidden As Boolean
        Dim lngMultiCnt As Long
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const CHK_NUM      As String = "[CheckNum] = "
        Const CHK_NUM1     As String = "[CheckNum] >= "
        Const CHK_NUM2     As String = "[CheckNum] <= "
        Const HIDDEN_TRX1  As String = "[ledger_HIDDEN] = True"
        Const HIDDEN_TRX2  As String = "[ledger_HIDDEN] = False"

36010   If strClause <> vbNullString And strFilter <> vbNullString Then
36020     If InStr(strClause, "[journaltype]") > 0 Then blnIsJType = True Else blnIsJType = False
36030     If InStr(strClause, "[CheckNum]") > 0 Then blnIsCheckNum = True Else blnIsCheckNum = False
36040     If InStr(strClause, "[ledger_HIDDEN]") > 0 Then blnIsHidden = True Else blnIsHidden = False
36050     With frm
36060       intPos01 = InStr(strFilter, strClause)
36070       If blnIsCheckNum = True And intPos01 = 0 Then
36080         intPos01 = InStr(strFilter, CHK_NUM1)
36090         If intPos01 = 0 Then
36100           intPos01 = InStr(strFilter, CHK_NUM2)
36110         End If
36120       End If
36130       If blnIsHidden = True And intPos01 = 0 Then
36140         intPos01 = InStr(strFilter, HIDDEN_TRX2)
36150       End If
36160       If intPos01 > 0 Then
              ' ** This clause is present
36170         Select Case blnIsCheckNum
              Case True
36180           intPos01 = InStr(strFilter, CHK_NUM)
36190           intPos02 = InStr(strFilter, CHK_NUM1)
36200           intPos03 = InStr(strFilter, CHK_NUM2)
36210           strTmp01 = vbNullString: strTmp02 = vbNullString
36220           If intPos01 > 0 Then
                  ' ** Simple equaltiy.
36230             strTmp01 = Left(strFilter, (intPos01 - 1))
36240             intPos04 = InStr((intPos01 + Len(CHK_NUM) + 1), strFilter, ANDF)
36250             If intPos04 > 0 Then
36260               strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
36270             End If
36280           ElseIf intPos02 > 0 And intPos03 > 0 Then
                  ' ** A range of check numbers.
36290             strTmp01 = Left(strFilter, (intPos02 - 1))
36300             If Right(Trim(strTmp01), 1) = "(" Then strTmp01 = Left(Trim(strTmp01), (Len(Trim(strTmp01)) - 1))
36310             intPos04 = InStr((intPos03 + Len(CHK_NUM2) + 1), strFilter, ANDF)  ' ** This will miss closing paren.
36320             If intPos04 > 0 Then
36330               strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
36340             End If
36350           ElseIf intPos02 > 0 Then
                  ' ** Greater than.
36360             strTmp01 = Left(strFilter, (intPos02 - 1))
36370             intPos04 = InStr((intPos02 + Len(CHK_NUM1) + 1), strFilter, ANDF)
36380             If intPos04 > 0 Then
36390               strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
36400             End If
36410           ElseIf intPos03 > 0 Then
                  ' ** Less than.
36420             strTmp01 = Left(strFilter, (intPos03 - 1))
36430             intPos04 = InStr((intPos03 + Len(CHK_NUM2) + 1), strFilter, ANDF)
36440             If intPos04 > 0 Then
36450               strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
36460             End If
36470           End If
36480           If Trim(strTmp01) <> vbNullString Then
36490             If Right(Trim(strTmp01), 4) = " And" Then
                    ' ** Remove this clause's preceeding 'And'.
36500               strTmp01 = Trim(Left(Trim(strTmp01), (Len(Trim(strTmp01)) - 3)))
36510             End If
36520           Else
                  ' ** Clause is beginning of line.
36530             strTmp01 = Trim(strTmp01)
36540           End If
36550           strFilter = strTmp01 & strTmp02  ' ** Both may be vbNullString.
36560         Case False
36570           Select Case blnIsHidden
                Case True
36580             intPos01 = InStr(strFilter, HIDDEN_TRX1)
36590             intPos02 = InStr(strFilter, HIDDEN_TRX2)
36600             strTmp01 = vbNullString: strTmp02 = vbNullString
36610             If intPos01 > 0 Then
                    ' ** Only hidden transactions.
36620               strTmp01 = Left(strFilter, (intPos01 - 1))
36630               intPos04 = InStr(intPos01, strFilter, ANDF)
36640               If intPos04 > 0 Then
36650                 strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
36660               End If
36670             ElseIf intPos02 > 0 Then
                    ' ** Excludes hidden transactions.
36680               strTmp01 = Left(strFilter, (intPos02 - 1))
36690               intPos04 = InStr(intPos02, strFilter, ANDF)
36700               If intPos04 > 0 Then
36710                 strTmp02 = Mid(strFilter, intPos04)  ' ** Includes space before 'And'.
36720               End If
36730             End If
36740             If Trim(strTmp01) <> vbNullString Then
36750               If Right(Trim(strTmp01), 4) = " And" Then
                      ' ** Remove this clause's preceeding 'And'.
36760                 strTmp01 = Trim(Left(Trim(strTmp01), (Len(Trim(strTmp01)) - 3)))
36770               End If
36780             Else
                    ' ** Clause is beginning of line.
36790               strTmp01 = Trim(strTmp01)
36800             End If
36810             strFilter = strTmp01 & strTmp02  ' ** Both may be vbNullString.
36820           Case False
36830             Select Case blnIsJType
                  Case True
36840               lngMultiCnt = CharCnt(strFilter, strClause, True)  ' ** Module Function: modStringFuncs.
36850               Select Case lngMultiCnt
                    Case 1&  ' ** For 1 occurance only.
36860                 intPos02 = InStr(intPos01, strFilter, ")" & ANDF)  ' ** Searches for following ') And '.
36870               Case 2&, 3&  ' ** For 1st of 2 or 3 occurances.
36880                 intPos02 = InStr(intPos01, strFilter, ORF)  ' ** Searches for following ' Or '.
36890               End Select
36900             Case False
36910               intPos02 = InStr(intPos01, strFilter, ANDF)  ' ** Searches for following ' And '.
36920             End Select
36930             If intPos01 = 1 Or (blnIsJType = True And intPos01 = 2) Then
                    ' ** This clause starts the filter.
36940               If intPos02 = 0 Then
                      ' ** This clause is the only clause.
36950                 strFilter = vbNullString
36960               Else
                      ' ** There's another clause after this one.
36970                 Select Case blnIsJType
                      Case True
36980                   Select Case lngMultiCnt
                        Case 1&  ' ** For 1 occurance only.
36990                     strFilter = Mid(strFilter, (intPos02 + Len(")" & ANDF)))  ' ** Pick up after ') And '.
37000                   Case 2&, 3&  ' ** For 1st of 2 or 3 occurances.
37010                     strFilter = Mid(strFilter, (intPos02 + Len(ORF)))  ' ** Pick up after ' Or '.
37020                   End Select
37030                 Case False
37040                   strFilter = Mid(strFilter, (intPos02 + Len(ANDF)))  ' ** Pick up after ' And '.
37050                 End Select
37060               End If
37070             Else
                    ' ** There's another clause before this one.
37080               If intPos02 = 0 Then
                      ' ** This clause is the last one.
37090                 Select Case blnIsJType
                      Case True  ' ** For 1 occurance only (otherwise it wouldn't be the last one).
37100                   strFilter = Left(strFilter, (intPos01 - Len(ANDF & "(")))  ' ** Remove final ' And ('.
37110                 Case False
37120                   strFilter = Left(strFilter, (intPos01 - Len(ANDF)))  ' ** Remove final ' And '.
37130                 End Select
37140               Else
                      ' ** There's a clause both before and after this one.
37150                 Select Case blnIsJType
                      Case True
37160                   Select Case lngMultiCnt
                        Case 1&  ' ** For 1 occurance only.
37170                     strFilter = Left(strFilter, (intPos01 - 2)) & Mid(strFilter, (intPos02 + Len(")" & ANDF)))
37180                   Case 2&, 3&  ' ** For 1st of 2 or 3 occurances.
37190                     strFilter = Left(strFilter, (intPos01 - 1)) & Mid(strFilter, (intPos02 + Len(ORF)))
37200                   End Select
37210                 Case False
37220                   strFilter = Left(strFilter, (intPos01 - 1)) & Mid(strFilter, (intPos02 + Len(ANDF)))  ' ** Don't leave 2 spaces or 2 And's.
37230                 End Select
37240               End If
37250             End If
37260           End Select  ' ** blnIsHidden.
37270         End Select  ' ** blnIsCheckNum.
37280         .Filter = strFilter
37290         DoEvents
37300         dblFilterRecs = .RecCnt  ' ** Form Function: frmTransaction_Audit_Sub.
37310         .Parent.CurrentFilter1 = strFilter
37320         .Parent.FilterRecs_Set strFilter, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit.
37330         frmCrit.FilterRecs_Set strFilter, dblFilterRecs  ' ** Form Procedure: frmTransaction_Audit_Sub_Criteria.
37340         If strFilter = vbNullString Then
37350           .FilterOn = False
37360         Else
37370           .FilterOn = True
37380         End If
37390       End If
37400     End With
37410   End If

EXITP:
37420   Exit Sub

ERRH:
37430   Select Case ERR.Number
        Case Else
37440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37450   End Select
37460   Resume EXITP

End Sub

Public Function FldArray_Get_TA() As Variant

37500 On Error GoTo ERRH

        Const THIS_PROC As String = "FldArray_Get_TA"

        Dim arr_varRetVal As Variant

37510   arr_varRetVal = arr_varFld

EXITP:
37520   FldArray_Get_TA = arr_varRetVal
37530   Exit Function

ERRH:
37540   Select Case ERR.Number
        Case Else
37550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37560   End Select
37570   Resume EXITP

End Function
