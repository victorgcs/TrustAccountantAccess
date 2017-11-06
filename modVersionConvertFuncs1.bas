Attribute VB_Name = "modVersionConvertFuncs1"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modVersionConvertFuncs1"

'VGC 09/28/2017: CHANGES!

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' ** CHECK OTHER COUNTRY FIELDS!
' ** CompanyInformation
' ** Account
' ** Location
' ** RecurringItems
' ** AccountContacts
' ** Version_Input
'NOW, DON'T FORGET NEW CURR_ID'S!
' ** Tables in TrustDta.mdb with curr_id:
' **   account                         Version_Upgrade_04()  OK!
' **   Balance                         Version_Upgrade_04()  OK!
' **   masterasset                     Version_Upgrade_04()  OK!
' **   ActiveAssets                    Version_Upgrade_04()  OK!
' **   ledger                          Version_Upgrade_05()  OK!
' **   journal                         Version_Upgrade_05()  OK!
' **   tblPricing_MasterAsset_History  Version_Upgrade_06()  OK!
' **   tblCurrency                     Version_Upgrade_07()  OK!
' **   tblCurrency_History             Version_Upgrade_07()  OK!
' **   tblCountry_Currency             {not converted}
' **   tblCurrency_Symbol              {not converted}
' **   tblCountry_Currency_Primary     {not converted}
' **   tblCurrency_Country_Primary     {not converted}
' **   journal Map                     {not converted}
' ** Tables in TrstArch.mdb with curr_id:
' **   ledger                          Version_Upgrade_06()  OK!
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

'########################
'X PASSWORD CHANGES: DONE! (At least it looks like it.)
'X Users: ALL DEFAULTS NOW INCLUDED!
'X tblSecurity_License
'X tblJournal_Memo
'X Schedule
'X ScheduleDetail
'########################

'HOW DO WE DEAL WITH THAT USER WHOSE Admin PASSWORD GOT CHANGED?
'AAH! LOG IN AS superuser, AND RESET THE Admin PASSWORD!

'CONSIDER MOVING PROCEDURES THAT ONLY CHECK IF A CONVERSION IS NECESSARY TO modStartupFuncs,
'SO THIS MODULE DOESN'T GET LOADED IF IT'S NOT GOING TO BE USED FOR A CONVERSION.

' ** The number of steps in the conversion are set in Version_Status(), below.
' **   dblPB_Steps = 25#

'On an old v1.6.nn, if Trust.mde is available, I can use a Master.mdw
'to link or import the 'License Name' table, where the version is listed.
'REVIEW VERSION CHECKING IN ORDER TO AVOID strOldVersion EQUALLING 'v1.6.00?'.

' ** Related forms.
Public Const FRM_CNV_STATUS As String = "frmVersion_Main"
Public Const FRM_CNV_INPUT1 As String = "frmVersion_Input"

' ** Bad field names to watch out for.
Private lngBadNames As Long, arr_varBadName() As Variant
Private blnBadName As Boolean ', lngBadElem As Long
Private lngErrNum As Long, lngErrLine As Long, strErrDesc As String, lngDupeNum As Long

' ** Array: arr_varDupeUnk().
Private lngDupeUnks As Long, arr_varDupeUnk() As Variant
Private Const DU_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const DU_TYP As Integer = 0
Private Const DU_TBL As Integer = 1
' ** The #DUPE identifier may show up in these tables:
' **   adminofficer
' **   Location
' **   RecurringItems
' **   Schedule
' **   m_REVCODE
' **   Users
' **   account
' ** The UNKNOWN identifier may show up in these tables:
' **   adminofficer
' **   Schedule
' **   masterasset

Private strFailMsg As String, strSummaryMsg As String, strTruncatedFields As String
Public gintConvertResponse As Integer
Private lngStats As Long, arr_varStat() As Variant

' ##################################################
' ## Unsettled issues:
' ## 1. Reconverting, change 'BAK' to 'mdb'.
' ## 2. Mixed prefix/non-prefix account nums.
' ## 3. Dupes, with an 'X'.
' ##################################################

'At the beginning of the install,
'Wise copies existing TrustDta.mdb to
'  %TEMP%\Trust_bak\TrustDta.mdb
'Wise copies existing TrstArch.mdb to
'  %TEMP%\Trust_bak\TrstArch.mdb
'As well as TrustSec.mdw, TA.lic, DDTrust.ini

'Then, after running TrustUpd.mde,
'Wise copied %TEMP%\Trust_bak\TrstArch.BAK to
'  %TRSTARCH%\TrstArch.BAK
'Wise copied %TEMP%\Trust_bak\TrstArch.INST to
'  %TRSTARCH%\TrstArch.INST
'Wise copied %TEMP%\Trust_bak\TrustDta.BAK to
'  %TRUSTDTA%\TrustDta.BAK
'Wise copied %TEMP%\Trust_bak\TrustDta.INST to
'  %TRUSTDTA%\TrustDta.INST
'It also copied TrustSec.mdw, TA.lic, DDTrust.ini
'to the local \Trust Accountant\Database\ folder.

'Instead, at the end, after Wise has run, we want
'the wrapper to copy the old TrustDta.mdb to
'  %TRUSTDTA%\Convert_New\
'and TrstArch.mdb to
'  %TRSTARCH%\Convert_New\
'and TrustSec.mdw, TA.lic, DDTrust.ini to
'  %TRUSTSEC%\TrustSec.mdw
'  %TALIC%\TA.lic
'  %TRUSTINI%\DDTrust.ini

'So, how do we want to do this?
'I would say, include an empty new database, from NewWorking.
'We'll need to ask them where the old data is, as we do now in the Upgrade Wise installer.
'In this case, DO NOT invoke TrustUpd.mdb; let an empty get installed and linked,
'then when it opens and goes to Backend_Update(), do everything off of that!
'We still have to move the old MDBs to a safe place on the side.
'If it's a standard location, check every time. If it finds it,
'that means it's an upgrade and needs to be processed.
'I'd say, DON'T Link the old MDBs (plural if it has a TrstArch.mdb),
'just connect via DAO in code.
'I could check for its version, but what would that gain me?
'Do I want to simply compile the table and field list on-the-fly, or
'rely on the version tables I've created?
'I could start with the existing list, then cross-check it with the actual.
'Then, one-at-a-time, read and transfer the data via recordsets.
'I'd have the most control that way.
'Create a special Conversion progress form, and show each table
'and record count as it's being converted.
'The biggest thing will be to test some Access 97 MDBs.
'Updated ID's will have to be kept track of, and trigger
'the cascade of all related tables.
'Include, of course, checking for orphans:
'  1. Ledger entries.
'  2. Journal entries.
'  3. Locations.
'  4. Recurring Items.
'  5. Statement, Review, and Fee Frequency stuff.
'     (Might I encounter a version old enough to have data in those tables?)
'  6. Portfolio Model?
'  7. MasterAsset and ActiveAssets especially.
'  8. Balance.
'Use the Convert_New directory. Copy an old MDB there, and process from there.
'Link to an empty set of MDBs, also there.
'Produce a summary screen of the convert, with the option to print.
'It'll have anomalies found, as well as listing User Names and Passwords reset.

' **************************************************
' **************************************************
' ** Conversion order:
' ** =================
' **
' ****************************
' ** Customer Static Data:
' ****************************
' **   These must be converted first, to provide
' **   references used in core dynamic data.
' **     adminofficer
' **     CompanyInformation
' **     Location
' **     RecurringItems
' **     Schedule
' **     ScheduleDetail
' **     DDate
' **     PostingDate
' **     Statement Date
' **     m_REVCODE
' **     tblSecurity_License
' **     Users
' **     _~xusr
' **
' ****************************
' ** Customer Dynamic Data:
' ****************************
' **   These hold the core TA customer data.
' **   Their conversion order is most important.
' **     Account
' **     Balance
' **     masterasset
' **     ActiveAssets
' **     ledger
' **     journal
' **     LedgerHidden
' **     LedgerArchive
' **     tblPricing_MasterAsset_History (tblAssetPricing)
' **     tblJournal_Memo
' **
' ****************************
' ** TA Static:
' ****************************
' **   These won't be converted;
' **   new copies replace them.
' **     accounttype
' **     AccountTypeGrouping
' **     assettype
' **     AssetTypeGrouping
' **     HiddenType
' **     InvestmentObjective
' **     journaltype
' **     m_REVCODE_TYPE
' **     m_TBL
' **     m_VD
' **     RecurringType
' **     State
' **     taxcode  (requires conversion)
' **
' ****************************
' ** TA Dynamic (Temporary):
' ****************************
' **   These only hold temporary data,
' **   so won't be converted.
' **     asset
' **     FeeCalculations
' **     Journal Map
' **     PortfolioModel
' **     tblErrorLog
' **     tblPortfolioModeling
' **
' ****************************
' ** TA Obsolete:
' ****************************
' **   These are remnants of earlier
' **   versions, and no longer used.
' **     assetsub
' **     feefreq
' **     jurisdiction
' **     Lock
' **     masterasset temp
' **     reviewfreq
' **     statementfreq
' **     tblAveragePrice
' **
' **************************************************
' **************************************************

' ** Previous versions documented and tested:
' **   Ver_1_7_00 X
' **   Ver_1_7_10 X
' **   Ver_1_7_20 X
' **   Ver_1_7_40 X
' **   Ver_2_0_00 X
' **   Ver_2_1_00 X
' **   Ver_2_1_20 X
' **   Ver_2_1_40 X
' **   Ver_2_1_44 X
' **   Ver_2_1_45 X
' **   Ver_2_1_46 X
' **   Ver_2_1_47 X
' **   Ver_2_1_50 X
' **   Ver_2_1_55 X
' **   Ver_2_1_56 X
' **   Ver_2_1_57 x

' ** Client files tested:
' **   Citizens
' **   Condit
' **   First Broken Arrow
' **   First Clarksdale
' **   First Palmerton
' **   First Pratt
' **   First Security
' **   Grundy
' **   Hedvig Klockars
' **   HFO
' **   Godfrey Kahn/LG
' **   Merchants & Planters
' **   Nachtigall
' **
' **
' **   Suplee & Shea

' ***************************************************************************************************************
' ** GetSpecialFolder Method:
' ** ========================
' **
' **   Description:
' **     Returns the special folder specified.
' **
' **   Syntax:
' **     Object.GetSpecialFolder (folderspec)
' **
' **     The GetSpecialFolder method syntax has these parts:
' **     Part          Description
' **     ============  ===========
' **     Object        Required. Always the name of a FileSystemObject.
' **     folderspec    Required. The name of the special folder to be returned.
' **                   Can be any of the constants shown in the Settings section.
' **
' **   Settings:
' **     The folderspec argument can have any of the following values:
' **     Constant           Value  Description
' **     =================  =====  ===========
' **     WindowsFolder        0    The Windows folder contains files installed by the Windows operating system.
' **     SystemFolder         1    The System folder contains libraries, fonts, and device drivers.
' **     TemporaryFolder      2    The Temp folder is used to store temporary files.
' **                               Its path is found in the TMP environment variable.
' ***************************************************************************************************************

' ***************************************************************************************************************
' ** CancelUpdate Method:
' ** ====================
' **
' **   Description:
' **     Cancels any pending updates for a Recordset object.
' **
' **   Syntax:
' **     recordset.CancelUpdate type
' **
' **     The CancelUpdate method syntax has these parts:
' **     Part         Description
' **     ===========  ===========
' **     recordset    Required. An object variable that represents the Recordset
' **                  object for which you are canceling pending updates.
' **     type         Optional. A constant indicating the type of update, as specified in Settings.
' **
' **   Settings:
' **     You can use the following values for the type argument only if batch updating is enabled.
' **     Constant           Value  Description
' **     =================  =====  ===========
' **     dbUpdateRegular      1    Default. Cancels pending changes that aren’t cached.
' **     dbUpdateBatch        4    Cancels pending changes in the update cache.
' **
' **   Remarks:
' **     You can use the CancelUpdate method to cancel any pending updates resulting from an Edit or AddNew
' **     operation. For example, if a user invokes the Edit or AddNew method and hasn't yet invoked the
' **     Update method, CancelUpdate cancels any changes made after Edit or AddNew was invoked.
' **     Check the EditMode property of the Recordset to determine if there is a pending
' **     operation that can be canceled.
' **
' **     Note: Using the CancelUpdate method has the same effect as moving to another record
' **     without using the Update method, except that the current record doesn't change,
' **     and various properties, such as BOF and EOF, aren't updated.
' ***************************************************************************************************************

' ***************************************************************************************************************
' ** EditMode Property:
' ** ==================
' **
' **   Description:
' **     Returns a value that indicates the state of editing for the current record.
' **
' **   Return Values:
' **     The return value is a Long that indicates the state of editing, as listed in the following table.
' **     Constant            Value  Description
' **     ==================  =====  ===========
' **     dbEditNone            0    No editing operation is in progress.
' **     dbEditInProgress      1    The Edit method has been invoked, and
' **                                the current record is in the copy buffer.
' **     dbEditAdd             2    The AddNew method has been invoked, and the current record in the
' **                                copy buffer is a new record that hasn't been saved in the database.
' **
' **   Remarks:
' **     The EditMode property is useful when an editing process is interrupted, for example, by an error
' **     during validation. You can use the value of the EditMode property to determine whether you should
' **     use the Update or CancelUpdate method.
' **     You can also check to see if the LockEdits property setting is True and the EditMode
' **     property setting is dbEditInProgress to determine whether the current page is locked.
' ***************************************************************************************************************

' ***************************************************************************************************************
' ** LockEdits Property:
' ** ===================
' **
' **   Description:
' **     Sets or returns a value indicating the type of locking that is in effect while editing.
' **
' **   Settings and Return Values:
' **     The setting or return value is a Boolean that indicates the type of locking,
' **     as specified in the following table.
' **     Value  Description
' **     =====  ===========
' **     True   Default. Pessimistic locking is in effect. The 2K page containing the
' **            record you're editing is locked as soon as you call the Edit method.
' **     False  Optimistic locking is in effect for editing. The 2K page containing
' **            the record is not locked until the Update method is executed.
' **
' **   Remarks:
' **     You can use the LockEdits property with updatable Recordset objects.
' **     If a page is locked, no other user can edit records on the same page. If you set LockEdits to
' **     True and another user already has the page locked, an error occurs when you use the Edit method.
' **     Other users can read data from locked pages.
' **     If you set the LockEdits property to False and later use the Update method while another user
' **     has the page locked, an error occurs. To see the changes made to your record by another user,
' **     use the Move method with 0 as the argument; however, if you do this, you will lose your changes.
' **     When working with Microsoft Jet-connected ODBC data sources, the LockEdits property
' **     is always set to False, or optimistic locking. The Microsoft Jet database engine
' **     has no control over the locking mechanisms used in external database servers.
' **
' **     Note: You can preset the value of LockEdits when you first open the Recordset by
' **     setting the lockedits argument of the OpenRecordset method. Setting the lockedits
' **     argument to dbPessimistic will set the LockEdits property to True, and setting
' **     lockedits to any other value will set the LockEdits property to False.
' ***************************************************************************************************************

Private Const A99_INC As String = "INCOME O/U"
Private Const A99_SUS As String = "SUSPENSE"

Private Const strVerTbl As String = "tblVersion_Conversion"
Private Const strKeyTbl As String = "tblVersion_Key"

' ** Temporary variables.
Public arr_varDataXFer() As Variant
Private varTmp00 As Variant, arr_varTmp01() As Variant, arr_varTmp02 As Variant, arr_varTmp03 As Variant
Private strTmp04 As String, strTmp05 As String, strTmp06 As String, strTmp07 As String
Private strTmp08 As String, strTmp09 As String, strTmp10 As String, strTmp11 As String, strTmp12 As String
Private lngTmp13 As Long, lngTmp14 As Long, lngTmp15 As Long, lngTmp16 As Long ', lngTmp17 As Long, lngTmp18 As Long, lngTmp19 As Long
Private lngTmp20 As Long, lngTmp21 As Long
Private blnTmp22 As Boolean, blnTmp23 As Boolean, blnTmp24 As Boolean, blnTmp25 As Boolean, blnTmp26 As Boolean, blnTmp27 As Boolean
Private datTmp28 As Date
Private strAcct99_IncomeOU As String, strAcct99_Suspense As String, blnAcct99_Both As Boolean
Private strOldVersion As String, lngVerCnvID As Long

' ** Various processing variables.
Private wrkLoc As DAO.Workspace, dbsLoc As DAO.Database, wrkLnk As DAO.Workspace, dbsLnk As DAO.Database
Private rstLoc1 As DAO.Recordset, rstLoc2 As DAO.Recordset, rstLoc3 As DAO.Recordset, rstLnk As DAO.Recordset
Private tdf As DAO.TableDef, fld As DAO.Field, doc As DAO.Document, prp As DAO.Property, qdf As DAO.QueryDef
Private fso As Scripting.FileSystemObject, fsfd As Scripting.Folder
Private fsfls As Scripting.FILES, fsfl As Scripting.File
Private lngOldFiles As Long, arr_varOldFile() As Variant
Private lngOldTbls As Long, arr_varOldTbl() As Variant
Private lngOldFlds As Long, arr_varOldFld() As Variant
Private blnConvert_TrustDta As Boolean, blnConvert_TrstArch As Boolean, blnArchiveNotPresent As Boolean
Private lngTrustDbsID As Long, lngTrustDtaDbsID As Long, lngTrstArchDbsID As Long
Private strCurrTblName As String, lngCurrTblID As Long
Private strCurrKeyFldName As String, lngCurrKeyFldID As Long
Private intWrkType As Integer
Private lngDtaElem As Long, lngArchElem As Long
Private lngAccts As Long, arr_varAcct() As Variant
Private lngRecs As Long, lngFlds As Long
Private lngOld_RECUR_I_TO_P_ID As Long, lngOld_RECUR_P_TO_I_ID As Long
Private lngSchedules As Long, arr_varSchedule() As Variant
Private lngRevCodes As Long, arr_varRevCode() As Variant
Private lngRevDefCodes As Long, arr_varRevDefCode() As Variant
Private lngInvestObjs As Long, arr_varInvestObj() As Variant
Private lngAcctTypes As Long, arr_varAcctType As Variant
Private lngAssetTypes As Long, arr_varAssetType As Variant
Private lngTaxDefCodes As Long, arr_varTaxDefCode As Variant
Private lngRecurTypes As Long, arr_varRecurType As Variant
Private lngLedgerEmptyDels As Long, lngLedgerArchEmptyDels As Long
Private lngMasterAssets As Long, arr_varMasterAsset() As Variant
Private lngArchiveRecs As Long
Private blnFound As Boolean, blnFound2 As Boolean
Private arr_varRetVal As Variant
Private lngX As Long, lngY As Long, lngZ As Long, lngE As Long
Private blnContinue As Boolean
'gdblCrtRpt_CostTot  ' ** Borrowing this for the sweep MarketValueCurrent.
'gstrCrtRpt_Version  ' ** Borrowing this for the sweep asset description.

' ** Progress bar variables.
Private strSp As String
Private dblPB_Steps As Double, dblPB_StepSubs As Double
Private dblPB_Width As Double, dblPB_ThisWidth As Double
Private dblPB_ThisStep As Double, dblPB_ThisStepSub As Double
Private arr_dblPB_ThisIncr() As Double, dblPB_ThisIncrSub As Double, strPB_ThisPct As String

' ** Array: arr_varOldFile().
Private Const F_ELEMS As Integer = 11  ' ** Array's first-element UBound().
Private Const F_FNAM    As Integer = 0
Private Const F_PTHFIL  As Integer = 1
Private Const F_DATA    As Integer = 2
Private Const F_CONV    As Integer = 3
Private Const F_TA_VER  As Integer = 4
Private Const F_ACC_VER As Integer = 5
Private Const F_TBLS    As Integer = 6
Private Const F_T_ARR   As Integer = 7
Private Const F_M_VER   As Integer = 8
Private Const F_APPVER  As Integer = 9
Private Const F_APPDATE As Integer = 10
Private Const F_NOTE    As Integer = 11

' ** Array: arr_varOldTbl().
Private Const T_ELEMS As Integer = 3  ' ** Array's first-element UBound().
Private Const T_TNAM  As Integer = 0
Private Const T_TNAMN As Integer = 1
Private Const T_FLDS  As Integer = 2
Private Const T_F_ARR As Integer = 3

' ** Array: arr_varOldFld().
Private Const FD_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const FD_FNAM As Integer = 0
Private Const FD_TYP  As Integer = 1
Private Const FD_SIZ  As Integer = 2

' ** Array: arr_varAcct().
Private Const A_ELEMS As Integer = 10  ' ** Array's first-element UBound().
Private Const A_NUM     As Integer = 0
Private Const A_NUM_N   As Integer = 1
Private Const A_NAM     As Integer = 2
Private Const A_TYP     As Integer = 3
Private Const A_ADMIN   As Integer = 4
Private Const A_ADMIN_N As Integer = 5
Private Const A_SCHED   As Integer = 6
Private Const A_SCHED_N As Integer = 7
Private Const A_DROPPED As Integer = 8
Private Const A_ACCT99  As Integer = 9
Private Const A_DASTNO  As Integer = 10

' ** Array: arr_varRecurType().
'Private Const RT_CODE As Integer = 0
Private Const RT_DESC As Integer = 1
'Private Const RT_JTYP As Integer = 2

' ** Array: arr_varSchedule().
Private Const S_ELEMS As Integer = 2  ' ** Arry's first-element UBound().
Private Const S_ID_OLD As Integer = 0
Private Const S_ID_NEW As Integer = 1
Private Const S_DETS   As Integer = 2

' ** Array: arr_varRevCode().
Private Const R_ELEMS As Integer = 10  ' ** Array's first-element UBound().
Private Const R_REC As Integer = 0
Private Const R_ID  As Integer = 1
Private Const R_DSC As Integer = 2
Private Const R_TYP As Integer = 3
Private Const R_ORD As Integer = 4
Private Const R_ACT As Integer = 5
Private Const R_NSO As Integer = 6  ' ** New Sort Order.
Private Const R_NID As Integer = 7  ' ** New ID.
Private Const R_EIM As Integer = 8  ' ** Element# It Matches.
Private Const R_DEL As Integer = 9
Private Const R_FND As Integer = 10

' ** Array: arr_varRevDefCode().
Private Const RD_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const RD_ID_OLD As Integer = 0
Private Const RD_ID_NEW As Integer = 1
Private Const RD_DSC    As Integer = 2

' ** Array: arr_varTaxDefCode().
'Private Const TD_ID_OLD As Integer = 0
'Private Const TD_ID_NEW As Integer = 1
'Private Const TD_DSC    As Integer = 2
'Private Const TD_TYP    As Integer = 3

' ** Array: arr_varInvestObj().
Private Const IO_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const IO_ID  As Integer = 0
Private Const IO_NAM As Integer = 1
Private Const IO_NEW As Integer = 2

' ** Array: arr_varAcctType(), arr_varAssetType().
Private Const AT_TYP As Integer = 0
'Private Const AT_DSC As Integer = 1

' ** Array: arr_varJType().
'Private Const JT_TYP  As Integer = 0
'Private Const JT_DSC  As Integer = 1
'Private Const JT_SORT As Integer = 2

' ** Array: arr_varMasterAsset().
Private Const MA_ELEMS As Integer = 6  ' ** Array's first-element UBound().
Private Const MA_OLD_ANO As Integer = 0
Private Const MA_NEW_ANO As Integer = 1
Private Const MA_NAM     As Integer = 2
Private Const MA_OLD_MVC As Integer = 3
Private Const MA_NEW_MVC As Integer = 4
Private Const MA_ERR     As Integer = 5
Private Const MA_ERRDESC As Integer = 6

' ** Array: arr_varStat().
Private Const STAT_ELEMS As Integer = 3  ' ** Array's first-element UBound().
Private Const STAT_ORD As Integer = 0
Private Const STAT_NAM As Integer = 1
Private Const STAT_CNT As Integer = 2
Private Const STAT_DSC As Integer = 3

' ** Array: arr_varBadName().
Private Const BN_ELEMS As Integer = 3  ' ** Array's first-element UBound().
Private Const BN_BAD   As Integer = 0
Private Const BN_GOOD  As Integer = 1
Private Const BN_FILE  As Integer = 2
Private Const BN_TABLE As Integer = 3
' **

Public Function Version_Upgrade_01() As Integer
' ** This begins by initializing everything, then sending it to each of the Version_Upgrade_{nn}()
' ** functions in succession, processing each of their responses, and reporting their errors.
' **
' ** Conversion is broken up into sections because, for the FIRST time when working with Access,
' ** I got the message that the procedure was too big! It exceeded the maximum size of 64K!

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Upgrade_01"

        Dim strPath As String
        Dim lngTmp01 As Long, arr_varTmp02 As Variant, lngTmp03 As Long, arr_varTmp04 As Variant, lngTmp05 As Long, arr_varTmp06 As Variant
        Dim intRetVal2 As Integer
        Dim intRetVal As Integer

110     intRetVal = 0
120     blnContinue = True

130     If gblnDev_NoErrHandle = True Then
140   On Error GoTo 0
150     End If

        ' ** This Public variable will be used to pass True/False to frmVersion_Main.
160     gblnMessage = True

170     dblPB_Steps = 0#: dblPB_StepSubs = 0#: dblPB_Width = 0#: dblPB_ThisWidth = 0#
180     dblPB_ThisStep = 0#: dblPB_ThisStepSub = 0#: dblPB_ThisIncrSub = 0#
190     strPB_ThisPct = vbNullString

200     lngErrNum = 0&: lngErrLine = 0&: strErrDesc = vbNullString
210     lngLedgerEmptyDels = 0&: lngLedgerArchEmptyDels = 0&

220     lngDupeUnks = 0&
230     ReDim arr_varDupeUnk(DU_ELEMS, 0)
240     lngDupeNum = 0&

250     strTmp04 = vbNullString: strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
260     strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
270     blnTmp22 = False: blnTmp23 = False: blnTmp24 = False: blnTmp25 = False: blnTmp26 = False
280     arr_varTmp02 = Empty: arr_varTmp03 = Empty: lngTmp13 = 0&: lngTmp14 = 0&: varTmp00 = Empty
290     strAcct99_IncomeOU = vbNullString: strAcct99_Suspense = vbNullString: blnAcct99_Both = False
300     strOldVersion = vbNullString: lngVerCnvID = -1&
310     lngArchiveRecs = 0&
320     strTruncatedFields = vbNullString

        ' ** This is currently set up for only a single bad name.
330     lngBadNames = 0&
340     ReDim arr_varBadName(BN_ELEMS, 0)

350     lngBadNames = lngBadNames + 1&
360     lngE = lngBadNames - 1&
370     ReDim Preserve arr_varBadName(BN_ELEMS, lngE)
        ' ************************************************
        ' ** Array: arr_varBadName()
        ' **
        ' **   Element  Name                  Constant
        ' **   =======  ====================  ==========
        ' **      0     Wrong Field Name      BN_BAD
        ' **      1     Correct Field name    BN_GOOD
        ' **      2     Database Name         BN_FILE
        ' **      3     Table Name            BN_TABLE
        ' **
        ' ************************************************
380     arr_varBadName(BN_BAD, lngE) = "revcode_KD"
390     arr_varBadName(BN_GOOD, lngE) = "revcode_ID"
400     arr_varBadName(BN_FILE, lngE) = "TrstArch.mdb"
410     arr_varBadName(BN_TABLE, lngE) = "ledger"
420     blnBadName = False

430     lngStats = 0&
440     ReDim arr_varStat(STAT_ELEMS, 0)

        ' ******************************************************************
        ' ******************************************************************
        ' ** IN THE ALPHA AND RELEASE VERSIONS, USE gstrTrustDataLocation!
        ' ** THIS WILL PUT \Convert_New\ IN THE NETWORK LOCATION.
450     strTmp04 = CurrentAppPath  ' ** Module Function: modFileUtilities.
460     If Left(strTmp04, 37) = Left(gstrDir_Dev, 37) Then        ' ** It's one of my test directories.
470       If Left(strTmp04, 47) = gstrDir_Dev Then                 ' ** If it's in my main test directory...
            ' ** "C:\VictorGCS_Clients\TrustAccountant\NewWorking"  '## OK
480         strPath = (CurrentAppPath & LNK_SEP & gstrDir_Convert)  ' ** Module Function: modFileUtilities.
490       ElseIf Parse_File(Left(strTmp04, 44)) = "NewDemo" Then  ' ** Module Function: modFileUtilities.
500         strPath = (CurrentAppPath & LNK_SEP & gstrDir_Convert)  ' ** Module Function: modFileUtilities.
510       ElseIf Parse_File(Left(strTmp04, 44)) = "Clients" And _
              Right(strTmp04, 27) = "Delta Data\Trust Accountant" Then
520         strPath = (CurrentAppPath & LNK_SEP & "Database" & LNK_SEP & gstrDir_Convert)
530       Else
540         Beep
550         MsgBox "Where am I?", vbCritical + vbOKOnly, "Where Am I?"
560       End If
570     Else                                                     ' ** Otherwise...
580       strPath = (gstrTrustDataLocation & gstrDir_Convert)     ' ** gstrTrustDataLocation INCLUDES FINAL BACKSLASH!
590     End If
        ' ******************************************************************
        ' ******************************************************************

        ' ******************************************************************
        ' ** Version_Upgrade_01() {this}
        ' **   Initialization of variables.
        ' **   Calls and processes the response from other sections.

        ' ** Version_Upgrade_02()
        ' **   Step 0:  Determine whether conversion is necessary.

        ' ** Version_Upgrade_03()  :  Below.
        ' **   Step 1:  Beginning. (Connect, get list of tables, get list of accounts.)
        ' **   Step 2:  CompanyInformation
        ' **   Step 3:  adminofficer
        ' **   Step 4:  Location
        ' **   Step 5:  RecurringItems
        ' **   Step 6:  Schedule, ScheduleDetail
        ' **   Step 7:  DDate, PostingDate, Statement Date
        ' **   Step 8:  m_REVCODE
        ' **   Step 9:  Users
        ' **   Step 10: _~xusr
        ' **   Step 11: tblSecurity_License

        ' ** Version_Upgrade_04()  :  modVersionConvertFuncs3
        ' **   Step 12: account
        ' **   Step 13: Balance
        ' **   Step 14: masterasset
        ' **   Step 15: ActiveAssets

        ' ** Version_Upgrade_05()  :  modVersionConvertFuncs2
        ' **   Step 16: ledger
        ' **   Step 17: journal
        ' **   Step 18: LedgerHidden

        ' ** Version_Upgrade_06()  :  modVersionConvertFuncs2
        ' **   Step 19: LedgerArchive
        ' **   Step 20: tblPricing_MasterAsset_History (tblAssetPricing)
        ' **   Step 21: tblJournal_Memo

        ' ** Version_Upgrade_07()  :  modVersionConvertFuncs2
        ' **   Step 22: tblCurrency
        ' **   Step 23: tblCurrency_History
        ' **   Step 24: tblCurrency_Account
        ' **   Step 25: tblLedgerHidden

        ' ** Version_Upgrade_08()  :  modVersionConvertFuncs3
        ' **   Step 26: tblCheckMemo
        ' **   Step 27: tblCheckReconcile_Account
        ' **   Step 28: tblCheckReconcile_Item
        ' **   Step 29: tblCheckPOSPay
        ' **   Step 30: tblCheckPOSPay_Detail
        ' **   Step 31: tblCheckBank
        ' **   Step 32: tblCheckVoid
        ' **   Step 33: tblRecurringAux1099  (not used yet)

        ' In ("CompanyInformation","adminofficer","Location","RecurringItems","Schedule","ScheduleDetail","DDate","PostingDate","Statement Date","m_REVCODE","Users","_~xusr","tblSecurity_License","account","Balance","masterasset","ActiveAssets","ledger","journal","LedgerHidden","LedgerArchive","tblPricing_MasterAsset_History","tblJournal_Memo","tblCurrency","tblCurrency_History","tblCurrency_Account","tblLedgerHidden","tblCheckMemo","tblCheckReconcile_Account","tblCheckReconcile_Item","tblCheckPOSPay","tblCheckPOSPay_Detail","tblCheckBank","tblCheckVoid","tblRecurringAux1099")

        ' ** Version_Upgrade_01()
        ' **   Step 34: Rename files in \Convert_New\
        ' ******************************************************************

        ' ** Get list of files in \Convert_New\ (if it's there),
        ' ** get version being converted, and open frmVersion_Main.
600     intRetVal2 = Version_Upgrade_02(strPath)  ' ** Function: Below.
        'strOldVersion, SET ABOVE, WILL BE NEEDED IN Version_DataXFer() IN modVersionConvertFuncs2!
        ' ** Return values:
        ' **    1  Unnecessary
        ' **    0  OK
        ' **   -1  Can't Connect
        ' **   -2  Can't Open
        ' **   -3  Canceled Status
        ' **   -9  Error

        ' ** The 1st If/EndIf block in Version_Upgrade_03() looks for the \Convert_New\ directory in the data location.
        ' ** If it doesn't find it, the Wise installation wasn't an Upgrade version.
        ' ** Between the 1st and 2nd If/EndIf blocks, it looks for .BAK files in the \Convert_New\ directory.
        ' ** The 2nd If/EndIf block is run if no .BAK files were found.

        ' ** In all cases, blnRetVal may return True.
        ' ** It's blnContinue that prevents or allows it to go into successive If/EndIf blocks.
        ' ** That means a non-conversion will still drop down into both the following procedures, and that's OK.

        ' ** blnContinue may be False! That can cause a problem with Version_Status()!
610     If intRetVal2 = 0 And blnContinue = True Then
          ' ** At this point, frmVersion_Main should be open.

620       intRetVal2 = Version_Upgrade_03  ' ** Function: Below.
          ' ** Return values:
          ' **    0  OK
          ' **   -4  Acount Empty
          ' **   -5  Canceled CoInfo
          ' **   -6  Index/Key
          ' **   -9  Error

630       If intRetVal2 = 0 And blnContinue = True Then
            ' ** frmVersion_Main is still open.

640         lngTmp01 = lngInvestObjs
650         arr_varTmp02 = arr_varInvestObj

660         lngTmp03 = lngStats
670         arr_varTmp04 = arr_varStat

680         lngTmp05 = lngDupeUnks
690         arr_varTmp06 = arr_varDupeUnk

700         intRetVal2 = Version_Upgrade_04(blnContinue, blnConvert_TrustDta, lngAccts, arr_varAcct, lngAcctTypes, arr_varAcctType, _
              lngAssetTypes, arr_varAssetType, lngMasterAssets, lngTmp01, arr_varTmp02, lngTmp03, arr_varTmp04, lngTmp05, arr_varTmp06, _
              lngOldFiles, arr_varOldFile, lngOldTbls, arr_varOldTbl, dblPB_ThisStep, lngTrustDtaDbsID, strKeyTbl, strTruncatedFields, _
              strAcct99_IncomeOU, strAcct99_Suspense, lngArchElem, lngDupeNum, wrkLnk, dbsLnk, wrkLoc, dbsLoc)  ' ** Module Function: modVersionConvertFuncs3.
            ' ** Return values:
            ' **    0  OK
            ' **   -6  Index/Key
            ' **   -9  Error

710         lngRecs = lngTmp01
720         For lngX = 0& To (lngRecs - 1&)
730           lngInvestObjs = lngInvestObjs + 1&
740           lngE = lngInvestObjs - 1&
750           ReDim Preserve arr_varInvestObj(IO_ELEMS, lngE)
760           For lngY = 0& To IO_ELEMS
770             arr_varInvestObj(lngY, lngE) = arr_varTmp02(lngY, lngX)
780           Next
790         Next

800         lngStats = 0&
810         ReDim arr_varStat(STAT_ELEMS, 0)

820         lngRecs = lngTmp03
830         For lngX = 0& To (lngRecs - 1&)
840           lngStats = lngStats + 1&
850           lngE = lngStats - 1&
860           ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
870           For lngY = 0& To STAT_ELEMS
880             arr_varStat(lngY, lngE) = arr_varTmp04(lngY, lngX)
890           Next
900         Next

910         lngDupeUnks = 0&
920         ReDim arr_varDupeUnk(DU_ELEMS, 0)

930         lngRecs = lngTmp05
940         For lngX = 0& To (lngRecs - 1&)
950           lngDupeUnks = lngDupeUnks + 1&
960           lngE = lngDupeUnks - 1&
970           ReDim Preserve arr_varDupeUnk(DU_ELEMS, lngE)
980           For lngY = 0& To DU_ELEMS
990             arr_varDupeUnk(lngY, lngE) = arr_varTmp06(lngY, lngX)
1000          Next
1010        Next

1020        If intRetVal2 = 0 And blnContinue = True Then

              'intRetVal2 = Version_Upgrade_05  ' ** Function: Below.
1030          intRetVal2 = Version_Upgrade_05(blnContinue, blnConvert_TrustDta, lngTrustDtaDbsID, lngLedgerEmptyDels, _
                lngDupeNum, lngAccts, arr_varAcct, lngRevCodes, arr_varRevCode, lngTaxDefCodes, arr_varTaxDefCode, _
                lngDupeUnks, arr_varDupeUnk, lngBadNames, arr_varBadName, lngOldFiles, arr_varOldFile, lngOldTbls, _
                arr_varOldTbl, lngStats, arr_varStat, dblPB_ThisStep, strKeyTbl, wrkLoc, wrkLnk, dbsLoc, dbsLnk)  ' ** Module Function: modVersionConvertFuncs2.
              ' ** Return values:
              ' **    0  OK
              ' **   -6  Index/Key
              ' **   -9  Error

1040          If intRetVal2 = 0 And blnContinue = True Then

                'intRetVal2 = Version_Upgrade_06  ' ** Function: Below.
1050            intRetVal2 = Version_Upgrade_06(blnContinue, blnConvert_TrstArch, intWrkType, lngDtaElem, lngArchElem, _
                  lngTrustDtaDbsID, lngTrstArchDbsID, lngAccts, arr_varAcct, lngRevCodes, arr_varRevCode, _
                  lngTaxDefCodes, arr_varTaxDefCode, lngBadNames, arr_varBadName, lngOldFiles, arr_varOldFile, _
                  lngStats, arr_varStat, dblPB_ThisStep, strKeyTbl, wrkLoc, dbsLoc)  ' ** Module Function: modVersionConvertFuncs2.

                ' ** Return values:
                ' **    0  OK
                ' **   -6  Index/Key
                ' **   -7  Can't Open
                ' **   -9  Error

1060            If intRetVal2 = 0 And blnContinue = True Then

                  'intRetVal2 = Version_Upgrade_07  ' ** Function: Below.
1070              intRetVal2 = Version_Upgrade_07(blnContinue, blnConvert_TrustDta, lngTrustDtaDbsID, strKeyTbl, dblPB_ThisStep, _
                    lngOldTbls, arr_varOldTbl, lngAccts, arr_varAcct, lngStats, arr_varStat, wrkLoc, wrkLnk, dbsLoc, dbsLnk)  ' ** Module Function: modVersionConvertFuncs2.
                  ' ** Return values:
                  ' **    0  OK
                  ' **   -6  Index/Key
                  ' **   -7  Can't Open
                  ' **   -9  Error

1080              If intRetVal2 = 0 And blnContinue = True Then

1090                intRetVal2 = Version_Upgrade_08(blnContinue, blnConvert_TrustDta, lngTrustDtaDbsID, strKeyTbl, dblPB_ThisStep, _
                      lngOldTbls, arr_varOldTbl, lngAccts, arr_varAcct, lngStats, arr_varStat, wrkLoc, wrkLnk, dbsLoc, dbsLnk)  ' ** Module Function: modVersionConvertFuncs3.
                    ' ** Return values:
                    ' **    0  OK
                    ' **   -6  Index/Key
                    ' **   -7  Can't Open
                    ' **   -9  Error

1100                If intRetVal2 = 0 And blnContinue = True Then

                      ' ******************************
                      ' ** Rename Files In Convert_New
                      ' ******************************

                      ' ** Step 34: Rename files.
1110                  Version_Up1_Etc strPath, dblPB_Steps, dblPB_ThisStep, 1  ' ** Module Procedure: modVersionConvertFuncs2.

1120                End If

1130              End If

1140            End If

1150          End If

1160        End If

1170      End If
1180    Else
          ' ** This is in an effort to prevent that 'Subscript out of range' I got on a very old version.
          ' ** I think I've covered that now in Version_GetOldVer().
1190      If blnContinue = False And intRetVal2 = 0 Then
1200        intRetVal2 = -9
1210      End If
          ' ** intRetVal2 will be either negative or positive, but 0 continues into the block above.
1220      If intRetVal2 > 0 Then
            ' ** Positive: No conversion, and no open window.
1230        intRetVal = intRetVal2
1240      Else
            ' ** Negative: Problem.
1250        If intRetVal2 = -9 Then intRetVal2 = -10  ' ** Message below slightly different for opening error.
1260      End If
1270    End If

1280    If intRetVal2 < 0 Then
          ' ** Return values:
          ' **    1  Unnecessary
          ' **    0  OK
          ' **   -1  Can't Connect
          ' **   -2  Can't Open {TrustDta.mdb}
          ' **   -3  Canceled Status
          ' **   -4  Acount Empty
          ' **   -5  Canceled CoInfo
          ' **   -6  Index/Key
          ' **   -7  Can't Open {TrstArch.mdb}
          ' **   -8  Tables Not Empty
          ' **   -9  Error

          ' ** NOTE: THAT ODD LONG MsgBox() ABOUT FINDING BOTH MDB AND BAK IS IN Version_Upgrade_02(), BELOW.

1290      intRetVal = intRetVal2
1300      DoCmd.Hourglass False
1310      Beep
1320      Select Case intRetVal2
          Case -1
            ' **   -1  Can't Connect
1330        If lngVerCnvID > 0& Then
1340          Set dbsLoc = CurrentDb
1350          With dbsLoc
                ' ** Update tblVersion_Conversion, set vercnv_error = True, by specified [vercid], [vernot].
1360            Set qdf = .QueryDefs("qryVersion_Convert_07")
1370            With qdf.Parameters
1380              ![vercid] = lngVerCnvID
1390              ![vernot] = "Workspace Failed"
1400            End With  ' ** Parameters.
1410            qdf.Execute
1420            .Close
1430          End With  ' ** dbsLoc.
1440        End If
1450        strFailMsg = "Trust Accountant has detected that this is new installation," & vbCrLf & _
              "with data from an earlier version awaiting conversion." & vbCrLf & vbCrLf & _
              "However, Trust Accountant was unable to establish a workable connection to your old data." & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc., for assistance." & vbCrLf & vbCrLf & _
              "Step: " & CStr(dblPB_ThisStep)
1460        MsgBox strFailMsg, vbCritical + vbOKOnly, ("Workspace Failed At Step " & CStr(dblPB_ThisStep) & Space(40))
1470        If IsLoaded(FRM_CNV_STATUS, acForm) = True Then  ' ** Module Function: modFileUtilities.
              ' ** Shouldn't be open.
1480          DoCmd.Close acForm, FRM_CNV_STATUS
1490        End If
1500      Case -2, -7
            ' **   -2  Can't Open (TrustDta.mdb)
            ' **   -7  Can't Open (TrstArch.mdb)
1510        If lngVerCnvID > 0& Then
1520          Set dbsLoc = CurrentDb
1530          With dbsLoc
                ' ** Update tblVersion_Conversion, set vercnv_error = True, by specified [vercid], [vernot].
1540            Set qdf = .QueryDefs("qryVersion_Convert_07")
1550            With qdf.Parameters
1560              ![vercid] = lngVerCnvID
1570              ![vernot] = "Database Failed: " & IIf(intRetVal2 = -2, "TrustDta.mdb", "TrstArch.mdb")
1580            End With  ' ** Parameters.
1590            qdf.Execute
1600            .Close
1610          End With  ' ** dbsLoc.
1620        End If
1630        strFailMsg = "Trust Accountant has detected that this is new installation," & vbCrLf & _
              "with data from an earlier version awaiting conversion." & vbCrLf & vbCrLf & _
              "However, Trust Accountant was unable to open your old data file." & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc., for assistance." & vbCrLf & vbCrLf & _
              "Step: " & CStr(dblPB_ThisStep)
1640        strTmp04 = IIf(intRetVal2 = -2, "TrustDta.mdb", "TrstArch.mdb")
1650        MsgBox strFailMsg, vbCritical + vbOKOnly, _
              (strTmp04 & " Database Failed At Step " & CStr(dblPB_ThisStep) & Space(40))
1660        strFailMsg = strFailMsg & "  " & strTmp04
1670        If IsLoaded(FRM_CNV_STATUS, acForm) = True Then  ' ** Module Function: modFileUtilities.
              ' ** Shouldn't be open.
1680          DoCmd.Close acForm, FRM_CNV_STATUS
1690        End If
1700      Case -3
            ' **   -3  Canceled Status
1710        Set dbsLoc = CurrentDb
1720        With dbsLoc
              ' ** Update tblVersion_Conversion, set vercnv_cancel = True, by specified [vercid].
1730          Set qdf = .QueryDefs("qryVersion_Convert_06")
1740          With qdf.Parameters
1750            ![vercid] = lngVerCnvID
1760          End With  ' ** Parameters.
1770          qdf.Execute
1780          .Close
1790        End With  ' ** dbsLoc.
1800        strFailMsg = "You chose to cancel the conversion process." & vbCrLf & vbCrLf & _
              "Your old files remain unchanged, and no data" & vbCrLf & _
              "was transferred to your current Trust Accountant." & vbCrLf & vbCrLf & _
              "Step: " & CStr(dblPB_ThisStep)
1810        MsgBox strFailMsg, vbExclamation + vbOKOnly, ("Conversion Canceled At Step " & CStr(dblPB_ThisStep) & Space(40))
1820        If IsLoaded(FRM_CNV_STATUS, acForm) = True Then  ' ** Module Function: modFileUtilities.
              ' ** Shouldn't be open.
1830          DoCmd.Close acForm, FRM_CNV_STATUS
1840        End If
1850      Case -4
            ' **   -4  Account Empty
1860        Set dbsLoc = CurrentDb
1870        With dbsLoc
              ' ** Update tblVersion_Conversion, set vercnv_error = True, by specified [vercid], [vernot].
1880          Set qdf = .QueryDefs("qryVersion_Convert_07")
1890          With qdf.Parameters
1900            ![vercid] = lngVerCnvID
1910            ![vernot] = "Account Table Empty"
1920          End With  ' ** Parameters.
1930          qdf.Execute
1940          .Close
1950        End With  ' ** dbsLoc.
1960        strFailMsg = "The primary Trust Accountant table in your old data file is empty!" & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc., for assistance." & vbCrLf & vbCrLf & _
              "Step: " & CStr(dblPB_ThisStep)
1970        MsgBox strFailMsg, vbCritical + vbOKOnly, ("Conversion Failed At Step " & CStr(dblPB_ThisStep) & Space(40))
1980        Forms(FRM_CNV_STATUS).cmdCancel.Enabled = True
1990        Forms(FRM_CNV_STATUS).cmdCancel.Caption = "&Close"
2000        Forms(FRM_CNV_STATUS).cmdCancel.SetFocus
2010      Case -5
2020        Set dbsLoc = CurrentDb
2030        With dbsLoc
              ' ** Update tblVersion_Conversion, set vercnv_cancel = True, by specified [vercid].
2040          Set qdf = .QueryDefs("qryVersion_Convert_06")
2050          With qdf.Parameters
2060            ![vercid] = lngVerCnvID
2070          End With  ' ** Parameters.
2080          qdf.Execute
2090          .Close
2100        End With  ' ** dbsLoc.
            ' **   -5  Canceled CoInfo
2110        strFailMsg = "You chose to cancel the conversion process." & vbCrLf & vbCrLf & _
              "Your old files remain unchanged, and no data" & vbCrLf & _
              "was transferred to your current Trust Accountant." & vbCrLf & vbCrLf & _
              "Step: " & CStr(dblPB_ThisStep)
2120        MsgBox strFailMsg, vbExclamation + vbOKOnly, ("Conversion Canceled At Step " & CStr(dblPB_ThisStep) & Space(40))
2130        If IsLoaded(FRM_CNV_STATUS, acForm) = True Then  ' ** Module Function: modFileUtilities.
2140          DoCmd.Close acForm, FRM_CNV_STATUS
2150        End If
2160      Case -6
            ' **   -6  Index/Key
2170        Set dbsLoc = CurrentDb
2180        With dbsLoc
              ' ** Update tblVersion_Conversion, set vercnv_error = True, by specified [vercid], [vernot].
2190          Set qdf = .QueryDefs("qryVersion_Convert_07")
2200          With qdf.Parameters
2210            ![vercid] = lngVerCnvID
2220            ![vernot] = "Table Update Failed; Error: " & CStr(lngErrNum) & "; Line: " & CStr(lngErrLine) & "; " & strErrDesc
2230          End With  ' ** Parameters.
2240          qdf.Execute
2250          .Close
2260        End With  ' ** dbsLoc.
2270        strFailMsg = "Because of the error, Trust Accountant is unable to continue with the conversion." & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc., for assistance." & vbCrLf & vbCrLf & _
              "Step: " & CStr(dblPB_ThisStep)
2280        MsgBox strFailMsg, vbCritical + vbOKOnly, ("Conversion Failed At Step " & CStr(dblPB_ThisStep) & Space(40))
2290        Forms(FRM_CNV_STATUS).cmdCancel.Enabled = True
2300        Forms(FRM_CNV_STATUS).cmdCancel.Caption = "&Close"
2310        Forms(FRM_CNV_STATUS).cmdCancel.SetFocus
2320      Case -8
            ' ** Tables Not Empty
2330        Set dbsLoc = CurrentDb
2340        With dbsLoc
              ' ** Update tblVersion_Conversion, set vercnv_error = True, by specified [vercid], [vernot].
2350          Set qdf = .QueryDefs("qryVersion_Convert_07")
2360          With qdf.Parameters
2370            ![vercid] = lngVerCnvID
2380            ![vernot] = "Tables Not Empty"
2390          End With  ' ** Parameters.
2400          qdf.Execute
2410          .Close
2420        End With  ' ** dbsLoc.
2430        strFailMsg = "Though a conversion is indicated, there is already data present in Trust Accountant." & vbCrLf & vbCrLf & _
              "If you wish to do the conversion over again, please contact Delta Data, Inc., for assistance." & vbCrLf & vbCrLf & _
              strTmp04 & vbCrLf & vbCrLf & "Step: " & CStr(dblPB_ThisStep)
2440        MsgBox strFailMsg, vbExclamation + vbOKOnly, ("Conversion Stopped At Step " & CStr(dblPB_ThisStep) & Space(40))
2450        If IsLoaded(FRM_CNV_STATUS, acForm) = True Then  ' ** Module Function: modFileUtilities.
2460          DoCmd.Close acForm, FRM_CNV_STATUS
2470        End If
2480      Case -9
            ' **   -9  Error
2490        Set dbsLoc = CurrentDb
2500        With dbsLoc
              ' ** Update tblVersion_Conversion, set vercnv_error = True, by specified [vercid], [vernot].
2510          Set qdf = .QueryDefs("qryVersion_Convert_07")
2520          With qdf.Parameters
2530            ![vercid] = lngVerCnvID
2540            strErrDesc = Left(strErrDesc, InStr(strErrDesc, "."))
2550            ![vernot] = ("Error: " & CStr(lngErrNum) & "; Line: " & CStr(lngErrLine) & "; " & strErrDesc)
2560          End With  ' ** Parameters.
2570          qdf.Execute
2580          .Close
2590        End With  ' ** dbsLoc.
2600        strFailMsg = "Because of the error, Trust Accountant is unable to continue with the conversion." & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc., for assistance." & vbCrLf & vbCrLf & _
              "Step: " & CStr(dblPB_ThisStep)
2610        MsgBox strFailMsg, vbCritical + vbOKOnly, ("Conversion Failed At Step " & CStr(dblPB_ThisStep) & Space(40))
2620        Forms(FRM_CNV_STATUS).cmdCancel.Enabled = True
2630        Forms(FRM_CNV_STATUS).cmdCancel.Caption = "&Close"
2640        Forms(FRM_CNV_STATUS).cmdCancel.SetFocus
2650      Case -10
            ' **   -9  Error --> -10
2660        Set dbsLoc = CurrentDb
2670        With dbsLoc
              ' ** Update tblVersion_Conversion, set vercnv_error = True, by specified [vercid], [vernot].
2680          Set qdf = .QueryDefs("qryVersion_Convert_07")
2690          With qdf.Parameters
2700            ![vercid] = lngVerCnvID
2710            strErrDesc = Left(strErrDesc, InStr(strErrDesc, "."))
2720            ![vernot] = ("Error: " & CStr(lngErrNum) & "; Line: " & CStr(lngErrLine) & "; " & strErrDesc)
2730          End With  ' ** Parameters.
2740          qdf.Execute
2750          .Close
2760        End With  ' ** dbsLoc.
2770        strFailMsg = "Trust Accountant has detected that this is new installation," & vbCrLf & _
              "with data from an earlier version awaiting conversion." & vbCrLf & vbCrLf & _
              "However, because of the error, it was unable to continue with the analysis." & vbCrLf & vbCrLf & _
              "Please contact Delta Data, Inc., for assistance." & vbCrLf & vbCrLf & _
              "Step: " & CStr(dblPB_ThisStep)
2780        MsgBox strFailMsg, vbCritical + vbOKOnly, ("Error During Detection Phase At Step " & CStr(dblPB_ThisStep) & Space(40))
2790        If IsLoaded(FRM_CNV_STATUS, acForm) = True Then  ' ** Module Function: modFileUtilities.
              ' ** Shouldn't be open.
2800          DoCmd.Close acForm, FRM_CNV_STATUS
2810        End If
2820      End Select  ' ** intRetVal2.

2830    End If  ' ** intRetVal2 < 0.

        ' ** Reset, lest it interfere with other places I'm using it!
2840    gblnMessage = False

2850    DoCmd.Hourglass False

EXITP:
2860    Set fsfl = Nothing
2870    Set fsfls = Nothing
2880    Set fsfd = Nothing
2890    Set fso = Nothing
2900    Set prp = Nothing
2910    Set doc = Nothing
2920    Set fld = Nothing
2930    Set tdf = Nothing
2940    Set rstLnk = Nothing
2950    Set dbsLnk = Nothing
2960    Set wrkLnk = Nothing
2970    Set rstLoc1 = Nothing
2980    Set rstLoc2 = Nothing
2990    Set rstLoc3 = Nothing
3000    Set qdf = Nothing
3010    Set dbsLoc = Nothing
3020    Set wrkLoc = Nothing
3030    Version_Upgrade_01 = intRetVal
3040    Exit Function

ERRH:
3050    intRetVal = -9
3060    DoCmd.Hourglass False
3070    lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
3080    Select Case ERR.Number
        Case Else
3090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3100    End Select
3110    Resume EXITP

End Function

Public Function Version_Upgrade_02(strPath As String) As Integer
' ** This looks for files in the Convert_New directory, and decides whether conversion is warranted.
' ** When this is called, frmVersion_Main is not yet open.
' **
' ** Table IDs recorded by tblVersion_Key:
' **   adminofficer
' **   Location
' **   RecurringItems
' **   Schedule
' **   m_REVCODE
' **   masterasset
' **
' ** Return values:
' **    1  Unnecessary
' **    0  OK
' **   -1  Can't Connect
' **   -2  Can't Open
' **   -3  Canceled Status
' **   -8  Tables Not Empty
' **   -9  Error

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Upgrade_02"

        Dim blnCheckForBoth_TrustDta As Boolean, blnCheckForBoth_TrstArch As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal1 As Integer, intRetVal2 As Integer, intRetVal3 As Integer

3210    If gblnDev_NoErrHandle = True Then
3220  On Error GoTo 0
3230    End If

3240    intRetVal1 = 0

3250    DoCmd.Hourglass True

3260    blnConvert_TrustDta = False: blnConvert_TrstArch = False: blnArchiveNotPresent = False

3270    If gstrTrustDataLocation = vbNullString Then
3280      IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
3290    End If

3300    If DirExists(strPath) Then

          ' ** The Wise Installer Script Editor first created a temporary Trust
          ' ** Accountant directory, and copied the existing files to it.
          ' **   Create Directory %TEMP%\Trust_bak
          ' **   Copy local file from %TRUSTDTA%\TrustDta.mdb to %TEMP%\Trust_bak\TrustDta.mdb
          ' **   Copy local file from %TRSTARCH%\TrstArch.mdb to %TEMP%\Trust_bak\TrstArch.mdb
          ' **   Copy local file from %TRUSTSEC%\TrustSec.mdw to %TEMP%\Trust_bak\TrustSec.mdw
          ' **   Copy local file from %TALIC%\TA.lic to %TEMP%\Trust_bak\TA.lic
          ' **   Copy local file from %TRUSTINI%\DDTrust.ini to %TEMP%\Trust_bak\DDTrust.ini
          ' ** Then will copy them back to C:\Program Files\Delta Data\Trust Accountant\Database\Convert_New\ after the installation.
          ' **   New, v2.1.63:
          ' **   Copy local file from %TEMP%\Trust_bak\TrustDta.mdb to %TRUSTDTA%\Convert_New\TrustDta.mdb
          ' **   Copy local file from %TEMP%\Trust_bak\TrstArch.mdb to %TRSTARCH%\Convert_New\TrstArch.mdb
          ' **   -----------------------------------------------------------------------------------------
          ' **   Previously, vb2.1.62 and earlier:
          ' **   Copy local file from %TEMP%\Trust_bak\TrstArch.BAK to %TRSTARCH%\TrstArch.BAK
          ' **   Copy local file from %TEMP%\Trust_bak\TrstArch.INST to %TRSTARCH%\TrstArch.INST
          ' **   Copy local file from %TEMP%\Trust_bak\TrustDta.BAK to %TRUSTDTA%\TrustDta.BAK
          ' **   Copy local file from %TEMP%\Trust_bak\TrustDta.INST to %TRUSTDTA%\TrustDta.INST
          ' ** These will have copied over the new, installed versions.
          ' **   Copy local file from %TEMP%\Trust_bak\TA.lic to %TALIC%\TA.lic
          ' **   Copy local file from %TEMP%\Trust_bak\TrustSec.mdw to %TRUSTSEC%\TrustSec.mdw
          ' **   Copy local file from %TEMP%\Trust_bak\DDTrust.ini to %TRUSTINI%\DDTrust.ini

          ' ** Instead, at the end, after Wise has run, we want the Wise Installer Script Editor
          ' ** to no longer run TrustUpd.mde, but just to copy the old TrustDta.mdb, as is, to:
          ' **   %TRUSTDTA%\Convert_New\
          ' ** And TrstArch.mdb, as is, to:
          ' **   %TRSTARCH%\Convert_New\
          ' ** And, as it is doing now, TrustSec.mdw, TA.lic, DDTrust.ini to:
          ' **   %TRUSTSEC%\TrustSec.mdw
          ' **   %TALIC%\TA.lic
          ' **   %TRUSTINI%\DDTrust.ini
          ' ** This should leave the old TrustDta.mdb and TrstArch.mdb in \Convert_New\.

3310      lngOldFiles = 0&
3320      ReDim arr_varOldFile(F_ELEMS, 0)

3330      Set fso = CreateObject("Scripting.FileSystemObject")
3340      With fso

3350        Set fsfd = .GetFolder(strPath)
3360        Set fsfls = fsfd.FILES

3370        lngOldFiles = fsfls.Count
3380        If lngOldFiles > 0& Then

3390          ReDim arr_varOldFile(F_ELEMS, (lngOldFiles - 1&))

              ' ** Get all the files in the Convert_New folder.
              ' ** If file extensions are still standard MDB, then convert hasn't yet taken place.
3400          lngX = -1&
3410          For Each fsfl In fsfls
3420            With fsfl
3430              lngX = lngX + 1&
3440              ReDim Preserve arr_varOldFile(F_ELEMS, lngX)
                  ' *******************************************************
                  ' ** Array: arr_varOldFile()
                  ' **
                  ' **   Element  Name                        Constant
                  ' **   =======  ==========================  ===========
                  ' **      0     Name                        F_FNAM
                  ' **      1     Path and File               F_PTHFIL
                  ' **      2     Data File (Y/N)             F_DATA
                  ' **      3     Already Converted (Y/N)     F_CONV
                  ' **      4     Trust Accountant Version    F_TA_VER
                  ' **      5     Access Version              F_ACC_VER
                  ' **      6     Tables                      F_TBLS
                  ' **      7     Table Array                 F_T_ARR
                  ' **      8     m_Vx Version                F_M_VER
                  ' **      9     AppVersion                  F_APPVER
                  ' **     10     AppDate                     F_APPDATE
                  ' **     11     Note                        F_NOTE
                  ' **
                  ' *******************************************************
3450              arr_varOldFile(F_FNAM, lngX) = .Name
3460              arr_varOldFile(F_PTHFIL, lngX) = .Path
3470              If Rem_Ext(arr_varOldFile(F_FNAM, lngX)) = Rem_Ext(gstrFile_DataName) Then  ' ** Module Function: modFileUtilities.
3480                arr_varOldFile(F_DATA, lngX) = CBool(True)
3490              ElseIf Rem_Ext(arr_varOldFile(F_FNAM, lngX)) = Rem_Ext(gstrFile_ArchDataName) Then  ' ** Module Function: modFileUtilities.
3500                arr_varOldFile(F_DATA, lngX) = CBool(True)
3510              Else
3520                arr_varOldFile(F_DATA, lngX) = CBool(False)
3530              End If
3540              arr_varOldFile(F_CONV, lngX) = CBool(False)
3550              arr_varOldFile(F_TA_VER, lngX) = vbNullString
3560              arr_varOldFile(F_ACC_VER, lngX) = vbNullString
3570              arr_varOldFile(F_TBLS, lngX) = CLng(0)
3580              arr_varOldFile(F_T_ARR, lngX) = Empty
3590              arr_varOldFile(F_M_VER, lngX) = vbNullString
3600              arr_varOldFile(F_APPVER, lngX) = vbNullString
3610              arr_varOldFile(F_APPDATE, lngX) = Null
3620              arr_varOldFile(F_NOTE, lngX) = vbNullString
3630            End With  ' ** This file: fsfl.
3640          Next  ' ** For each file: fsfl.

3650        End If
3660      End With  ' ** fso.

          ' ########################
          ' ## Provide a mechanism for renaming .BAK back to .MDB to reconvert.
          ' ########################

3670      For lngX = 0& To (lngOldFiles - 1&)
3680        If arr_varOldFile(F_DATA, lngX) = True Then
3690          If Parse_Ext(arr_varOldFile(F_FNAM, lngX)) = "BAK" Then  ' ** Module Function: modFileUtilities.
3700            arr_varOldFile(F_CONV, lngX) = True
3710          End If
3720        End If
3730      Next

          ' ** If both data files are already BAK, the conversion's been done.
          ' ** What if both BAK and MDB are found? What would that mean?
          ' ** I'd say, that means a reconversion is desired, since
          ' ** this conversion process renames the existing files.
          ' ** If that's the case, we'd have to ask the user whether to
          ' ** replace ALL the data with a new conversion, or to append it.
          ' ** Since absolutely no provision exists in Trust Accountant to
          ' ** append additional data, that'd have to be a whole new process.
          ' ** In any case, I'm not writing code for this possibility now, so
          ' ** the presence of BAK's will flag the conversion as already done.

          ' ** 08/03/2011: NEW!
          ' ** With a succession of upgrades, there may now be both .BAK's and the new .MDB's.
          ' ** Just get rid of the .BAK's by either deleting or moving them.

          ' ** First, look for BAK's.
          ' ** If the BAK is present, look no further, and don't run convert.
3740      blnConvert_TrustDta = True: blnConvert_TrstArch = True: blnArchiveNotPresent = True
3750      For lngX = 0& To (lngOldFiles - 1&)
3760        If Rem_Ext(arr_varOldFile(F_FNAM, lngX)) = Rem_Ext(gstrFile_DataName) Then
3770          lngDtaElem = lngX
3780          If Parse_Ext(arr_varOldFile(F_FNAM, lngX)) = "BAK" Then  ' ** Module Functions: modFileUtilities.
3790            blnConvert_TrustDta = False
3800          End If
3810        End If
3820        If Rem_Ext(arr_varOldFile(F_FNAM, lngX)) = Rem_Ext(gstrFile_ArchDataName) Then
3830          lngArchElem = lngX
3840          blnArchiveNotPresent = False
3850          If Parse_Ext(arr_varOldFile(F_FNAM, lngX)) = "BAK" Then  ' ** Module Functions: modFileUtilities.
3860            blnConvert_TrstArch = False
3870          End If
3880        End If
3890      Next

          ' ** If TrstArch isn't present, then there's really no conversion necessary.
          ' ** This new version has an empty one attached and all set to go.
3900      If blnArchiveNotPresent = True Then blnConvert_TrstArch = False

          ' ** Then look for MDB's.
          ' ** If there's no BAK, make sure the MDB is a real one.
3910      If blnConvert_TrustDta = True Then
3920        blnConvert_TrustDta = False
3930        For lngX = 0& To (lngOldFiles - 1&)
3940          If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
3950            blnConvert_TrustDta = True
3960            Exit For
3970          End If
3980        Next
3990      End If
4000      If blnConvert_TrstArch = True Then
4010        blnConvert_TrstArch = False
4020        For lngX = 0& To (lngOldFiles - 1&)
4030          If arr_varOldFile(F_FNAM, lngX) = gstrFile_ArchDataName Then
4040            blnConvert_TrstArch = True
4050            Exit For
4060          End If
4070        Next
4080      End If

          ' ** Because of the problem on Rich's machine, and the current
          ' ** inability of the Uninstall process to get files and folders
          ' ** put there by the WiseScript wrapper, let's check for both.
4090      If blnConvert_TrustDta = False And blnConvert_TrstArch = False Then

4100        lngTmp13 = 0&: lngTmp14 = 0&
4110        blnCheckForBoth_TrustDta = False
4120        For lngX = 0& To (lngOldFiles - 1&)
4130          If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
4140            lngTmp13 = lngX
4150            blnCheckForBoth_TrustDta = True
4160            Exit For
4170          End If
4180        Next
4190        blnCheckForBoth_TrstArch = False
4200        For lngX = 0& To (lngOldFiles - 1&)
4210          If arr_varOldFile(F_FNAM, lngX) = gstrFile_ArchDataName Then
4220            lngTmp14 = lngX
4230            blnCheckForBoth_TrstArch = True
4240            Exit For
4250          End If
4260        Next

4270        If (blnConvert_TrustDta = False And blnCheckForBoth_TrustDta = True) Or _
                ((blnConvert_TrstArch = False And blnArchiveNotPresent = False) And blnCheckForBoth_TrstArch = True) Then

              ' ** We'll have to ask them!
              ' ** 08/03/2011: NO!
              'strTmp04 = vbNullString
              'If blnCheckForBoth_TrustDta = True And blnCheckForBoth_TrstArch = True Then
              '  strTmp04 = "Both the files " & gstrFile_DataName & " And " & gstrFile_ArchDataName & " have been found "
              '  strTmp05 = "backup copies of the same files."
              'ElseIf blnCheckForBoth_TrustDta = True Then
              '  strTmp04 = "The file " & gstrFile_DataName & " has been found "
              '  strTmp05 = "a backup copy of the same file."
              'ElseIf blnCheckForBoth_TrstArch = True Then
              '  strTmp04 = "The file " & gstrFile_ArchDataName & " has been found "
              '  strTmp05 = "a backup copy of the same file."
              'End If
              'strTmp04 = strTmp04 & "in the folder:" & vbCrLf
              'strTmp04 = strTmp04 & "    " & strPath & vbCrLf
              'strTmp04 = strTmp04 & "along with " & strTmp05 & vbCrLf & vbCrLf
              'strTmp04 = strTmp04 & "Since originals and backups should not both be present," & vbCrLf & _
              '  "Trust Accountant needs your help to understand what to do next." & vbCrLf & vbCrLf
              'strTmp04 = strTmp04 & "If you were expecting a Conversion to take place now, or" & vbCrLf
              'strTmp04 = strTmp04 & "    it's your intention to Re-Convert your data," & vbCrLf
              'strTmp04 = strTmp04 & "    click 'YES'." & vbCrLf & vbCrLf
              'strTmp04 = strTmp04 & "If your data is already converted, or" & vbCrLf
              'strTmp04 = strTmp04 & "    you have no idea what this means," & vbCrLf
              'strTmp04 = strTmp04 & "    click 'NO'."
              'msgResponse = MsgBox(strTmp04, vbQuestion + vbYesNo, ("Backups And Originals Both Present" & Space(55)))
4280          msgResponse = vbYes

4290          Version_Up2_Etc strPath, msgResponse, lngOldFiles, arr_varOldFile, blnConvert_TrustDta, blnConvert_TrstArch, _
                blnCheckForBoth_TrustDta, lngDtaElem, lngArchElem, intRetVal1, lngTmp13, lngTmp14, 1  ' ** Module Procedure: modVersionConvertFuncs2.

4300          strTmp04 = vbNullString: strTmp05 = vbNullString: strTmp06 = vbNullString
4310          strTmp09 = vbNullString: strTmp10 = vbNullString
4320          lngTmp13 = 0&: lngTmp14 = 0&

4330        End If  ' ** blnCheckForBoth_TrustDta, blnCheckForBoth_TrstArch.
4340      End If  ' ** blnConvert_TrustDta, blnConvert_TrstArch.

4350      If blnConvert_TrustDta = True Or blnConvert_TrstArch = True Then

4360        strTmp04 = vbNullString
4370        If Version_IsEmpty(strTmp04) = True Then  ' ** Module Function: modVersionConvertFuncs2.

4380          gblnIncomeTaxCoding = False
4390          gblnRevenueExpenseTracking = False
4400          gblnAccountNoWithType = False
4410          gblnSeparateCheckingAccounts = False
4420          gblnTabCopyAccount = False
4430          gblnLinkRevTaxCodes = False
4440          gstrCo_Name = vbNullString
4450          gstrCo_Address1 = vbNullString
4460          gstrCo_Address2 = vbNullString
4470          gstrCo_City = vbNullString
4480          gstrCo_State = vbNullString
4490          gstrCo_Zip = vbNullString
4500          gstrCo_Country = vbNullString
4510          gstrCo_PostalCode = vbNullString
4520          gstrCo_Phone = vbNullString
4530          CoOptions_Read  ' ** Module Function: modUtilities.

              ' ** Create tblVersion_Conversion record.
4540          Set dbsLoc = CurrentDb
              ' ** This opens for the entire rest of the upgrade process.
4550          With dbsLoc
4560            Set rstLoc1 = .OpenRecordset(strVerTbl, dbOpenDynaset, dbConsistent)
4570            With rstLoc1
4580              .AddNew
4590              ![vercnv_date] = Now()
4600              If gstrCo_Name <> vbNullString Then
4610                ![vercnv_name] = gstrCo_Name
4620              Else
4630                ![vercnv_name] = "{to come}"
4640              End If
4650              ![vercnv_verold] = "0.0.0"
4660              ![vercnv_vernew] = AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
4670              ![vercnv_step] = dblPB_ThisStep
4680              ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
4690              ![vercnv_user] = GetUserName  ' ** Module Function: modFileUtilities.
4700              ![vercnv_datemodified] = Now()
4710              .Update
4720              .Bookmark = .LastModified
4730              lngVerCnvID = ![vercnv_id]
4740              .Close
4750            End With  ' ** rstLoc1.
                '.Close' ** This stays open for the entire rest of the upgrade process.
4760          End With  ' ** dbsLoc.
              ' ** dbsLoc opened and closed 8 times in Version_Upgrade_01() error message section.

4770          lngOldTbls = 0&
4780          ReDim arr_varOldTbl(T_ELEMS, 0)

4790          lngOldFlds = 0&
4800          ReDim arr_varOldFld(F_ELEMS, 0)

4810          intRetVal2 = Version_GetOldVer(blnContinue, blnConvert_TrustDta, blnConvert_TrstArch, _
                intWrkType, strOldVersion, lngVerCnvID, lngDtaElem, arr_varOldFile)  ' ** Module Function: modVersionConvertFuncs2.
              ' ** Return values:
              ' **    0  OK
              ' **   -1  Can't Connect
              ' **   -2  Can't Open
              ' **   -9  Error

4820          If blnConvert_TrstArch = True Then
4830            intRetVal3 = Version_ArchCheck(blnContinue, blnConvert_TrstArch, lngArchiveRecs, intWrkType, _
                  lngArchElem, arr_varOldFile)  ' ** Module Function: modVersionConvertFuncs2.
                ' ** Return values:
                ' **    0  OK
                ' **   -7  Can't Open {TrstArch.mdb}
                ' **   -9  Error
4840          End If

4850          If intRetVal2 = 0 Then

                ' ** Open frmVersion_Main status window.
4860            Version_Status 1  ' ** Function: Below.
                'Version_Status IS EXPECTING strOldVersion, SET IN Version_Upgrade_02(), ABOVE!
                ' ** Window opened acDialog, so it won't return till
                ' ** the user either clicks cmdConvert or cmdCancel.
4870            If gblnMessage = False Then
4880              intRetVal1 = -3
4890              blnContinue = False
4900            Else
4910              Version_Status 2  ' ** Function: Below.
                  ' ** Window opened acNormal, so Version_Status() will return
                  ' ** immediately and exit this procedure, returning to Version_Upgrade_01()
                  ' ** with blnContinue = True to conintue with the conversion.
4920            End If

4930          Else
4940            intRetVal1 = intRetVal2
4950          End If

4960        Else
              ' ** It's supposed to do a conversion, but the tables aren't empty!
4970          intRetVal1 = -8
4980        End If

4990      Else
            ' ** Conversion already done.
5000        blnContinue = False
5010        intRetVal1 = 1
5020      End If

5030    Else
          ' ** Not a conversion.
5040      blnContinue = False
5050      intRetVal1 = 1
5060    End If

EXITP:
5070    Version_Upgrade_02 = intRetVal1
5080    Exit Function

ERRH:
5090    intRetVal1 = -9
5100    DoCmd.Hourglass False
5110    lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
5120    Select Case ERR.Number
        Case Else
5130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5140    End Select
5150    Resume EXITP

End Function

Public Function Version_Upgrade_03() As Integer
' ** This function sets up the link to the old data, collects some basic info,
' ** then begins the actual conversion process with the first group of tables.
' ** Tables converted here:
' **   CompanyInformation
' **   adminofficer
' **   Location
' **   RecurringItems
' **   Schedule
' **   ScheduleDetail
' **   DDate
' **   PostingDate
' **   Statement Date
' **   m_REVCODE
' **   Users
' **   _~xusr
' **   tblSecurity_License
' **
' ** At this point, frmVersion_Main should be open and ready to
' ** receive status updates for the duration of the conversion.
' ** Return values:
' **    0  OK
' **   -4  Acount Empty
' **   -5  Canceled CoInfo
' **   -6  Index/Key
' **   -9  Error

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Upgrade_03"

        Dim datCoInfo As Date
        Dim intRetVal As Integer

5210    If gblnDev_NoErrHandle = True Then
5220  On Error GoTo 0
5230    End If

5240    intRetVal = 0

5250    DoCmd.Hourglass True

5260    If blnContinue = True Then  ' ** Is a conversion.

5270      If blnContinue = True And blnConvert_TrustDta = True Or blnConvert_TrstArch = True Then

            ' ** Step 1: Beginning.
5280        dblPB_ThisStep = 1#
5290        Version_Status 3, dblPB_ThisStep, "Start"  ' ** Function: Below.

5300        Set fsfl = Nothing: Set fsfls = Nothing: Set fsfd = Nothing: Set fso = Nothing

5310        Set wrkLoc = DBEngine.Workspaces(0)
5320        Set dbsLoc = wrkLoc.Databases(0)

5330        If blnConvert_TrustDta = True Then

              'WHAT IF ONLY THE ARCHIVE IS BEING CONVERTED?  !!!!

              ' ** Open the workspace with type found in Version_GetOldVer(), modVersionConvertFuncs2.
5340          Select Case intWrkType
              Case 1
5350            Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
5360          Case 2
5370            Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
5380          Case 3
5390            Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
5400          Case 4
5410            Set wrkLnk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
5420          Case 5
5430            Set wrkLnk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
5440          Case 6
5450            Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
5460          Case 7
5470            Set wrkLnk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
5480          End Select

5490          If blnContinue = True Then  ' ** Workspace opens.

5500            With wrkLnk

5510              Set dbsLnk = .OpenDatabase(arr_varOldFile(F_PTHFIL, lngDtaElem), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
                  ' ** Not Exclusive, Read-Only.

5520              If blnContinue = True Then  ' ** Database opens.

5530                With dbsLnk

5540                  blnBadName = False
5550                  For lngX = 0& To (lngBadNames - 1&)
5560                    If arr_varBadName(BN_FILE, lngX) = Parse_File(.Name) Then  ' ** Module Function: modFileUtilities.
5570                      blnBadName = True  ' ** Means there's a bad field name in this database.
5580                      Exit For
5590                    End If
5600                  Next

5610                  For Each tdf In .TableDefs
5620                    With tdf
5630                      If Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "~TMP" And _
                              .Connect = vbNullString Then  ' ** Skip those pesky system tables.

5640                        lngOldTbls = lngOldTbls + 1&
5650                        lngY = lngOldTbls - 1&
5660                        ReDim Preserve arr_varOldTbl(T_ELEMS, lngY)
                            ' **********************************************
                            ' ** Array: arr_varOldTbl()
                            ' **
                            ' **   Element  Name              Constant
                            ' **   =======  ================  ============
                            ' **      0     Name              T_TNAM
                            ' **      1     New Table Name    T_TNAMN
                            ' **      2     Fields            T_FLDS
                            ' **      3     Field Array       T_F_ARR
                            ' **
                            ' **********************************************
5670                        arr_varOldTbl(T_TNAM, lngY) = .Name
5680                        Select Case .Name
                            Case "ReocurringItems"
5690                          arr_varOldTbl(T_TNAMN, lngY) = "RecurringItems"
5700                        Case Else
5710                          arr_varOldTbl(T_TNAMN, lngY) = vbNullString
5720                        End Select
5730                        lngOldFlds = .Fields.Count
5740                        If lngOldFlds = 0& Then
5750                          arr_varOldFile(F_NOTE, lngDtaElem) = arr_varOldFile(F_NOTE, lngDtaElem) & _
                                " TBL: " & .Name & "  FLDS: " & CStr(lngOldFlds)
5760                          arr_varOldFile(F_NOTE, lngDtaElem) = Trim(arr_varOldFile(F_NOTE, lngDtaElem))
5770                        Else
5780                          arr_varOldTbl(T_FLDS, lngY) = lngOldFlds
5790                          arr_varOldTbl(T_F_ARR, lngY) = Empty
5800                          ReDim arr_varOldFld(FD_ELEMS, (lngOldFlds - 1&))
                              ' **********************************
                              ' ** Array: arr_varOldFld()
                              ' **
                              ' **   Element  Name    Constant
                              ' **   =======  ======  ==========
                              ' **      0     Name    FD_FNAM
                              ' **      1     Type    FD_TYP
                              ' **      2     Size    FD_SIZ
                              ' **
                              ' **********************************
5810                        End If

5820                        lngZ = -1&
5830                        For Each fld In .Fields
5840                          With fld
5850                            lngZ = lngZ + 1&
5860                            arr_varOldFld(FD_FNAM, lngZ) = .Name
5870                            arr_varOldFld(FD_TYP, lngZ) = .Type
5880                            arr_varOldFld(FD_SIZ, lngZ) = .Size
5890                          End With  ' ** fld.
5900                        Next

5910                        arr_varOldFile(F_TBLS, lngDtaElem) = lngOldTbls
5920                        arr_varOldTbl(T_F_ARR, lngY) = arr_varOldFld

5930                      End If  ' ** Not a system table.
5940                    End With  ' ** This table: tdf.
5950                  Next  ' ** For each table: tdf.

5960                  arr_varOldFile(F_T_ARR, lngDtaElem) = arr_varOldTbl

5970                End With  ' ** dbsLnk.

                    ' ************************************************************
                    ' ************************************************************
                    ' ** DO THE CONVERSION!
                    ' ** OK, we've got all the info on this TrustDta.mdb,
                    ' ** so now transfer it to our new tables.
                    ' ************************************************************
                    ' ************************************************************

5980                With dbsLnk

5990                  lngRecs = 0&

6000                  lngAccts = 0&
6010                  ReDim arr_varAcct(A_ELEMS, 0)

6020                  lngRevCodes = 0&
6030                  ReDim arr_varRevCode(R_ELEMS, 0)

6040                  lngRevDefCodes = 0&
6050                  ReDim arr_varRevDefCode(RD_ELEMS, 0)

6060                  lngSchedules = 0&
6070                  ReDim arr_varSchedule(S_ELEMS, 0)

6080                  lngInvestObjs = 0&
6090                  ReDim arr_varInvestObj(IO_ELEMS, 0)

6100                  lngAcctTypes = 0&
6110                  lngAssetTypes = 0&
6120                  lngRecurTypes = 0&

6130                  varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & gstrFile_App & "." & _
                        Parse_Ext(CurrentDb.Name) & "'")  ' ** Module Function: modFileUtilities.
6140                  If IsNull(varTmp00) = True Then
                        ' ** Update tblDatabase, for dbs_name = 'Trust.mde'.
6150                    Set qdf = dbsLoc.QueryDefs("qrySecurity_Stat_08")
6160                    qdf.Execute
6170                    Set qdf = Nothing
6180                    DoEvents
6190                    varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & gstrFile_App & "." & _
                          Parse_Ext(CurrentDb.Name) & "'")  ' ** Module Function: modFileUtilities.
6200                  End If
6210                  lngTrustDbsID = Nz(varTmp00, 1&)

6220                  lngTrustDtaDbsID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & gstrFile_DataName & "'")
6230                  lngTrstArchDbsID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & gstrFile_ArchDataName & "'")

                      ' ##################################################
                      ' ## Remember to convert and clean up any residual
                      ' ## references to the old '99-' account numbers.
                      'THESE AREN'T OLD, THEY'RE gblnAccountNoWithType
                      ' ##   99-INCOME O/U -->  INCOME O/U
                      ' ##   99-SUSPENSE   -->  SUSPENSE
                      ' ##################################################
6240                  strAcct99_IncomeOU = vbNullString
6250                  strAcct99_Suspense = vbNullString
6260                  blnAcct99_Both = False

                      ' ** Collect a basic list of accounts, along with some links we'll have to check.
6270                  Set rstLnk = .OpenRecordset("account", dbOpenDynaset, dbReadOnly)
6280                  With rstLnk
6290                    If .BOF = True And .EOF = True Then
                          ' ** Something's horribly wrong!
6300                      intRetVal = -4
6310                      blnContinue = False
6320                    Else
6330                      .MoveLast
6340                      lngRecs = .RecordCount
6350                      .MoveFirst
6360                      For lngX = 1& To lngRecs
                            ' ** Of 16 example TrustDta.mdb's, all have all 5 checked fields.
6370                        lngAccts = lngAccts + 1&
6380                        lngE = lngAccts - 1&
6390                        ReDim Preserve arr_varAcct(A_ELEMS, lngE)
                            ' *******************************************************
                            ' ** Array: arr_varAcct()
                            ' **
                            ' **   Field  Element  Name                 Constant
                            ' **   =====  =======  ===================  ===========
                            ' **     1       0     accountno            A_NUM
                            ' **     2       1     New accountno        A_NUM_N
                            ' **     3       2     shortname            A_NAM
                            ' **     4       3     accounttype          A_TYP
                            ' **     5       4     admin (adminno)      A_ADMIN
                            ' **     6       5     New adminno          A_ADMIN_N
                            ' **     7       6     Schedule_ID          A_SCHED
                            ' **     8       7     New Schedule_ID      A_SCHED_N
                            ' **     9       8     Deleted (T/F)        A_DROPPED
                            ' **    10       9     IncomeIO/Suspense    A_ACCT99
                            ' **    11      10     taxlot (assetno)     A_DASTNO  'Text field.
                            ' **
                            ' *******************************************************
6400                        arr_varAcct(A_NUM, lngE) = ![accountno]
6410                        arr_varAcct(A_ACCT99, lngE) = vbNullString
6420                        If ![accountno] = A99_INC Or ![accountno] = A99_SUS Or _
                                ![accountno] = ("99-" & A99_INC) Or ![accountno] = ("99-" & A99_SUS) Then
6430                          Select Case ![accountno]
                              Case A99_INC, ("99-" & A99_INC)
6440                            If strAcct99_IncomeOU = vbNullString Then
6450                              strAcct99_IncomeOU = ![accountno]
6460                              arr_varAcct(A_ACCT99, lngE) = "10"  ' ** Just a flag field.
6470                            Else
6480                              blnAcct99_Both = True
6490                              arr_varAcct(A_ACCT99, lngE) = "12"
6500                            End If
6510                          Case A99_SUS, ("99-" & A99_SUS)
6520                            If strAcct99_Suspense = vbNullString Then
6530                              strAcct99_Suspense = ![accountno]
6540                              arr_varAcct(A_ACCT99, lngE) = "20"
6550                            Else
6560                              blnAcct99_Both = True
6570                              arr_varAcct(A_ACCT99, lngE) = "22"
6580                            End If
6590                          End Select
6600                        End If
6610                        arr_varAcct(A_NUM_N, lngE) = vbNullString
6620                        arr_varAcct(A_NAM, lngE) = ![shortname]
6630                        arr_varAcct(A_TYP, lngE) = ![accounttype]
6640                        blnFound = False: blnFound2 = False
6650                        For Each fld In .Fields
6660                          With fld
6670                            If .Name = "adminno" Then
                                  ' ** Has newly renamed Administrative Officer field.
6680                              blnFound = True
6690                            ElseIf .Name = "Schedule_ID" Then
                                  ' ** Has newly renamed Schedule_ID field.
6700                              blnFound2 = True
6710                            End If
6720                          End With
6730                        Next
6740                        If blnFound = True Then
6750                          arr_varAcct(A_ADMIN, lngE) = ![adminno]
6760                        Else
6770                          arr_varAcct(A_ADMIN, lngE) = ![Admin]
6780                        End If
6790                        arr_varAcct(A_ADMIN_N, lngE) = CLng(0)
6800                        If blnFound2 = True Then
6810                          arr_varAcct(A_SCHED, lngE) = ![Schedule_ID]
6820                        Else
6830                          arr_varAcct(A_SCHED, lngE) = ![Schedule ID]
6840                        End If
6850                        arr_varAcct(A_SCHED_N, lngE) = CLng(0)
6860                        arr_varAcct(A_DROPPED, lngE) = CBool(False)
6870                        arr_varAcct(A_DASTNO, lngE) = "0"
6880                        If IsNull(![taxlot]) = False Then
6890                          If IsNumeric(![taxlot]) = True Then
6900                            arr_varAcct(A_DASTNO, lngE) = ![taxlot]  ' ** Temporarily using this for the default AssetNo.
6910                          End If
6920                        End If
6930                        If lngX < lngRecs Then .MoveNext
6940                      Next
6950                    End If
6960                    .Close
6970                  End With  ' ** rstLnk.

6980                  lngStats = lngStats + 1&
6990                  lngE = lngStats - 1&
7000                  ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
7010                  arr_varStat(STAT_ORD, lngE) = CInt(2)
7020                  arr_varStat(STAT_NAM, lngE) = "Accounts: "
7030                  arr_varStat(STAT_CNT, lngE) = lngAccts
7040                  arr_varStat(STAT_DSC, lngE) = vbNullString

7050                  If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ** Empty tblVersion_Key.
7060                    Set qdf = dbsLoc.QueryDefs("qryVersion_Convert_01")
7070                    qdf.Execute

                        ' ** Append a dummy record to tblVersion_Key.
                        ' ** Numerous rstLoc2.MoveFirst statements will error if there's nothing there.
7080                    Set qdf = dbsLoc.QueryDefs("qryVersion_Convert_08")
7090                    qdf.Execute

                        ' ******************************
                        ' ** Table: CompanyInformation.
                        ' ******************************

                        ' ** Step 2: CompanyInformation.
7100                    dblPB_ThisStep = 2#
7110                    Version_Status 3, dblPB_ThisStep, "Company Information"  ' ** Function: Below.

7120                    strCurrTblName = "CompanyInformation"
7130                    lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

7140                    blnFound = False: lngRecs = 0&
7150                    For lngX = 0& To (lngOldTbls - 1&)
7160                      If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
7170                        blnFound = True
7180                        Exit For
7190                      End If
7200                    Next

7210                    strTmp04 = vbNullString: strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
7220                    strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
7230                    strTmp11 = vbNullString: strTmp12 = vbNullString
7240                    blnTmp22 = False: blnTmp23 = False: blnTmp24 = False: blnTmp25 = False: blnTmp26 = False: blnTmp27 = False
7250                    varTmp00 = False: lngTmp13 = 0  ' ** So I don't have to create more variables.

7260                    blnFound2 = False

7270                    If blnFound = True Then
7280                      lngRecs = 0&
7290                      Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
7300                      With rstLnk
7310                        If .BOF = True And .EOF = True Then
                              ' ** Empty!
7320                          blnFound = False
7330                        Else
                              ' ** Of 16 example TrustDta.mdb's, some have 7 fields, some have 12.
                              ' ** Current field count is 20 fields.
                              ' ** Table: CompanyInformation
                              ' **   ![CoInfo_ID]                (newest'er field)
                              ' **   ![CoInfo_Name]              (new name)
                              ' **   ![CoInfo_Address1]          (new name)
                              ' **   ![CoInfo_Address2]          (new name)
                              ' **   ![CoInfo_City]              (new name)
                              ' **   ![CoInfo_State]             (new name)
                              ' **   ![CoInfo_Zip]               (new name)
                              ' **   ![CoInfo_Country]           (new field)
                              ' **   ![CoInfo_PostalCode]        (new field)
                              ' **   ![CoInfo_Phone]             (new name)
                              ' **   ![IncomeTaxCoding]          (newer field)
                              ' **   ![RevenueExpenseTracking]   (newer field)
                              ' **   ![AccountNoWithType]        (newer field)
                              ' **   ![SeparateCheckingAccounts] (newer field)
                              ' **   ![TabCopyAccount]           (newer field)
                              ' **   ![LinkRevTaxCodes]          (newest field)
                              ' **   ![SpecialCapGainLoss]       (newest'est field)
                              ' **   ![SpecialCapGainLossOpt]    (newest'est field)
                              ' **   ![Username]                 (newest'er field)
                              ' **   ![CoInfo_DateModified]      (newest'er field)
                              ' ** Referenced by:
                              ' **   This table is not referenced by other tables.
7340                          .MoveLast  ' ** VGC 10/04/2009: CompanyInfo checking added.
                              ' ** Some tables I've seen have a blank first record, and all the info is in a second record!
7350                          lngRecs = .RecordCount
7360                          blnTmp23 = False: blnTmp24 = False: lngTmp15 = 0&: lngTmp16 = 0&
7370                          If lngRecs > 1 Then
7380                            .MoveFirst
7390                            For lngX = 1& To lngRecs
7400                              For Each fld In .Fields
7410                                Select Case fld.Type
                                    Case dbText
7420                                  If IsNull(.Fields(fld.Name)) = False Then
7430                                    blnTmp23 = True
7440                                    Exit For
7450                                  End If
7460                                Case dbBoolean
7470                                  If .Fields(fld.Name) = True Then
                                        ' ** A True would indicate a choice was made, and that this record was used.
                                        ' ** If text data has already been found, subsequent records won't be checked for Boolean choices.
7480                                    blnTmp24 = True
7490                                    lngTmp16 = lngX
7500                                  End If
7510                                Case Else
                                      ' ** Don't care.
7520                                End Select
7530                              Next
7540                              If blnTmp23 = True Then
                                    ' ** This record has info, so use it.
7550                                lngTmp15 = lngX
7560                                Exit For
7570                              Else
                                    ' ** No data in this record, so check the next one.
7580                                If lngX < lngRecs Then .MoveNext
7590                              End If
7600                            Next
7610                            If blnTmp23 = True Then
                                  ' ** Text data found, and because of the Exit For, it should be on that record.
7620                              If blnTmp24 = True And lngTmp16 <> lngTmp15 Then
                                    ' ** A choice was made on a record that isn't the one with the company info.
7630                                If lngRecs = 2& Then
                                      ' ** 6 Boolean fields.  'NOW 7, AND 1 INTEGER!
7640                                  .MoveFirst
7650                                  For Each fld In .Fields
7660                                    Select Case fld.Name
                                        Case "IncomeTaxCoding"
7670                                      blnTmp22 = ![IncomeTaxCoding]
7680                                    Case "RevenueExpenseTracking"
7690                                      blnTmp23 = ![RevenueExpenseTracking]
7700                                    Case "AccountNoWithType"
7710                                      blnTmp24 = ![AccountNoWithType]
7720                                    Case "SeparateCheckingAccounts"
7730                                      blnTmp25 = ![SeparateCheckingAccounts]
7740                                    Case "TabCopyAccount"
7750                                      blnTmp26 = ![TabCopyAccount]
7760                                    Case "LinkRevTaxCodes"
7770                                      blnTmp27 = ![LinkRevTaxCodes]
7780                                    Case "SpecialCapGainLoss"
7790                                      varTmp00 = ![SpecialCapGainLoss]     ' ** Even though it's a Boolean.
7800                                    Case "SpecialCapGainLossOpt"
7810                                      lngTmp13 = ![SpecialCapGainLossOpt]  ' ** Even though it's an Integer.
7820                                    End Select
7830                                  Next
                                      ' ** Go to the record with text data, and move on.
                                      ' ** A non-zero lngTmp15 and lngTmp16 will signal the code below to use these Boolean values.
7840                                  .MoveNext
7850                                Else
                                      ' ** I've never seen one with more than 2, so just forget it, and leave rstLnk where it is.
7860                                  blnTmp23 = False: blnTmp24 = False: lngTmp15 = 0&: lngTmp16 = 0&
7870                                End If
7880                              End If
7890                            Else
                                  ' ** No text data was found.
7900                              If blnTmp24 = True Then
                                    ' ** A boolean choice was made, so at least use that record.
7910                                If lngTmp16 = 1& Then
7920                                  .MoveFirst
7930                                ElseIf lngTmp16 = lngRecs Then
7940                                  .MoveLast
7950                                Else
                                      ' ** I've never seen one with more than 2!
7960                                  .MoveFirst
7970                                  If lngTmp16 >= 2& Then
7980                                    .MoveNext
7990                                  End If
8000                                  If lngTmp16 >= 3& Then
8010                                    .MoveNext
8020                                  End If
                                      ' ** Oh, that's enough!
8030                                End If
8040                              Else
                                    ' ** And no Boolean choices were made, so go to the first and move on.
8050                                .MoveFirst
8060                              End If
8070                              blnTmp23 = False: blnTmp24 = False: lngTmp15 = 0&: lngTmp16 = 0&
8080                            End If
8090                          Else
8100                            .MoveFirst
8110                          End If
8120                          Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, 1&  ' ** Function: Below.
8130                          Version_Status 4, dblPB_ThisStep, Null, 1&, 1&  ' ** Function: Below.
8140                          lngFlds = 0&
8150                          For lngX = 0& To (lngOldFiles - 1&)
8160                            If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
8170                              arr_varTmp02 = arr_varOldFile(F_T_ARR, lngX)
8180                              lngOldTbls = (UBound(arr_varTmp02, 2) + 1)
8190                              For lngY = 0& To (lngOldTbls - 1&)
8200                                If arr_varTmp02(T_TNAM, lngY) = strCurrTblName Then
8210                                  lngFlds = arr_varTmp02(T_FLDS, lngY)
8220                                  Exit For
8230                                End If
8240                              Next
8250                              Exit For
8260                            End If
8270                          Next
8280                          For Each fld In .Fields
8290                            Select Case fld.Name
                                Case "Name"
8300                              strTmp04 = Nz(![Name], vbNullString)
8310                            Case "CoInfo_Name"
8320                              strTmp04 = Nz(![CoInfo_Name], vbNullString)
8330                            Case "Address1"
8340                              strTmp05 = Nz(![Address1], vbNullString)
8350                            Case "CoInfo_Address1"
8360                              strTmp05 = Nz(![CoInfo_Address1], vbNullString)
8370                            Case "Address2"
8380                              strTmp06 = Nz(![Address2], vbNullString)
8390                            Case "CoInfo_Address2"
8400                              strTmp06 = Nz(![CoInfo_Address2], vbNullString)
8410                            Case "City"
8420                              strTmp07 = Nz(![City], vbNullString)
8430                            Case "CoInfo_City"
8440                              strTmp07 = Nz(![CoInfo_City], vbNullString)
8450                            Case "State"
8460                              strTmp08 = Nz(![state], vbNullString)
8470                            Case "CoInfo_State"
8480                              strTmp08 = Nz(![CoInfo_State], vbNullString)
8490                            Case "Zip"
8500                              strTmp09 = Nz(![Zip], vbNullString)
8510                            Case "CoInfo_Zip"
8520                              strTmp09 = Nz(![CoInfo_Zip], vbNullString)
8530                            Case "CoInfo_Country"
8540                              strTmp11 = Nz(![CoInfo_Country], vbNullString)
8550                            Case "CoInfo_PostalCode"
8560                              strTmp12 = Nz(![CoInfo_PostalCode], vbNullString)
8570                            Case "PhoneNumber"
8580                              strTmp10 = Nz(![PhoneNumber], vbNullString)
8590                            Case "CoInfo_Phone"
8600                              strTmp10 = Nz(![CoInfo_Phone], vbNullString)
8610                            Case "CoInfo_ID", "Username", "CoInfo_DateModified"
                                  ' ** These will all be set with the conversion.
8620                            End Select
8630                          Next  ' ** fld.
8640                          If strTmp04 = vbNullString Or strTmp05 = vbNullString Or strTmp07 = vbNullString Or _
                                  (strTmp11 = vbNullString And (strTmp08 = vbNullString Or strTmp09 = vbNullString)) Then
                                ' ** Some info is missing (excluding Address2, PhoneNumber).
8650                            blnFound = False
8660                          Else
8670                            If lngFlds = 7 Then
8680                              blnTmp22 = False: blnTmp23 = False: blnTmp24 = False: blnTmp25 = False: blnTmp26 = False: blnTmp27 = False
8690                              varTmp00 = False: lngTmp13 = 0
8700                            Else
8710                              If lngTmp15 = 0& And lngTmp16 = 0& Then  ' ** Non-zero means Boolean choices were found,
8720                                For Each fld In .Fields                ' ** above, on a record other than this.
8730                                  Select Case fld.Name
                                      Case "IncomeTaxCoding"
8740                                    blnTmp22 = ![IncomeTaxCoding]
8750                                  Case "RevenueExpenseTracking"
8760                                    blnTmp23 = ![RevenueExpenseTracking]
8770                                  Case "AccountNoWithType"
8780                                    blnTmp24 = ![AccountNoWithType]
8790                                  Case "SeparateCheckingAccounts"
8800                                    blnTmp25 = ![SeparateCheckingAccounts]
8810                                  Case "TabCopyAccount"
8820                                    blnTmp26 = ![TabCopyAccount]
8830                                  Case "LinkRevTaxCodes"
8840                                    blnTmp27 = ![LinkRevTaxCodes]
8850                                  Case "SpecialCapGainLoss"
8860                                    varTmp00 = ![SpecialCapGainLoss]
8870                                  Case "SpecialCapGainLossOpt"
8880                                    lngTmp13 = ![SpecialCapGainLossOpt]
8890                                  End Select
8900                                Next
8910                              Else
8920                                lngTmp15 = 0&: lngTmp16 = 0&  ' ** These have done their job.
8930                              End If
8940                            End If
8950                            If blnTmp22 = False And blnTmp23 = False And blnTmp24 = False And blnTmp25 = False And _
                                    blnTmp26 = False And blnTmp27 = False And varTmp00 = False Then
                                  ' ** If they're all False, I'm thinking that means it's never been dealt with.
8960                              blnFound = False
8970                            Else
                                  ' ** Enough pertinent info is present to copy as-is.
8980                            End If
8990                          End If
9000                        End If  ' ** Records present.
9010                        .Close
9020                      End With  ' ** rstLnk.
9030                    End If  ' ** blnFound = True.

9040                    If blnFound = False Then

                          ' ** Delay this at least 5 sec.; it comes too quickly!
9050                      datCoInfo = DateAdd("s", 3, Now())
9060                      Do Until Now() >= datCoInfo
                            ' ** Dum-dee-dum-dum...
9070                      Loop

                          ' ** Company Information needs to be supplied.
9080                      Version_DataXFer_CoInfoVars strTmp04, strTmp05, strTmp06, strTmp07, _
                            strTmp08, strTmp09, strTmp10, strTmp11, strTmp12, lngTmp13, blnTmp22, _
                            blnTmp23, blnTmp24, blnTmp25, blnTmp26, blnTmp27, CBool(varTmp00)  ' ** Module Procedure: modVersionConvertFuncs2.
9090                      blnContinue = Version_DataXFer("Set", "CoInfo")  ' ** Module Function: modVersionConvertFuncs2.
                          'ALL THE ABOVE VARIABLES HAVE TO GET TO modVersionConvertFuncs2!
                          ' ** Do a Set here to pass the data to the form,
                          ' ** which does a Get to populate its controls.
                          ' ** Then it does a Ret to pass the changes back,
                          ' ** where I do a Get to update the table.
9100                      If blnContinue = True Then
9110                        DoCmd.Hourglass False
9120                        DoCmd.OpenForm FRM_CNV_INPUT1, , , , , acDialog, "CoInfo"
9130                        arr_varRetVal = Version_DataXFer("Get", "CoInfo")  ' ** Module Function: modVersionConvertFuncs2.
9140                        DoCmd.Hourglass True
9150                        If Left(arr_varRetVal(0, 0), 1) <> "#" Then  ' ** Covers #ERROR, #EMPTY, and #CANCEL.
9160                          For lngX = 0& To 16&
9170                            If IsEmpty(arr_varRetVal(lngX, 0)) = True Then
9180                              arr_varRetVal(lngX, 0) = vbNullString
9190                            Else
9200                              If IsNull(arr_varRetVal(lngX, 0)) = True Then
9210                                arr_varRetVal(lngX, 0) = vbNullString
9220                              End If
9230                            End If
9240                          Next
9250                          strTmp04 = arr_varRetVal(0, 0)
9260                          strTmp05 = arr_varRetVal(1, 0)
9270                          strTmp06 = arr_varRetVal(2, 0)
9280                          strTmp07 = arr_varRetVal(3, 0)
9290                          strTmp08 = arr_varRetVal(4, 0)
9300                          strTmp09 = arr_varRetVal(5, 0)
9310                          strTmp10 = arr_varRetVal(6, 0)
9320                          If lngTmp15 = 0& And lngTmp16 = 0& Then  ' ** If only partial text data was found, and Boolean
9330                            blnTmp22 = arr_varRetVal(7, 0)         ' ** choices were made on a record other than this, as
9340                            blnTmp23 = arr_varRetVal(8, 0)         ' ** found above, we'll use those choices rather than these.
9350                            blnTmp24 = arr_varRetVal(9, 0)
9360                            blnTmp25 = arr_varRetVal(10, 0)
9370                            blnTmp26 = arr_varRetVal(11, 0)
9380                            blnTmp27 = arr_varRetVal(12, 0)
9390                            varTmp00 = arr_varRetVal(13, 0)
9400                            lngTmp13 = arr_varRetVal(14, 0)
9410                          Else
9420                            lngTmp15 = 0& And lngTmp16 = 0&  ' ** These have done their job.
9430                          End If
9440                          strTmp11 = NullStrIfNull(arr_varRetVal(15, 0))  ' ** Module Function: modStringFuncs.
9450                          strTmp12 = NullStrIfNull(arr_varRetVal(16, 0))  ' ** Module Function: modStringFuncs.
9460                          blnFound = True
9470                        Else
                              ' ** Some sort of problem...
9480                          blnContinue = False
9490                          blnFound = False
9500                          DoCmd.Hourglass False
9510                          If arr_varRetVal(0, 0) = "#CANCEL" Then
9520                            intRetVal = -5
9530                          Else
9540                            intRetVal = -9
9550                          End If
9560                        End If
9570                      Else
                            ' ** Unlikely to be a problem setting the array.
9580                        intRetVal = -9
9590                        blnContinue = False
9600                        blnFound = False
9610                      End If
9620                    End If  ' ** blnFound = False.

9630                    If blnFound = True Then
                          ' ** The data is ready for copying.
9640                      strCurrKeyFldName = vbNullString
9650                      lngCurrKeyFldID = 0&
9660                      Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                          'Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
9670                      With rstLoc1
9680                        If .BOF = True And .EOF = True Then
                              ' ** Shouldn't happen on a new install, but who knows.
9690                          .AddNew
9700                          ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
9710                          ![CoInfo_DateModified] = Now()
9720                          .Update
9730                        End If
9740                        .MoveFirst
9750                        .Edit
9760                        If strTmp04 <> vbNullString Then
9770                          ![CoInfo_Name] = strTmp04
9780                          gstrCo_Name = strTmp04
9790                        End If
9800                        If strTmp05 <> vbNullString Then
9810                          ![CoInfo_Address1] = strTmp05
9820                          gstrCo_Address1 = strTmp05
9830                        End If
9840                        If strTmp06 <> vbNullString Then
9850                          ![CoInfo_Address2] = strTmp06
9860                          gstrCo_Address2 = strTmp06
9870                        End If
9880                        If strTmp07 <> vbNullString Then
9890                          ![CoInfo_City] = strTmp07
9900                          gstrCo_City = strTmp07
9910                        End If
9920                        If strTmp08 <> vbNullString Then
9930                          ![CoInfo_State] = strTmp08
9940                          gstrCo_State = strTmp08
9950                        End If
9960                        If strTmp09 <> vbNullString Then
9970                          ![CoInfo_Zip] = strTmp09
9980                          gstrCo_Zip = strTmp09
9990                        End If
10000                       If strTmp11 <> vbNullString Then
10010                         ![CoInfo_Country] = strTmp11
10020                         gstrCo_Country = strTmp11
10030                       End If
10040                       If strTmp12 <> vbNullString Then
10050                         ![CoInfo_PostalCode] = strTmp12
10060                         gstrCo_PostalCode = strTmp12
10070                       End If
10080                       If strTmp10 <> vbNullString Then
10090                         ![CoInfo_Phone] = strTmp10
10100                         gstrCo_Phone = strTmp10
10110                       End If
                            ' ** Also update the Public variables reflecting these options.
10120                       ![IncomeTaxCoding] = blnTmp22
10130                       gblnIncomeTaxCoding = blnTmp22
10140                       ![RevenueExpenseTracking] = blnTmp23
10150                       gblnRevenueExpenseTracking = blnTmp23
10160                       ![AccountNoWithType] = blnTmp24
10170                       gblnAccountNoWithType = blnTmp24
10180                       ![SeparateCheckingAccounts] = blnTmp25
10190                       gblnSeparateCheckingAccounts = blnTmp25
10200                       ![TabCopyAccount] = blnTmp26
10210                       gblnTabCopyAccount = blnTmp26
10220                       ![LinkRevTaxCodes] = blnTmp27
10230                       gblnLinkRevTaxCodes = blnTmp27
10240                       ![SpecialCapGainLoss] = varTmp00
10250                       gblnSpecialCapGainLoss = varTmp00
10260                       ![SpecialCapGainLossOpt] = lngTmp13
10270                       gintSpecialCapGainLossOpt = lngTmp13
10280                       ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10290                       ![CoInfo_DateModified] = Now()
10300                       .Update
10310                       .Close
10320                     End With  ' ** rstLoc1.

                          ' ** Update tblVersion_Conversion, by specified [vercid], [vernam].
10330                     Set qdf = dbsLoc.QueryDefs("qryVersion_Convert_05")
10340                     With qdf.Parameters
10350                       ![vercid] = lngVerCnvID
10360                       ![vernam] = gstrCo_Name
10370                     End With  ' ** Parameters.
10380                     qdf.Execute

10390                   End If

10400                   lngStats = lngStats + 1&
10410                   lngE = lngStats - 1&
10420                   ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
10430                   arr_varStat(STAT_ORD, lngE) = CInt(1)
10440                   arr_varStat(STAT_NAM, lngE) = "Company: "
10450                   arr_varStat(STAT_CNT, lngE) = CLng(1)
10460                   arr_varStat(STAT_DSC, lngE) = gstrCo_Name

10470                 End If  ' ** blnContinue.

10480                 If blnContinue = True Then

                        ' ******************************
                        ' ** Table: adminofficer.
                        ' ******************************

                        ' ** Step 3: adminofficer.
10490                   dblPB_ThisStep = 3#
10500                   Version_Status 3, dblPB_ThisStep, "Account Admin Officer"  ' ** Function: Below.

10510                   strCurrTblName = "adminofficer"
10520                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

10530                   blnFound = False: lngRecs = 0&
10540                   For lngX = 0& To (lngOldTbls - 1&)
10550                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
10560                       blnFound = True
10570                       Exit For
10580                     End If
10590                   Next

10600                   If blnFound = True Then
10610                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
10620                     With rstLnk
10630                       If .BOF = True And .EOF = True Then
                              ' ** Not used.
10640                       Else
10650                         strCurrKeyFldName = "adminno"
10660                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
10670                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
10680                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, all have 2 fields.
                              ' ** Table: adminofficer
                              ' **   ![adminno]
                              ' **   ![officer]
                              ' ** Referenced by:
                              ' **   Table: account
                              ' **     Field: ![adminno]/![admin]
10690                         .MoveLast
10700                         lngRecs = .RecordCount
10710                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
10720                         .MoveFirst
10730                         For lngX = 1& To lngRecs
10740                           blnTmp22 = True
10750                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
10760                           If IsNull(![officer]) = False Then
10770                             If Trim(![officer]) <> vbNullString Then
10780                               If ![adminno] = 1& And ![officer] = "{Unassigned}" Then
                                      ' ** Cross-check adminno's against referenced admin's in Account table.
10790                                 For lngY = 0& To (lngAccts - 1&)
10800                                   If IsNull(arr_varAcct(A_ADMIN, lngY)) = False Then
10810                                     If arr_varAcct(A_ADMIN, lngY) = ![adminno] Then
10820                                       arr_varAcct(A_ADMIN_N, lngY) = CLng(1)
10830                                     End If
10840                                   End If
10850                                 Next
10860                               Else
                                      ' ** Add the record to the new table.
10870                                 rstLoc1.AddNew
10880                                 rstLoc1![officer] = ![officer]  ' ** dbText (25).
10890 On Error Resume Next
10900                                 rstLoc1.Update
10910                                 If ERR.Number <> 0 Then
10920                                   If ERR.Number = 3022 Then
                                          ' ** Error 3022: The changes you requested to the table were not successful because they
                                          ' **             would create duplicate values in the index, primary key, or relationship.
10930                                     If gblnDev_NoErrHandle = True Then
10940 On Error GoTo 0
10950                                     Else
10960 On Error GoTo ERRH
10970                                     End If
10980                                     lngDupeNum = lngDupeNum + 1&
10990                                     rstLoc1![officer] = Trim(Nz(![officer], vbNullString)) & " #DUPE" & CStr(lngDupeNum)
11000                                     rstLoc1.Update
11010                                     lngDupeUnks = lngDupeUnks + 1&
11020                                     ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
11030                                     arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "DUP"
11040                                     arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
11050                                   Else
11060                                     intRetVal = -6
11070                                     blnContinue = False
11080                                     lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
11090                                     MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
11100                                     rstLoc1.CancelUpdate
11110                                     If gblnDev_NoErrHandle = True Then
11120 On Error GoTo 0
11130                                     Else
11140 On Error GoTo ERRH
11150                                     End If
11160                                   End If
11170                                 Else
11180                                   If gblnDev_NoErrHandle = True Then
11190 On Error GoTo 0
11200                                   Else
11210 On Error GoTo ERRH
11220                                   End If
11230                                 End If
11240                                 If blnContinue = True Then
11250                                   rstLoc1.Bookmark = rstLoc1.LastModified
                                        ' ** Add the ID cross-reference to tblVersion_Key.
11260                                   rstLoc2.AddNew
11270                                   rstLoc2![tbl_id] = lngCurrTblID
11280                                   rstLoc2![tbl_name] = strCurrTblName  'adminofficer
11290                                   rstLoc2![fld_id] = lngCurrKeyFldID
11300                                   rstLoc2![fld_name] = strCurrKeyFldName
11310                                   rstLoc2![key_lng_id1] = ![adminno]
                                        'rstLoc2![key_txt_id1] =
11320                                   rstLoc2![key_lng_id2] = rstLoc1![adminno]
                                        'rstLoc2![key_txt_id2] =
11330                                   rstLoc2.Update
                                        ' ** Cross-check adminno's against referenced admin's in Account table.
11340                                   For lngY = 0& To (lngAccts - 1&)
11350                                     If IsNull(arr_varAcct(A_ADMIN, lngY)) = False Then
11360                                       If arr_varAcct(A_ADMIN, lngY) = ![adminno] Then
11370                                         arr_varAcct(A_ADMIN_N, lngY) = rstLoc1![adminno]
11380                                       End If
11390                                     End If
11400                                   Next
11410                                 Else
11420                                   blnTmp22 = False
11430                                 End If
11440                               End If
11450                             Else
                                    ' ** A Null String record.
11460                               For lngY = 0& To (lngAccts - 1&)
11470                                 If IsNull(arr_varAcct(A_ADMIN, lngY)) = False Then
11480                                   If arr_varAcct(A_ADMIN, lngY) = ![adminno] Then
11490                                     arr_varAcct(A_ADMIN_N, lngY) = CLng(0)
11500                                   End If
11510                                 End If
11520                               Next
11530                             End If
11540                           Else
                                  ' ** A Null record.
11550                             For lngY = 0& To (lngAccts - 1&)
11560                               If IsNull(arr_varAcct(A_ADMIN, lngY)) = False Then
11570                                 If arr_varAcct(A_ADMIN, lngY) = ![adminno] Then
11580                                   arr_varAcct(A_ADMIN_N, lngY) = CLng(0)
11590                                 End If
11600                               End If
11610                             Next
11620                           End If
11630                           If blnTmp22 = True Then
11640                             If lngX < lngRecs Then .MoveNext
11650                           Else
11660                             Exit For
11670                           End If
11680                         Next
11690                         If blnContinue = True Then
                                ' ** Check if all referenced admin's in Account table are present.
11700                           For lngX = 0& To (lngAccts - 1&)
11710                             If IsNull(arr_varAcct(A_ADMIN, lngX)) = False Then
11720                               If arr_varAcct(A_ADMIN_N, lngX) = 0& And arr_varAcct(A_ADMIN, lngX) <> 0& Then
                                      ' ** Admin in account table not found in adminofficer table.
11730                                 With rstLoc1
                                        ' ** Add the record to the new table.
11740                                   .AddNew
11750                                   strTmp04 = "UNKNOWN" & CStr(arr_varAcct(A_ADMIN, lngX))
11760                                   strTmp04 = Left(strTmp04 & " Acct# " & arr_varAcct(A_NUM, lngX), 25)  ' ** [accountno] is max 15 chars.
11770                                   ![officer] = strTmp04  ' ** dbText (25).
11780                                   .Update
11790                                   lngDupeUnks = lngDupeUnks + 1&
11800                                   ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
11810                                   arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
11820                                   arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
11830                                   .Bookmark = .LastModified
11840                                   arr_varAcct(A_ADMIN_N, lngX) = ![adminno]
                                        ' ** Update the rest of arr_varAcct().
11850                                   For lngY = 0& To (lngAccts - 1&)
11860                                     If lngY <> lngX Then  ' ** Skip the one we're on.
11870                                       If IsNull(arr_varAcct(A_ADMIN, lngY)) = False Then
11880                                         If arr_varAcct(A_ADMIN, lngY) = arr_varAcct(A_ADMIN, lngX) Then
11890                                           arr_varAcct(A_ADMIN_N, lngY) = arr_varAcct(A_ADMIN_N, lngX)
11900                                         End If
11910                                       End If
11920                                     End If
11930                                   Next
11940                                 End With  ' ** rstLoc1.
11950                                 With rstLoc2
                                        ' ** Add the ID cross-reference to tblVersion_Key.
11960                                   .AddNew
11970                                   ![tbl_id] = lngCurrTblID
11980                                   ![tbl_name] = strCurrTblName  'adminofficer
11990                                   ![fld_id] = lngCurrKeyFldID
12000                                   ![fld_name] = strCurrKeyFldName
12010                                   ![key_lng_id1] = arr_varAcct(A_ADMIN, lngX)
                                        '![key_txt_id1] =
12020                                   ![key_lng_id2] = arr_varAcct(A_ADMIN_N, lngX)
                                        '![key_txt_id2] =
12030                                   .Update
12040                                 End With  ' ** rstLoc2.
12050                               End If
12060                             End If
12070                           Next
12080                         End If  ' ** blnContinue.
12090                         rstLoc1.Close
12100                         rstLoc2.Close
12110                       End If  ' ** Records present.
12120                       .Close
12130                     End With  ' ** rstLnk.
                          ' ** When the referenced table(s) are converted,
                          ' ** do so using this tblVersion_Key.
                          ' ** If an orphan is found, add the record at that time.
                          ' ** Well... I guess I'm not doing that for this table and the Account table...

12140                   End If  ' ** blnFound.

12150                 End If  ' ** blnContinue.

12160                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: Location.
                        ' ******************************

                        ' ** Step 4: Location.
12170                   dblPB_ThisStep = 4#
12180                   Version_Status 3, dblPB_ThisStep, "Location"  ' ** Function: Below.

12190                   strCurrTblName = "Location"
12200                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

12210                   blnFound = False: blnFound2 = False: lngRecs = 0&
12220                   For lngX = 0& To (lngOldTbls - 1&)
12230                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
12240                       blnFound = True
12250                       Exit For
12260                     End If
12270                   Next

12280                   If blnFound = True Then
12290                     lngRecs = 0&
12300                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
12310                     With rstLnk
12320                       If .BOF = True And .EOF = True Then
                              ' ** Not used.
                              ' ** No records are needed. Make sure referenced [Location_ID] Nulls and 0's are all changed to 1's.
                              ' **   [Location_ID] = 1 : '{Unassigned}'
12330                         blnFound = False
12340                       Else
12350                         strCurrKeyFldName = "Location_ID"
12360                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
12370                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
12380                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, all have 8 fields.
                              ' ** Current field count is 13 fields.
                              ' ** Table: Location  {all new field names}
                              ' **   ![Location_ID]
                              ' **   ![Location_Name]
                              ' **   ![Location_Address1]
                              ' **   ![Location_Address2]
                              ' **   ![Location_City]
                              ' **   ![Location_State]
                              ' **   ![Location_Zip]
                              ' **   ![Location_Phone]
                              ' **   ![Username]                 {newest field}
                              ' **   ![Location_DateCreated]     {newest field}
                              ' **   ![Location_DateModified]    {newest field}
                              ' **   Table: ActiveAssets                  {Linked}
                              ' **     Field: [Location_ID]
                              ' **   Table: journal                       {Linked}
                              ' **     Field: [Location_ID]
                              ' **   Table: Journal Map                   {Linked}
                              ' **     Field: [Location_ID]
                              ' **   Table: ledger                        {Linked}
                              ' **     Field: [Location_ID]
                              ' **   Table: LedgerArchive                 {Linked}
                              ' **     Field: [Location_ID]
                              ' **   Table: tblJournal_Column             {Local}
                              ' **     Field: [Location_ID]
                              ' **            [Loc_Name]
                              ' **            [Loc_Name_display]
                              ' **   Table: tblTemplate_ActiveAssets      {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tmpEdit01                     {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tmpEdit04                     {Local}
                              ' **     Field: [Location_ID]
                              ' **            [Location_Namex]
                              ' **   Table: tblTemplate_Journal           {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tblTemplate_Ledger            {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tblTemplate_Location          {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tmpLocation                   {Local}
                              ' **     Field: [Location_ID]
                              ' **            [Location_ID_new]
                              ' **   Table: tmpLocation2                  {Local}
                              ' **     Field: [Location_ID]
                              ' **            [Location_ID_new]
                              ' **   Table: tmpXAdmin_ActiveAssets_02     {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tmpXAdmin_Ledger_01           {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tmpXAdmin_Ledger_02           {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tmpXAdmin_Ledger_03           {Local}
                              ' **     Field: [Location_ID]
                              ' **   Table: tmpXAdmin_Ledger_04           {Local}
                              ' **     Field: [Location_ID]
12390                         .MoveLast
12400                         lngRecs = .RecordCount
12410                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
12420                         .MoveFirst
12430                         lngFlds = 0&
12440                         For lngX = 0& To (lngOldFiles - 1&)
12450                           If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
12460                             arr_varTmp02 = arr_varOldFile(F_T_ARR, lngX)
12470                             lngOldTbls = (UBound(arr_varTmp02, 2) + 1)
12480                             For lngY = 0& To (lngOldTbls - 1&)
12490                               If arr_varTmp02(T_TNAM, lngY) = strCurrTblName Then
12500                                 lngFlds = arr_varTmp02(T_FLDS, lngY)
12510                                 Exit For
12520                               End If
12530                             Next
12540                             Exit For
12550                           End If
12560                         Next
12570                         For Each fld In .Fields
12580                           With fld
12590                             If .Name = "Name" Then
                                    ' ** Has old field names.
12600                               blnFound2 = True
12610                               Exit For
12620                             End If
12630                           End With
12640                         Next
12650                         strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
12660                         strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
12670                         strTmp06 = "Location_ID": strTmp07 = "Location ID"
12680                         If blnFound2 = False Then strTmp05 = strTmp06 Else strTmp05 = strTmp07
12690                         For lngX = 1& To lngRecs
12700                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
                                ' ** Check for '{Unassigned}', so we don't get 2 of them!
12710                           strTmp09 = "Location_Name": strTmp10 = "Name"
12720                           If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
12730                           If .Fields(strTmp08) = "{Unassigned}" Then
                                  ' ** If its [Location_ID] = 1, then just skip copying this one.
                                  ' ** It really only should be 1, since no one
                                  ' ** else would have ever added it to the table!
12740                           Else
                                  ' ** Add the record to the new table.
12750                             rstLoc1.AddNew
12760                             For Each fld In .Fields
12770                               varTmp00 = Empty
12780                               If fld.Name <> strTmp05 Then
12790                                 varTmp00 = fld.Value
12800                                 strTmp09 = "Location_Name": strTmp10 = "Name"
12810                                 If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
12820                                 If fld.Name = strTmp08 Then
12830                                   If IsNull(varTmp00) = True Then
                                          ' ** This section handles Null Location_Name's.
12840                                     Set rstLoc3 = dbsLoc.OpenRecordset("CompanyInformation", dbOpenDynaset, dbReadOnly)
12850                                     With rstLoc3
12860                                       If .BOF = True And .EOF = True Then
                                              ' ** Should have already been converted or filled!
12870                                       Else
12880                                         .MoveFirst
12890                                         If .Fields(strTmp08) = "Sidney C. Summey, Conservator" And rstLnk.Fields(strTmp05) = 27 Then
                                                ' ** Client-Specific Fix: White, Arnold, & Dowd.
12900                                           varTmp00 = "Alabama Department of Public Health"
12910                                         Else
                                                ' ** Generic Fix.
12920                                           varTmp00 = CStr("UNKNOWN " & CStr(rstLnk.Fields(strTmp05).Value))
12930                                         End If
12940                                       End If
12950                                       .Close
12960                                     End With
12970                                     Set rstLoc3 = Nothing
12980                                   End If
12990                                 End If
13000                                 strTmp08 = vbNullString
13010                                 Select Case blnFound2
                                      Case True
13020                                   Select Case fld.Name
                                        Case "Name"
13030                                     strTmp08 = "Location_Name"
13040                                   Case "Address1"
13050                                     strTmp08 = "Location_Address1"
13060                                   Case "Address2"
13070                                     strTmp08 = "Location_Address2"
13080                                   Case "City"
13090                                     strTmp08 = "Location_City"
13100                                   Case "State"
13110                                     strTmp08 = "Location_State"
13120                                   Case "Zip"
13130                                     strTmp08 = "Location_Zip"
13140                                   Case "Phone"
13150                                     strTmp08 = "Location_Phone"
13160                                   Case "Username"
13170                                     strTmp08 = "Username"  ' ** No name change.
13180                                   Case "DateCreated"
13190                                     strTmp08 = "Location_DateCreated"
13200                                   Case "DateModified"
13210                                     strTmp08 = "Location_DateModified"
13220                                   End Select
13230                                   If strTmp08 <> vbNullString Then
13240                                     rstLoc1.Fields(strTmp08) = varTmp00
13250                                   End If
13260                                 Case False
13270                                   rstLoc1.Fields(fld.Name) = varTmp00
13280                                 End Select
13290                               End If
13300                             Next
13310                             varTmp00 = Empty
13320                             If lngFlds = 8& Then
13330                               rstLoc1![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
13340                               rstLoc1![Location_DateCreated] = Now()
13350                               rstLoc1![Location_DateModified] = Now()
13360                             Else
                                    ' ** A very new version indeed!
                                    ' ** The values would've come over with the For/Next loop above.
13370                             End If
                                  ' ** What's the best way to handle dupes that'll hiccup
                                  ' ** in the new table, but weren't a problem in the old?
                                  ' ** Their ID may still be used, so I can't just throw them out.
                                  ' ** 1. Add an 'X' to the name, then re-update.
                                  ' ** 2. Save the record off to the side for handling later.
                                  ' ** 3. Check for dupes before beginning that table's transfer;
                                  ' **    though this would still require checking for its links later.
                                  ' ** ...
                                  ' ** OK, if I save it off to the side (whether or not I checked ahead of time),
                                  ' ** then, as each referenced table is processed, I'd check this side table
                                  ' ** for dupes, and switch their ID's then.
                                  ' ** Adding an 'X' is certainly the simplest.
                                  ' ** I could check if it's used later, then delete it or do the
                                  ' ** update of referenced tables then. I could then use canned queries.
13380 On Error Resume Next
13390                             rstLoc1.Update
13400                             If ERR.Number <> 0 Then
13410                               If ERR.Number = 3022 Then
                                      ' ** Error 3022: The changes you requested to the table were not successful because they
                                      ' **             would create duplicate values in the index, primary key, or relationship.
13420                                 If gblnDev_NoErrHandle = True Then
13430 On Error GoTo 0
13440                                 Else
13450 On Error GoTo ERRH
13460                                 End If
                                      'ANY WAY TO CHECK IF ONE HAS AN ADDRESS AND THE OTHER DOESN'T?
13470                                 lngDupeNum = lngDupeNum + 1&
13480                                 strTmp09 = "Location_Name": strTmp10 = "Name"
13490                                 If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
13500                                 strTmp04 = Trim(Nz(.Fields(strTmp08), vbNullString)) & " #DUPE" & CStr(lngDupeNum)
13510                                 If Len(strTmp04) > 35 Then
13520                                   strTmp04 = Left(Trim(Nz(.Fields(strTmp08), vbNullString)), (35 - Len(" #DUPE" & CStr(lngDupeNum)))) & _
                                          " #DUPE" & CStr(lngDupeNum)
13530                                 End If
13540                                 rstLoc1![Location_Name] = strTmp04  ' ** 35 chars max.
13550                                 rstLoc1.Update
13560                                 lngDupeUnks = lngDupeUnks + 1&
13570                                 ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
13580                                 arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "DUP"
13590                                 arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
13600                                 strTmp04 = vbNullString
13610                               Else
13620                                 intRetVal = -6
13630                                 blnContinue = False
13640                                 lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
13650                                 MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                        "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                        vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
13660                                 rstLoc1.CancelUpdate
13670                                 If gblnDev_NoErrHandle = True Then
13680 On Error GoTo 0
13690                                 Else
13700 On Error GoTo ERRH
13710                                 End If
13720                               End If
13730                             Else
13740                               If gblnDev_NoErrHandle = True Then
13750 On Error GoTo 0
13760                               Else
13770 On Error GoTo ERRH
13780                               End If
13790                             End If
13800                             If blnContinue = True Then
13810                               rstLoc1.Bookmark = rstLoc1.LastModified
                                    ' ** Add the ID cross-reference to tblVersion_Key.
13820                               rstLoc2.AddNew
13830                               rstLoc2![tbl_id] = lngCurrTblID
13840                               rstLoc2![tbl_name] = strCurrTblName  'Location
13850                               rstLoc2![fld_id] = lngCurrKeyFldID
13860                               rstLoc2![fld_name] = strCurrKeyFldName
13870                               rstLoc2![key_lng_id1] = .Fields(strTmp05)
                                    'rstLoc2![key_txt_id1] =
13880                               rstLoc2![key_lng_id2] = rstLoc1![Location_ID]
                                    'rstLoc2![key_txt_id2] =
13890                               rstLoc2.Update
13900                             Else
13910                               Exit For
13920                             End If
13930                           End If
13940                           If lngX < lngRecs Then .MoveNext
13950                         Next
13960                         rstLoc1.Close
13970                         rstLoc2.Close
13980                       End If  ' ** Records present.
13990                       .Close
14000                     End With  ' ** rstLnk.
14010                   End If  ' ** blnFound.
14020                   strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
14030                   strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString

14040                 End If  ' ** blnContinue.

14050                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ** Get a list of Recurring Types.
14060                   Set qdf = dbsLoc.QueryDefs("qryRecurringType_02")
14070                   Set rstLoc1 = qdf.OpenRecordset
14080                   With rstLoc1
14090                     .MoveLast
14100                     lngRecurTypes = .RecordCount
14110                     .MoveFirst
14120                     arr_varRecurType = .GetRows(lngRecurTypes)
                          ' *****************************************************
                          ' ** Array: arr_varRecurType()
                          ' **
                          ' **   Field  Element  Name                Constant
                          ' **   =====  =======  ==================  ==========
                          ' **     1       0     RecurringType_ID    RT_CODE
                          ' **     2       1     RecurringType       RT_DESC
                          ' **     3       2     journaltype         RT_JTYP
                          ' **
                          ' *****************************************************
14130                     .Close
14140                   End With  ' ** rstLoc1.

                        ' ******************************
                        ' ** Table: RecurringItems.
                        ' ******************************

                        ' ** Step 5: RecurringItems.
14150                   dblPB_ThisStep = 5#
14160                   Version_Status 3, dblPB_ThisStep, "Recurring Items"  ' ** Function: Below.

14170                   strCurrTblName = "RecurringItems"
14180                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

14190                   blnFound = False: blnFound2 = False: lngRecs = 0&: strTmp05 = vbNullString
14200                   strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
14210                   For lngX = 0& To (lngOldTbls - 1&)
14220                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
14230                       blnFound = True
14240                       Exit For
14250                     End If
14260                   Next

                        ' ** Table Name Change!
                        ' **   ReocurringItems -> RecurringItems
14270                   If blnFound = False Then
14280                     For lngX = 0& To (lngOldTbls - 1&)
14290                       If arr_varOldTbl(T_TNAMN, lngX) = strCurrTblName Then
                              ' ** Has old table and field names.
14300                         blnFound2 = True
14310                         blnFound = True
14320                         strTmp05 = arr_varOldTbl(T_TNAM, lngX)
14330                         Exit For
14340                       End If
14350                     Next
14360                   End If

14370                   If blnFound = True Then
14380                     lngRecs = 0&
14390                     Select Case blnFound2
                          Case True
14400                       Set rstLnk = .OpenRecordset(strTmp05, dbOpenDynaset, dbReadOnly)
14410                     Case False
14420                       Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
14430                     End Select

14440                     With rstLnk
14450                       If .BOF = True And .EOF = True Then
                              ' ** Not used.
                              ' ** No records are needed. Make sure referenced Name's are properly accounted for.
                              ' **   RecurringItem_ID = 1 : 'Transferred Income Cash to Principal Cash'
                              ' **   RecurringItem_ID = 2 : 'Transferred Principal Cash to Income Cash'
14460                         blnFound = False
14470                       Else
14480                         strCurrKeyFldName = "RecurringItem_ID"
14490                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
14500                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
14510                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, some have 2 fields, some have 6.
                              ' ** Current field count is 10 fields.
                              ' ** Table: RecurringItems  {all new field names}
                              ' **   ![RecurringItem_ID]
                              ' **   ![RecurringItem]                     dbText 50  {oldest field}
                              ' **   ![RecurringType]                                {oldest field}
                              ' **   ![RecurringItem_Address]
                              ' **   ![RecurringItem_City]
                              ' **   ![RecurringItem_State]
                              ' **   ![RecurringItem_Zip]
                              ' ** Referenced by:
                              ' **   Table: journal                       {Linked}
                              ' **     Field: [RecurringItem]             dbText 50
                              ' **   Table: ledger                        {Linked}
                              ' **     Field: [RecurringItem]             dbText 50
                              ' **   Table: LedgerArchive                 {Linked}
                              ' **     Field: [RecurringItem]             dbText 50
                              ' **   Table: tmpCourtReportData2           {Local}
                              ' **     Field: [RecurringItem]             dbText 255
                              ' **   Table: tmpTaxDisbursementsDeductions {Local}
                              ' **     Field: [RecurringItem]             dbText 255
                              ' **   Table: tmpTaxReceiptsIncome          {Local}
                              ' **     Field: [RecurringItem]             dbText 255
                              ' **   Table: tmpTaxReports                 {Local}
                              ' **     Field: [RecurringItem]             dbText 50
14520                         .MoveLast
14530                         lngRecs = .RecordCount
14540                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
14550                         .MoveFirst
14560                         lngFlds = 0&
14570                         For lngX = 0& To (lngOldFiles - 1&)
14580                           If arr_varOldFile(F_FNAM, lngX) = gstrFile_DataName Then
14590                             arr_varTmp02 = arr_varOldFile(F_T_ARR, lngX)
14600                             lngOldTbls = (UBound(arr_varTmp02, 2) + 1)
14610                             For lngY = 0& To (lngOldTbls - 1&)
14620                               If arr_varTmp02(T_TNAM, lngY) = strCurrTblName Or arr_varTmp02(T_TNAMN, lngY) = strCurrTblName Then
14630                                 lngFlds = arr_varTmp02(T_FLDS, lngY)
14640                                 Exit For
14650                               End If
14660                             Next
14670                             Exit For
14680                           End If
14690                         Next
                              ' ** Some have 2 fields, some have 6; currently there are 7.
14700                         lngOld_RECUR_I_TO_P_ID = 0&: lngOld_RECUR_P_TO_I_ID = 0&
14710                         For lngX = 1& To lngRecs
                                ' ** 09/17/2009: Remember: I've changed the the index
                                ' ** to allow for the same Name, but with different Type.
14720                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
                                ' ** Check for our new 'Transferred ...' records, so we don't get 2 sets of them!
                                ' **   RECUR_I_TO_P As String = "Transferred Income Cash to Principal Cash"
                                ' **   RECUR_I_TO_P_ID As Long = 1&
                                ' **   RECUR_P_TO_I As String = "Transferred Principal Cash to Income Cash"
                                ' **   RECUR_P_TO_I_ID As Long = 2&
14730                           strTmp09 = "RecurringItem": strTmp10 = "Name"
14740                           If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
14750                           If .Fields(strTmp08) = RECUR_I_TO_P Or .Fields(strTmp08) = RECUR_P_TO_I Then
14760                             If lngFlds = 7& Then
14770                               If .Fields(strTmp08) = RECUR_I_TO_P And ![RecurringItem_ID] <> RECUR_I_TO_P_ID Then
                                      ' ** It really only should be 1, since no one
                                      ' ** else would have ever added it to the table!
14780                                 lngOld_RECUR_I_TO_P_ID = ![RecurringItem_ID]
14790                               ElseIf .Fields(strTmp08) = RECUR_P_TO_I And ![RecurringItem_ID] <> RECUR_P_TO_I_ID Then
                                      ' ** It really only should be 2, since no one
                                      ' ** else would have ever added it to the table!
14800                                 lngOld_RECUR_P_TO_I_ID = ![RecurringItem_ID]
14810                               Else
                                      ' ** ID's the same, so just drop these.
14820                               End If
14830                             Else
                                    ' ** Since it didn't have a RecurringItem_ID to begin with, no need to watch out for it.
14840                             End If
14850                           ElseIf .Fields(strTmp08) = "Transfered Income Cash to Principal Cash" Or _
                                    .Fields(strTmp08) = "Transferred Income to Principal" Or _
                                    .Fields(strTmp08) = "Transfered Income to Principal" Or _
                                    .Fields(strTmp08) = "Moved Income Cash to Principal Cash" Or _
                                    .Fields(strTmp08) = "Moved Income to Principal" Or _
                                    .Fields(strTmp08) = "Transfer Income Cash to Principal Cash" Or _
                                    .Fields(strTmp08) = "Transfer Income to Principal" Then
                                  ' ** Yes, I'm including the misspellings.
14860                             If lngFlds = 7& Then
                                    ' ** This would be truly bizarre!
14870                               lngOld_RECUR_I_TO_P_ID = ![RecurringItem_ID]
14880                             Else
                                    ' ** Just drop it, though we could change its text in the referenced tables.
                                    ' ** Naw... That's too obsessive.
14890                             End If
14900                           ElseIf .Fields(strTmp08) = "Transfered Principal Cash to Income Cash" Or _
                                    .Fields(strTmp08) = "Transferred Principal to Income" Or _
                                    .Fields(strTmp08) = "Transfered Principal to Income" Or _
                                    .Fields(strTmp08) = "Moved Principal Cash to Income Cash" Or _
                                    .Fields(strTmp08) = "Moved Principal to Income" Or _
                                    .Fields(strTmp08) = "Transfer Principal Cash to Income Cash" Or _
                                    .Fields(strTmp08) = "Transfer Principal to Income" Then
14910                             If lngFlds = 7& Then
                                    ' ** And so would this!
14920                               lngOld_RECUR_P_TO_I_ID = ![RecurringItem_ID]
14930                             Else
                                    ' ** Drop it.
14940                             End If
14950                           Else
                                  ' ** Add the record to the new table.
14960                             rstLoc1.AddNew
14970                             rstLoc1![RecurringItem] = .Fields(strTmp08)
14980                             blnFound = False: strTmp04 = vbNullString
14990                             strTmp09 = "RecurringType": strTmp10 = "Type"
15000                             If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
15010                             If IsNull(.Fields(strTmp08)) = False Then
                                    ' ** One of our customer's had a type of '################'!
15020                               For lngY = 0& To (lngRecurTypes - 1&)
15030                                 If arr_varRecurType(RT_DESC, lngY) = .Fields(strTmp08) Then
15040                                   blnFound = True
15050                                   strTmp04 = .Fields(strTmp08)
15060                                   Exit For
15070                                 End If
15080                               Next
15090                             End If
15100                             If blnFound = False Then
15110                               strTmp04 = "Misc"
15120                             End If
15130                             rstLoc1![RecurringType] = strTmp04
15140                             blnFound = True: strTmp04 = vbNullString
15150                             If lngFlds >= 6 Then
15160                               Select Case blnFound2
                                    Case True
15170                                 rstLoc1![RecurringItem_Address] = ![Address]
15180                                 rstLoc1![RecurringItem_City] = ![City]
15190                                 rstLoc1![RecurringItem_State] = ![state]
15200                                 rstLoc1![RecurringItem_Zip] = ![Zip]
15210                                 rstLoc1![RecurringItem_DateModified] = Now()
15220                               Case False
15230                                 rstLoc1![RecurringItem_Address] = ![RecurringItem_Address]
15240                                 rstLoc1![RecurringItem_City] = ![RecurringItem_City]
15250                                 rstLoc1![RecurringItem_State] = ![RecurringItem_State]
15260                                 rstLoc1![RecurringItem_Zip] = ![RecurringItem_Zip]
15270                                 rstLoc1![RecurringItem_DateModified] = ![RecurringItem_DateModified]
15280                               End Select
15290                             Else
15300                               rstLoc1![RecurringItem_DateModified] = Now()
15310                             End If
15320 On Error Resume Next
15330                             rstLoc1.Update
15340                             If ERR.Number <> 0 Then
                                    ' ** 09/17/2009: Remember: I've changed the the index
                                    ' ** to allow for the same Name, but with different Type,
                                    ' ** so this should no longer produce a '#DUPE'!
15350                               If ERR.Number = 3022 Then
                                      ' ** Error 3022: The changes you requested to the table were not successful because they
                                      ' **             would create duplicate values in the index, primary key, or relationship.
15360                                 If gblnDev_NoErrHandle = True Then
15370 On Error GoTo 0
15380                                 Else
15390 On Error GoTo ERRH
15400                                 End If
15410                                 lngDupeNum = lngDupeNum + 1&
15420                                 strTmp09 = "RecurringItem": strTmp10 = "Name"
15430                                 If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
15440                                 rstLoc1![RecurringItem] = Trim(Nz(.Fields(strTmp08), vbNullString)) & " #DUPE" & CStr(lngDupeNum)
15450                                 rstLoc1.Update
15460                                 lngDupeUnks = lngDupeUnks + 1&
15470                                 ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
15480                                 arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "DUP"
15490                                 arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
15500                               Else
15510                                 intRetVal = -6
15520                                 blnContinue = False
15530                                 lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
15540                                 MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                        "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                        vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
15550                                 rstLoc1.CancelUpdate
15560                                 If gblnDev_NoErrHandle = True Then
15570 On Error GoTo 0
15580                                 Else
15590 On Error GoTo ERRH
15600                                 End If
15610                               End If
15620                             Else
15630                               If gblnDev_NoErrHandle = True Then
15640 On Error GoTo 0
15650                               Else
15660 On Error GoTo ERRH
15670                               End If
15680                             End If
15690                             If blnContinue = True Then
15700                               rstLoc1.Bookmark = rstLoc1.LastModified
                                    ' ** Add the ID cross-reference to tblVersion_Key.
15710                               rstLoc2.AddNew
15720                               rstLoc2![tbl_id] = lngCurrTblID
15730                               rstLoc2![tbl_name] = strCurrTblName  'RecurringItems
15740                               rstLoc2![fld_id] = lngCurrKeyFldID
15750                               rstLoc2![fld_name] = strCurrKeyFldName
15760                               If lngFlds = 7& Then
15770                                 rstLoc2![key_lng_id1] = ![RecurringItem_ID]
                                      'rstLoc2![key_txt_id1] =
15780                               Else
15790                                 rstLoc2![key_lng_id1] = 0&
                                      'rstLoc2![key_txt_id1] =
15800                               End If
15810                               rstLoc2![key_lng_id2] = rstLoc1![RecurringItem_ID]
                                    'rstLoc2![key_txt_id2] =
15820                               rstLoc2.Update
15830                             Else
15840                               Exit For
15850                             End If
15860                           End If
15870                           If lngX < lngRecs Then .MoveNext
15880                         Next
15890                         rstLoc1.Close
15900                         rstLoc2.Close
15910                       End If  ' ** Records present.
15920                       .Close
15930                     End With  ' ** rstLnk.
15940                   End If  ' ** blnFound.
15950                   blnFound2 = False: strTmp05 = vbNullString
15960                   strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString

15970                 End If  ' ** blnContinue.

15980                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: Schedule.
                        ' **        ScheduleDetail.
                        ' ******************************

                        ' ** Step 6: Schedule, ScheduleDetail.
15990                   dblPB_ThisStep = 6#
16000                   Version_Status 3, dblPB_ThisStep, "Schedule"  ' ** Function: Below.

16010                   strCurrTblName = "Schedule"
16020                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

16030                   blnFound = False: blnFound2 = False: lngRecs = 0&: strTmp04 = vbNullString
16040                   strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
16050                   strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
16060                   For lngX = 0& To (lngOldTbls - 1&)
16070                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
16080                       blnFound = True
16090                       Exit For
16100                     End If
16110                   Next

16120                   If blnFound = True Then
16130                     lngRecs = 0&
16140                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
16150                     With rstLnk
16160                       If .BOF = True And .EOF = True Then
                              ' ** Not used.
                              ' ** No records are needed. Make sure referenced Schedule_ID's are properly accounted for.
16170                         blnFound = False
16180                       Else
16190                         strCurrKeyFldName = "Schedule_ID"
16200                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
16210                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
16220                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Of 16 example TrustDta.mdb's, all have 4 fields.
                              ' ** Current field count is 6 fields.
                              ' ** Table: Schedule  {all new field names}
                              ' **   ![Schedule_ID]
                              ' **   ![Schedule_Name]
                              ' **   ![Schedule_Base]
                              ' **   ![Schedule_Minimum]
                              ' **   ![Schedule_DateCreated]     {newest field}
                              ' **   ![Schedule_DateModified]    {newest field}
                              ' ** Referenced by:
                              ' **   Table: account                       {Linked}
                              ' **     Field: [Schedule_ID]
                              ' **   Table: ScheduleDetail                {Linked}
                              ' **     Field: [Schedule_ID]
                              ' **   Table: tmpXAdmin_Account_01          {Local}
                              ' **     Field: [Schedule_ID]
                              ' **   Table: tmpXAdmin_Account_02          {Local}
                              ' **     Field: [Schedule_ID]
16230                         .MoveLast
16240                         lngRecs = .RecordCount
16250                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
16260                         .MoveFirst
16270                         For Each fld In .Fields
16280                           With fld
16290                             If .Name = "Schedule ID" Then
                                    ' ** Has old field names.
16300                               blnFound2 = True
16310                               Exit For
16320                             End If
16330                           End With
16340                         Next
16350                         For lngX = 1& To lngRecs
16360                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
                                ' ** Add the record to the new table.
16370                           rstLoc1.AddNew
16380                           strTmp04 = vbNullString
16390                           strTmp09 = "Schedule_Name": strTmp10 = "Schedule Name"
16400                           If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
16410                           If IsNull(.Fields(strTmp08)) = True Then
16420                             lngDupeNum = lngDupeNum + 1&
16430                             strTmp04 = "UNKNOWN " & CStr(lngDupeNum)
16440                             lngDupeUnks = lngDupeUnks + 1&
16450                             ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
16460                             arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
16470                             arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
16480                           Else
16490                             If Trim(.Fields(strTmp08)) = vbNullString Then
16500                               lngDupeNum = lngDupeNum + 1&
16510                               strTmp04 = "UNKNOWN " & CStr(lngDupeNum)
16520                               lngDupeUnks = lngDupeUnks + 1&
16530                               ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
16540                               arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
16550                               arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
16560                             Else
16570                               strTmp04 = Trim(.Fields(strTmp08))
16580                             End If
16590                           End If
16600                           rstLoc1![Schedule_Name] = strTmp04
16610                           Select Case blnFound2
                                Case True
16620                             rstLoc1![Schedule_Base] = ![base]
16630                             rstLoc1![Schedule_Minimum] = ![minimum]
16640                             rstLoc1![Schedule_DateCreated] = Now()
16650                             rstLoc1![Schedule_DateModified] = Now()
16660                           Case False
16670                             rstLoc1![Schedule_Base] = ![Schedule_Base]
16680                             rstLoc1![Schedule_Minimum] = ![Schedule_Minimum]
16690                             rstLoc1![Schedule_DateCreated] = ![Schedule_DateCreated]
16700                             rstLoc1![Schedule_DateModified] = ![Schedule_DateModified]
16710                           End Select
16720 On Error Resume Next
16730                           rstLoc1.Update
16740                           If ERR.Number <> 0 Then
16750                             If ERR.Number = 3022 Then
                                    ' ** Error 3022: The changes you requested to the table were not successful because they
                                    ' **             would create duplicate values in the index, primary key, or relationship.
16760                               If gblnDev_NoErrHandle = True Then
16770 On Error GoTo 0
16780                               Else
16790 On Error GoTo ERRH
16800                               End If
16810                               lngDupeNum = lngDupeNum + 1&
16820                               strTmp09 = "Schedule_Name": strTmp10 = "Schedule Name"
16830                               If blnFound2 = False Then strTmp08 = strTmp09 Else strTmp08 = strTmp10
16840                               rstLoc1![Schedule_Name] = Trim(Nz(.Fields(strTmp08), vbNullString)) & " #DUPE" & CStr(lngDupeNum)
16850                               rstLoc1.Update
16860                               lngDupeUnks = lngDupeUnks + 1&
16870                               ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
16880                               arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "DUP"
16890                               arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
16900                             Else
16910                               intRetVal = -6
16920                               blnContinue = False
16930                               lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
16940                               MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                      "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                      vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
16950                               rstLoc1.CancelUpdate
16960                               If gblnDev_NoErrHandle = True Then
16970 On Error GoTo 0
16980                               Else
16990 On Error GoTo ERRH
17000                               End If
17010                             End If
17020                           Else
17030                             If gblnDev_NoErrHandle = True Then
17040 On Error GoTo 0
17050                             Else
17060 On Error GoTo ERRH
17070                             End If
17080                           End If
17090                           If blnContinue = True Then
17100                             rstLoc1.Bookmark = rstLoc1.LastModified
                                  ' ** Add the ID cross-reference to tblVersion_Key.
17110                             lngSchedules = lngSchedules + 1&
17120                             lngE = lngSchedules - 1&
17130                             ReDim Preserve arr_varSchedule(S_ELEMS, lngE)
17140                             strTmp06 = "Schedule_ID": strTmp07 = "Schedule ID"
17150                             If blnFound2 = False Then strTmp05 = strTmp06 Else strTmp05 = strTmp07
17160                             arr_varSchedule(S_ID_OLD, lngE) = .Fields(strTmp05)
17170                             arr_varSchedule(S_ID_NEW, lngE) = rstLoc1![Schedule_ID]
17180                             arr_varSchedule(S_DETS, lngE) = CLng(0)  ' ** Detail records.
17190                             rstLoc2.AddNew
17200                             rstLoc2![tbl_id] = lngCurrTblID
17210                             rstLoc2![tbl_name] = strCurrTblName  'Schedule
17220                             rstLoc2![fld_id] = lngCurrKeyFldID
17230                             rstLoc2![fld_name] = strCurrKeyFldName
17240                             rstLoc2![key_lng_id1] = .Fields(strTmp05)
                                  'rstLoc2![key_txt_id1] =
17250                             rstLoc2![key_lng_id2] = rstLoc1![Schedule_ID]
                                  'rstLoc2![key_txt_id2] =
17260                             rstLoc2.Update
                                  ' ** Cross-check Schedule_ID's against those referenced in Account table.
17270                             For lngY = 0& To (lngAccts - 1&)
17280                               If IsNull(arr_varAcct(A_SCHED, lngY)) = False Then
17290                                 If arr_varAcct(A_SCHED, lngY) = .Fields(strTmp05) Then
17300                                   arr_varAcct(A_SCHED_N, lngY) = rstLoc1![Schedule_ID]
17310                                 End If
17320                               End If
17330                             Next
17340                           Else
17350                             Exit For
17360                           End If
17370                           If lngX < lngRecs Then .MoveNext
17380                         Next
17390                         If blnContinue = True Then
                                ' ** Check if all referenced Schedule_ID's in Account table are present.
17400                           For lngX = 0& To (lngAccts - 1&)
17410                             If IsNull(arr_varAcct(A_SCHED, lngX)) = False Then
17420                               If arr_varAcct(A_SCHED_N, lngX) = 0& Then
                                      ' ** Schedule_ID in account table not found in Schedule table.
17430                                 With rstLoc1
                                        ' ** Add the record to the new table.
17440                                   .AddNew
17450                                   strTmp04 = "UNKNOWN" & CStr(arr_varAcct(A_SCHED, lngX))
17460                                   strTmp04 = Left(strTmp04 & " Acct# " & arr_varAcct(A_NUM, lngX), 25)  ' ** [accountno] is max 15 chars.
17470                                   ![Schedule_Name] = strTmp04
17480                                   ![Schedule_Base] = CDbl(0)
17490                                   ![Schedule_Minimum] = CDbl(0)
17500                                   ![Schedule_DateCreated] = Now()
17510                                   ![Schedule_DateModified] = Now()
17520                                   .Update
17530                                   lngDupeUnks = lngDupeUnks + 1&
17540                                   ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
17550                                   arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "UNK"
17560                                   arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
17570                                   .Bookmark = .LastModified
17580                                   strTmp06 = "Schedule_ID": strTmp07 = "Schedule ID"
17590                                   If blnFound2 = False Then strTmp05 = strTmp06 Else strTmp05 = strTmp07
17600                                   arr_varAcct(A_SCHED_N, lngX) = .Fields(strTmp05)
                                        ' ** Update the rest of arr_varAcct().
17610                                   For lngY = 0& To (lngAccts - 1&)
17620                                     If lngY <> lngX Then  ' ** Skip the one we're on.
17630                                       If IsNull(arr_varAcct(A_SCHED, lngY)) = False Then
17640                                         If arr_varAcct(A_SCHED, lngY) = arr_varAcct(A_SCHED, lngX) Then
17650                                           arr_varAcct(A_SCHED_N, lngY) = arr_varAcct(A_SCHED_N, lngX)
17660                                         End If
17670                                       End If
17680                                     End If
17690                                   Next
17700                                 End With  ' ** rstLoc1.
17710                                 With rstLoc2
                                        ' ** Add the ID cross-reference to tblVersion_Key.
17720                                   .AddNew
17730                                   ![tbl_id] = lngCurrTblID
17740                                   ![tbl_name] = strCurrTblName  'Schedule
17750                                   ![fld_id] = lngCurrKeyFldID
17760                                   ![fld_name] = strCurrKeyFldName
17770                                   ![key_lng_id1] = arr_varAcct(A_SCHED, lngX)
                                        '![key_txt_id1] =
17780                                   ![key_lng_id2] = arr_varAcct(A_SCHED_N, lngX)
                                        '![key_txt_id2] =
17790                                   .Update
17800                                 End With  ' ** rstLoc2.
                                      ' ** Its detail record will be handled below.
17810                               End If
17820                             End If
17830                           Next
17840                         End If  ' ** blnContinue.
17850                         rstLoc1.Close
17860                         rstLoc2.Close
17870                       End If  ' ** Records present.
17880                       .Close
17890                     End With  ' ** rstLnk.
17900                   End If  ' ** blnFound.
17910                   strTmp04 = vbNullString
17920                   strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
17930                   strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString

17940                   If blnContinue = True And blnFound = True Then
17950                     If lngSchedules > 0& Then

17960                       strCurrTblName = "ScheduleDetail"
17970                       lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                              "[tbl_name] = '" & strCurrTblName & "'")

17980                       blnFound = False: blnFound2 = False: lngRecs = 0&
17990                       strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
18000                       strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
18010                       For lngX = 0& To (lngOldTbls - 1&)
18020                         If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
18030                           blnFound = True
18040                           Exit For
18050                         End If
18060                       Next

18070                       If blnFound = True Then
18080                         lngRecs = 0&
18090                         Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
18100                         With rstLnk
18110                           If .BOF = True And .EOF = True Then
                                  ' ** Not used.
                                  ' ** We do want at least 1 detail record for each Schedule!
18120                             blnFound = False
18130                           Else
18140                             strCurrKeyFldName = "ScheduleDetail_ID"
18150                             lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                    "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                    "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
18160                             Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
                                  'Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                                  ' ** Table: ScheduleDetail  {all new field names}
                                  ' **   ![Schedule_ID]
                                  ' **   ![ScheduleDetail_ID]
                                  ' **   ![ScheduleDetail_Rate]
                                  ' **   ![ScheduleDetail_Amount]
                                  ' **   ![ScheduleDetail_DateModified]
                                  ' ** Referenced by:
                                  ' **   Table: Schedule                      {Linked}
                                  ' **     Field: [Schedule_ID]
18170                             .MoveLast
18180                             lngRecs = .RecordCount
18190                             .MoveFirst
18200                             For Each fld In .Fields
18210                               With fld
18220                                 If .Name = "Schedule ID" Then
                                        ' ** Has old field names.
18230                                   blnFound2 = True
18240                                   Exit For
18250                                 End If
18260                               End With
18270                             Next
18280                             For lngX = 1& To lngRecs
18290                               strTmp06 = "Schedule_ID": strTmp07 = "Schedule ID"
18300                               If blnFound2 = False Then strTmp05 = strTmp06 Else strTmp05 = strTmp07
18310                               lngTmp13 = Nz(.Fields(strTmp05), 0)
18320                               lngTmp14 = 0&
18330                               If lngTmp13 > 0& Then
18340                                 For lngY = 0& To (lngSchedules - 1&)
18350                                   If arr_varSchedule(S_ID_OLD, lngY) = lngTmp13 Then
18360                                     lngTmp14 = arr_varSchedule(S_ID_NEW, lngY)
18370                                     arr_varSchedule(S_DETS, lngY) = arr_varSchedule(S_DETS, lngY) + 1&
18380                                     Exit For
18390                                   End If
18400                                 Next
18410                                 If lngTmp14 > 0& Then
                                        ' ** Add the record to the new table.
18420                                   rstLoc1.AddNew
18430                                   rstLoc1![Schedule_ID] = lngTmp14
18440                                   Select Case blnFound2
                                        Case True
18450                                     rstLoc1![ScheduleDetail_Rate] = ![rate]
18460                                     rstLoc1![ScheduleDetail_Amount] = ![amount]
18470                                     rstLoc1![ScheduleDetail_DateModified] = Now()
18480                                   Case False
18490                                     rstLoc1![ScheduleDetail_Rate] = ![ScheduleDetail_Rate]
18500                                     rstLoc1![ScheduleDetail_Amount] = ![ScheduleDetail_Amount]
18510                                     rstLoc1![ScheduleDetail_DateModified] = ![ScheduleDetail_DateModified]
18520                                   End Select
18530                                   rstLoc1.Update
18540                                 Else
                                        ' ** An orphan detail record; throw out.
18550                                 End If
18560                               Else
                                      ' ** An orphan detail record; throw out.
18570                               End If
18580                               If lngX < lngRecs Then .MoveNext
18590                             Next
18600                             For lngX = 0& To (lngSchedules - 1&)
18610                               If arr_varSchedule(S_DETS, lngY) = 0& Then
                                      ' ** Any one missing a detail will cause it to be checked below.
18620                                 blnFound = False
18630                                 Exit For
18640                               End If
18650                             Next
18660                             rstLoc1.Close
                                  'rstLoc2.Close
18670                           End If  ' ** Records present.
18680                           If blnFound = False Then
                                  ' ** Since we're here because there are Schedule records, this
                                  ' ** means that none have detail (or at least one is missing a detail), so we have to create them.
                                  ' **   "... enter a '0' in the rate and '1' in the amount columns."
18690                             Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
18700                             With rstLoc1
18710                               For lngX = 0& To (lngSchedules - 1&)
18720                                 If arr_varSchedule(S_DETS, lngY) = 0& Then
                                        ' ** Add the record to the new table.
18730                                   rstLoc1.AddNew
18740                                   rstLoc1![Schedule_ID] = arr_varSchedule(S_ID_NEW, lngY)
18750                                   rstLoc1![ScheduleDetail_Rate] = 0#
18760                                   rstLoc1![ScheduleDetail_Amount] = 1#
18770                                   rstLoc1![ScheduleDetail_DateModified] = Now()
18780                                   rstLoc1.Update
18790                                   arr_varSchedule(S_DETS, lngY) = 1&
18800                                 End If
18810                               Next
18820                               .Close
18830                             End With  ' ** rstLoc1.
18840                           End If
18850                           .Close
18860                         End With  ' ** rstLnk.
18870                       Else
                              ' ** I don't think it's possible for Schedule to be present, and ScheduleDetail to be missing!
18880                       End If  ' ** blnFound.
18890                     Else
                            ' ** Since there are no Schedules, no need for ScheduleDetail
18900                     End If  ' ** lngSchedules > 0.
18910                   End If  ' ** blnContinue.
18920                   strTmp05 = vbNullString: strTmp06 = vbNullString: strTmp07 = vbNullString
18930                   strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString

18940                 End If  ' ** blnContinue.

18950                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: DDate.
                        ' **        PostingDate.
                        ' **        Statement Date.
                        ' ******************************

                        ' ** Step 7: DDate, PostingDate, Statement Date.
18960                   dblPB_ThisStep = 7#
18970                   Version_Status 3, dblPB_ThisStep, "Posting / Statement Date"  ' ** Function: Below.

                        ' ** DDate is unnecessary.
                        ' ** Only used in preparing for frmAccruedIncome.
                        'strCurrTblName = "DDate"
                        'lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                        '  "[tbl_name] = '" & strCurrTblName & "'")

18980                   strCurrTblName = "PostingDate"
18990                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

19000                   blnFound = False
19010                   For lngX = 0& To (lngOldTbls - 1&)
19020                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
19030                       blnFound = True
19040                       Exit For
19050                     End If
19060                   Next

19070                   Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, 1&  ' ** Function: Below.
19080                   Version_Status 4, dblPB_ThisStep, strCurrTblName, 1&, 1&  ' ** Function: Below.

19090                   If blnFound = True Then
19100                     Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
19110                     With rstLoc1
19120                       If .BOF = True And .EOF = True Then
                              ' ** There should always be at least one record here!
19130                         .AddNew
19140                         .Update
19150                       End If
19160                     End With  ' ** rstLoc1.
19170                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
19180                     With rstLnk
19190                       If .BOF = True And .EOF = True Then
                              ' ** There should always be at least one record here!
19200                       Else
19210                         .MoveFirst
19220                         strTmp04 = vbNullString
19230                         For Each fld In .Fields
19240                           If fld.Name = "Posting Date" Then  '#### DONE POSTINGDATE
19250                             strTmp04 = "Posting Date"
19260                             Exit For
19270                           ElseIf fld.Name = "Posting_Date" Then
19280                             strTmp04 = "Posting_Date"
19290                             Exit For
19300                           End If
19310                         Next
19320                         If IsNull(.Fields(strTmp04)) = False Then
19330                           rstLoc1.MoveFirst
19340                           rstLoc1.Edit
19350                           rstLoc1![Posting_Date] = .Fields(strTmp04)
19360                           rstLoc1.Update
19370                         End If
19380                       End If
19390                       .Close
19400                     End With  ' ** rstLnk.
19410                     rstLoc1.Close
19420                   End If  ' ** blnFound.

19430                   strCurrTblName = "Statement Date"
19440                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

19450                   blnFound = False
19460                   For lngX = 0& To (lngOldTbls - 1&)
19470                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
19480                       blnFound = True
19490                       Exit For
19500                     End If
19510                   Next

19520                   If blnFound = True Then
19530                     Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
19540                     With rstLoc1
19550                       If .BOF = True And .EOF = True Then
                              ' ** There should always be at least one record here!
19560                         .AddNew
19570                         .Update
19580                       End If
19590                     End With  ' ** rstLoc1.
19600                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
19610                     With rstLnk
19620                       If .BOF = True And .EOF = True Then
                              ' ** There should always be at least one record here!
19630                       Else
19640                         .MoveFirst
19650                         If IsNull(![Statement_Date]) = False Then
19660                           rstLoc1.MoveFirst
19670                           rstLoc1.Edit
19680                           rstLoc1![Statement_Date] = ![Statement_Date]
19690                           rstLoc1.Update
19700                         End If
19710                       End If
19720                       .Close
19730                     End With  ' ** rstLnk.
19740                     rstLoc1.Close
19750                   End If  ' ** blnFound.

19760                 End If  ' ** blnContinue.

19770                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: m_REVCODE.
                        ' ******************************

                        ' ** Step 8: m_REVCODE.
19780                   dblPB_ThisStep = 8#
19790                   Version_Status 3, dblPB_ThisStep, "Income/Expense Codes"  ' ** Function: Below.

19800                   strCurrTblName = "m_REVCODE"
19810                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

19820                   blnFound = False: lngRecs = 0&
19830                   For lngX = 0& To (lngOldTbls - 1&)
19840                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
19850                       blnFound = True
19860                       Exit For
19870                     End If
19880                   Next

19890                   If blnFound = True Then
19900                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
19910                     With rstLnk
19920                       If .BOF = True And .EOF = True Then
                              ' ** If they've got it, there should be at least one record!
19930                       Else
19940                         strCurrKeyFldName = "revcode_ID"
19950                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
19960                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
19970                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Ver 1.x didn't have it!
                              ' ** Of 11 example TrustDta.mdb's, all have 5 fields.
                              ' ** Current field count is 5 fields.
                              ' ** Table: m_REVCODE
                              ' **   ![revcode_ID]
                              ' **   ![revcode_DESC]
                              ' **   ![revcode_TYPE]
                              ' **   ![revcode_SORTORDER]
                              ' **   ![revcode_ACTIVE]
                              ' ** Referenced by:
                              ' **   Table: journal                       {Linked}
                              ' **     Field: [revcode_ID]
                              ' **   Table: ledger                        {Linked}
                              ' **     Field: [revcode_ID]
                              ' **   Table: LedgerArchive                 {Linked}
                              ' **     Field: [revcode_ID]
                              ' **   Table: tblTemplate_Journal           {Local}
                              ' **     Field: [revcode_ID]
                              ' **   Table: tblTemplate_Ledger            {Local}
                              ' **     Field: [revcode_ID]
                              ' **   Table: tblTemplate_m_REVCODE         {Local}
                              ' **     Field: [revcode_ID]
                              ' **   Table: tmpCourtReportData            {Local}
                              ' **     Field: [revcode_ID]
                              ' **   Table: tmpCourtReportData2           {Local}
                              ' **     Field: [revcode_ID]
                              ' **   Table: tmpCourtReportData6           {Local}
                              ' **     Field: [revcode_ID]
                              ' **   Table: tmpEdit02                     {Local}
                              ' **     Field: [revcode_ID]
                              ' ** Make sure these defaults don't get brought over a 2nd time, as well as checking their ID's.
                              ' **   revcode_ID  revcode_DESC         revcode_TYPE  revcode_SORTORDER  revcode_ACTIVE
                              ' **   ==========  ===================  ============  =================  ==============
                              ' **   1           Unspecified Income   1             1                  True
                              ' **   2           Unspecified Expense  2             1                  True
                              ' **   3           OC Other Charges     1             2                  True
                              ' **   4           OC Other Credits     2             2                  True
                              ' **   5           Ordinary Dividend    1             3                  True
                              ' **   6           Interest Income      1             4                  True
19980                         .MoveLast
19990                         lngRecs = .RecordCount
20000                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
20010                         .MoveFirst
                              ' ** Key errors here are most likely because of revcode_SORTORDER, not revcode_DESC.
                              ' ** Let's see if we can resort them before putting them into the new table!
                              ' ** Collect their list.
20020                         For lngX = 1& To lngRecs
20030                           lngRevCodes = lngRevCodes + 1&
20040                           lngE = lngRevCodes - 1&
20050                           ReDim Preserve arr_varRevCode(R_ELEMS, lngE)
20060                           arr_varRevCode(R_REC, lngE) = lngX                                      'Long
20070                           arr_varRevCode(R_ID, lngE) = ![revcode_ID]                              'Long
20080                           arr_varRevCode(R_DSC, lngE) = Trim(Nz(![revcode_DESC], vbNullString))  'String
20090                           arr_varRevCode(R_TYP, lngE) = Nz(![revcode_TYPE], 1)                    'Long
20100                           arr_varRevCode(R_ORD, lngE) = Nz(![revcode_SORTORDER], 80)  'arbitrary  'Long
20110                           arr_varRevCode(R_ACT, lngE) = ![revcode_ACTIVE]                         'Boolean
20120                           arr_varRevCode(R_NSO, lngE) = CLng(-1)        'New Sort Order           'Long
20130                           arr_varRevCode(R_NID, lngE) = CLng(0)         'New ID                   'Long
20140                           arr_varRevCode(R_EIM, lngE) = CLng(-1)        'Element of ?             'Long
20150                           arr_varRevCode(R_DEL, lngE) = CBool(False)                              'Boolean
20160                           arr_varRevCode(R_FND, lngE) = CBool(False)                              'Boolean
20170                           If lngX < lngRecs Then .MoveNext
20180                         Next
                              'WE'VE GOT THEIR ORIGINAL ID, AND THEIR ORIGINAL SORT ORDER,
                              'AS WELL AS THEIR DESCRIPTION FOR COMPARISSON!

                              ' ** arr_varRevDefCode() is the new Default revcode_ID cross-reference.
                              ' ** tblVersion_Key is the new saved revcode_ID cross-reference.
                              ' ** arr_varRevCode() has any dupes, and their revcode_ID cross-reference.
                              ' ** Tables with revcode_ID:
                              ' **   Journal
                              ' **   Ledger
                              ' **   LedgerArchive

                              ' ** m_REVCODE now has two unique indexes, the 2nd of which wasn't there before:
                              ' **   1. [revcode_TYPE], [revcode_SORTORDER]  ' ** Taken care of by new sort order (R_NSO).
                              ' **   2. [revcode_DESC], [revcode_TYPE]
                              ' ** Example of customer's dupes:
                              ' **   revcode_DESC      revcode_TYPE  revcode_ID  revcode_SORTORDER  revcode_ACTIVE
                              ' **   ================  ============  ==========  =================  ==============
                              ' **   Health Insurance  2             115         12                 Yes
                              ' **   Health insurance  2             113         9                  No
                              ' **   Other Income      1             108         30                 Yes
                              ' **   Other Income      1             5           6                  No
                              ' **   Trustee Fees      2             155         51                 Yes
                              ' **   Trustee Fees      2             3           3                  Yes

                              ' ** Check for [revcode_DESC]/[revcode_TYPE] dupes.
                              ' ** This just looks at dupes in their list, and doesn't
                              ' ** consider the new 'Ordinary Dividend' nor 'Interest Income'.
20190                         For lngX = 0& To (lngRevCodes - 1&)
20200                           If arr_varRevCode(R_DEL, lngX) = False Then
20210                             strTmp04 = CStr(arr_varRevCode(R_TYP, lngX)) & " " & arr_varRevCode(R_DSC, lngX)
20220                             blnTmp22 = arr_varRevCode(R_ACT, lngX)
20230                             For lngY = 0& To (lngRevCodes - 1&)
20240                               If lngY <> lngX And arr_varRevCode(R_DEL, lngY) = False Then
20250                                 If (CStr(arr_varRevCode(R_TYP, lngY)) & " " & arr_varRevCode(R_DSC, lngY)) = strTmp04 Then
20260                                   If blnTmp22 = False And arr_varRevCode(R_ACT, lngY) = True Then
                                          ' ** The first one encountered is inactive, but this one is active.
                                          ' ** Switch them around, since we'll be giving them all new revcode_ID's anyway.
20270                                     arr_varRevCode(R_ACT, lngY) = blnTmp22  ' ** This one now inactive.
20280                                     blnTmp22 = True
20290                                     arr_varRevCode(R_ACT, lngX) = blnTmp22  ' ** That one now active.
20300                                   End If
20310                                   arr_varRevCode(R_DEL, lngY) = CBool(True)
20320                                   arr_varRevCode(R_EIM, lngY) = lngX
20330                                 End If
20340                               End If
20350                             Next
20360                           End If
20370                         Next
                              'IF DUPES WERE FOUND, THE 2ND ONE GETS POINTED TO THE 1ST VIA THE R_EIM ELEMENT!
                              'lngTmp15
                              'lngTmp16
                              'strTmp04
                              'lngTmp17
                              'lngTmp18
                              'blnTmp22
                              'lngTmp19
                              'lngTmp20
                              'lngTmp21
                              'blnTmp23
                              'blnTmp24
                              'lngTmp15 = 0&: lngTmp16 = 0&: lngTmp17 = 0&: lngTmp18 = 0&: lngTmp19 = 0&: lngTmp20 = 0&: lngTmp21 = 0&:
                              'blnTmp22 = False: blnTmp23 = False: blnTmp24 = False

                              ' ** Binary Sort arr_varRevCode() array (alphabetize).
                              'DO SORT ORDER INSTEAD!
20380                         For lngX = UBound(arr_varRevCode, 2) To 1& Step -1
20390                           For lngY = 0 To (lngX - 1)
20400                             If arr_varRevCode(R_ORD, lngY) > arr_varRevCode(R_ORD, (lngY + 1)) Then
20410                               varTmp00 = Empty
20420                               For lngZ = 0& To R_ELEMS
20430                                 varTmp00 = arr_varRevCode(lngZ, lngY)
20440                                 arr_varRevCode(lngZ, lngY) = arr_varRevCode(lngZ, (lngY + 1))
20450                                 arr_varRevCode(lngZ, (lngY + 1)) = varTmp00
20460                                 varTmp00 = Empty
20470                               Next
20480                             End If
20490                           Next
20500                         Next

                              ' ***********************************************
                              ' ***********************************************
                              ' ** Look for existing Dividend/Interest codes!
                              ' ***********************************************

                              ' ***********************************************
                              ' ** Dividend.
                              ' ***********************************************
20510                         blnTmp26 = False: blnTmp27 = False
20520                         For lngX = 0& To (lngRevCodes - 1&)
20530                           If arr_varRevCode(R_TYP, lngX) = REVTYP_INC Then  ' ** Only check Income.
20540                             Select Case arr_varRevCode(R_DSC, lngX)
                                  Case "Ordinary Dividend"
                                    ' ** This is one of ours.
20550                               blnTmp26 = True
20560                             Case "Dividend", "Dividends", "Dividend Income", "Dividends Income", "Income Dividend", _
                                      "Income Dividends", "Income From Dividends", "Ordinary Dividends"
                                    ' ** This is one of theirs.
20570                               blnTmp27 = True
20580                             End Select
20590                           End If
20600                         Next  ' ** lngX.
20610                         If blnTmp26 = True And blnTmp27 = False Then
                                ' ** All they have is ours. (Or they happened to choose the same name.)
20620                         ElseIf blnTmp26 = False And blnTmp27 = True Then
                                ' ** We'll take over theirs.
20630                           For lngX = 0& To (lngRevCodes - 1&)
20640                             If arr_varRevCode(R_TYP, lngX) = REVTYP_INC Then  ' ** Only check Income.
20650                               Select Case arr_varRevCode(R_DSC, lngX)
                                    Case "Dividend", "Dividends", "Dividend Income", "Dividends Income", "Income Dividend", _
                                        "Income Dividends", "Income From Dividends", "Ordinary Dividends"
                                      ' ** Just take the first hit.
20660                                 arr_varRevCode(R_DSC, lngX) = "Ordinary Dividend"
20670                                 Exit For
20680                               End Select
20690                             End If
20700                           Next  ' ** lngX.
20710                         ElseIf blnTmp26 = False And blnTmp27 = False Then
                                ' ** We'll have to add ours.
20720                         ElseIf blnTmp26 = True And blnTmp27 = True Then
                                ' ** They've got both, and that's OK, too.
20730                         End If

                              ' ***********************************************
                              ' ** Interest.
                              ' ***********************************************
20740                         blnTmp26 = False: blnTmp27 = False
20750                         For lngX = 0& To (lngRevCodes - 1&)
20760                           If arr_varRevCode(R_TYP, lngX) = REVTYP_INC Then  ' ** Only check Income.
20770                             Select Case arr_varRevCode(R_DSC, lngX)
                                  Case "Interest Income"
                                    ' ** This is one of ours.
20780                               blnTmp26 = True
20790                             Case "Interest", "Interests", "Income Interest", "Income From Interest", "Ordinary Interest"
                                    ' ** This is one of theirs.
20800                               blnTmp27 = True
20810                             End Select
20820                           End If
20830                         Next  ' ** lngX.
20840                         If blnTmp26 = True And blnTmp27 = False Then
                                ' ** All they have is ours. (Or they happened to choose the same name.)
20850                         ElseIf blnTmp26 = False And blnTmp27 = True Then
                                ' ** We'll take over theirs.
20860                           For lngX = 0& To (lngRevCodes - 1&)
20870                             If arr_varRevCode(R_TYP, lngX) = REVTYP_INC Then  ' ** Only check Income.
20880                               Select Case arr_varRevCode(R_DSC, lngX)
                                    Case "Interest", "Interests", "Income Interest", "Income From Interest", "Ordinary Interest"
                                      ' ** Just take the first hit.
20890                                 arr_varRevCode(R_DSC, lngX) = "Interest Income"
20900                                 Exit For
20910                               End Select
20920                             End If
20930                           Next  ' ** lngX.
20940                         ElseIf blnTmp26 = False And blnTmp27 = False Then
                                ' ** We'll have to add ours.
20950                         ElseIf blnTmp26 = True And blnTmp27 = True Then
                                ' ** They've got both, and that's OK, too.
20960                         End If
                              ' ***********************************************
                              ' ***********************************************

20970                         strTmp04 = vbNullString
                              ' ** So, what if one of the Switches/Dupes is one of the defaults?
                              ' ** Whichever one is active (or switched to active) gets saved, and if it's
                              ' ** a dupe, it's now referring to the first one. So, I think we're covered.
                              ' ** First, renumber the defaults.
20980                         lngTmp13 = 0&: lngTmp14 = 0&: blnTmp23 = False
                              ' ** blnTmp23 = Yes, a change was made in a Default.
                              ' ** lngTmp13 = Income sort order.
                              ' ** lngTmp14 = Expense sort order.
20990                         For lngX = 0& To (lngRevCodes - 1&)
21000                           If arr_varRevCode(R_DEL, lngX) = False Then
21010                             If arr_varRevCode(R_DSC, lngX) = "Unspecified Income" Or arr_varRevCode(R_DSC, lngX) = "Unspecified Expense" Or _
                                      arr_varRevCode(R_DSC, lngX) = "OC Other Charges" Or arr_varRevCode(R_DSC, lngX) = "OC Other Credits" Or _
                                      arr_varRevCode(R_DSC, lngX) = "Ordinary Dividend" Or arr_varRevCode(R_DSC, lngX) = "Interest Income" Then
21020                               Select Case arr_varRevCode(R_DSC, lngX)
                                    Case "Unspecified Income"   ' ** lngTmp13 = Income = 1.
21030                                 If arr_varRevCode(R_ORD, lngX) <> 1& Then
21040                                   blnTmp23 = True
21050                                   arr_varRevCode(R_NSO, lngX) = 1&
21060                                 End If
21070                                 lngTmp13 = 1&
21080                               Case "Unspecified Expense"  ' ** lngTmp14 = Expense = 2.
21090                                 If arr_varRevCode(R_ORD, lngX) <> 1& Then
21100                                   blnTmp23 = True
21110                                   arr_varRevCode(R_NSO, lngX) = 1&
21120                                 End If
21130                                 lngTmp14 = 1&
21140                               Case "OC Other Charges"
21150                                 If arr_varRevCode(R_NSO, lngX) <> 2& Then  ' ** Regardless of their original sort order.
21160                                   blnTmp23 = True
21170                                   arr_varRevCode(R_NSO, lngX) = 2&
21180                                 End If
21190                                 lngTmp13 = REVID_EXP  ' ** Since we sorted above, we shouldn't encounter an Unspecified after this.
21200                               Case "OC Other Credits"
21210                                 If arr_varRevCode(R_NSO, lngX) <> 2& Then  ' ** Regardless of their original sort order.
21220                                   blnTmp23 = True
21230                                   arr_varRevCode(R_NSO, lngX) = 2&
21240                                 End If
21250                                 lngTmp14 = 2&  ' ** Since we sorted above, we shouldn't encounter an Unspecified after this.
21260                               Case "Ordinary Dividend"
21270                                 If arr_varRevCode(R_NSO, lngX) <> 3& Then  ' ** Regardless of their original sort order.
21280                                   blnTmp23 = True
21290                                   arr_varRevCode(R_NSO, lngX) = 3&
21300                                 End If
21310                                 lngTmp13 = 3&
21320                               Case "Interest Income"
21330                                 If arr_varRevCode(R_NSO, lngX) <> 4& Then  ' ** Regardless of their original sort order.
21340                                   blnTmp23 = True
21350                                   arr_varRevCode(R_NSO, lngX) = 4&
21360                                 End If
21370                                 lngTmp13 = 4&
21380                               End Select
21390                             End If
21400                           End If
21410                         Next
21420                         If lngTmp13 < 2& Then
21430                           blnTmp23 = True
21440                           lngTmp13 = 4&  ' ** Leave room for the OC's, Div, and Int if they didn't have them.
21450                         End If
21460                         If lngTmp14 < 2& Then
21470                           blnTmp23 = True
21480                           lngTmp14 = 2&
21490                         End If
21500                         If blnTmp23 = True Then
                                ' ** Only renumber if we have to!
21510                           For lngX = 0& To (lngRevCodes - 1&)
                                  ' ** Since I sorted by R_ORD, this should keep them in the same order!
21520                             If arr_varRevCode(R_DEL, lngX) = False Then
21530                               If arr_varRevCode(R_NSO, lngX) = -1& Then
21540                                 Select Case arr_varRevCode(R_TYP, lngX)
                                      Case 1&   ' ** lngTmp13 = Income = 1.
21550                                   lngTmp13 = lngTmp13 + 1&
21560                                   arr_varRevCode(R_NSO, lngX) = lngTmp13
21570                                 Case 2&  ' ** lngTmp14 = Expense = 2.
21580                                   lngTmp14 = lngTmp14 + 1&
21590                                   arr_varRevCode(R_NSO, lngX) = lngTmp14
21600                                 End Select
21610                               End If
21620                             End If
21630                           Next
                                'THEIR SORT ORDER MAY BE DIFFERENT, BUT THEIR ID'S ARE, SO FAR, STILL THE SAME!
21640                         End If

21650                         .MoveFirst  ' ** Though now we don't move through them again.
21660                         blnTmp24 = False: blnTmp25 = False
                              ' ** blnTmp24 = Yes, a defualt ID was changed.
21670                         For lngX = 0& To (lngRevCodes - 1&)
21680                           If arr_varRevCode(R_DEL, lngX) = False Then
21690                             Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX + 1&, lngRevCodes  ' ** Function: Below.
21700                             If arr_varRevCode(R_DSC, lngX) = "Unspecified Income" Or arr_varRevCode(R_DSC, lngX) = "Unspecified Expense" Or _
                                      arr_varRevCode(R_DSC, lngX) = "OC Other Charges" Or arr_varRevCode(R_DSC, lngX) = "OC Other Credits" Or _
                                      arr_varRevCode(R_DSC, lngX) = "Ordinary Dividend" Or arr_varRevCode(R_DSC, lngX) = "Interest Income" Then
                                    'blnTmp23 SAYS WHETHER THEIR SORT ORDER HAD TO BE CHANGED,
                                    'BUT THIS STILL HAS TO CHECK WHETHER THEIR ID MAY HAVE TO CHANGE AS WELL!
21710                               Select Case arr_varRevCode(R_DSC, lngX)
                                    Case "Unspecified Income"
21720                                 If arr_varRevCode(R_ID, lngX) <> 1& Then
21730                                   blnTmp24 = True
21740                                   lngRevDefCodes = lngRevDefCodes + 1&
21750                                   lngE = lngRevDefCodes - 1&
21760                                   ReDim Preserve arr_varRevDefCode(RD_ELEMS, lngE)
21770                                   arr_varRevDefCode(RD_ID_OLD, lngE) = arr_varRevCode(R_ID, lngX)
21780                                   arr_varRevDefCode(RD_ID_NEW, lngE) = CLng(1)
21790                                   arr_varRevDefCode(RD_DSC, lngE) = "Unspecified Income"
                                        'HERE'S WHERE THEY'LL GET A NEW ID!
21800                                 End If
21810                                 arr_varRevCode(R_NID, lngX) = 1&
21820                               Case "Unspecified Expense"
21830                                 If arr_varRevCode(R_ID, lngX) <> 2& Then
21840                                   blnTmp24 = True
21850                                   lngRevDefCodes = lngRevDefCodes + 1&
21860                                   lngE = lngRevDefCodes - 1&
21870                                   ReDim Preserve arr_varRevDefCode(RD_ELEMS, lngE)
21880                                   arr_varRevDefCode(RD_ID_OLD, lngE) = arr_varRevCode(R_ID, lngX)
21890                                   arr_varRevDefCode(RD_ID_NEW, lngE) = CLng(2)
21900                                   arr_varRevDefCode(RD_DSC, lngE) = "Unspecified Expense"
                                        'HERE'S WHERE THEY'LL GET A NEW ID!
21910                                 End If
21920                                 arr_varRevCode(R_NID, lngX) = 2&
21930                               Case "OC Other Charges"
21940                                 If arr_varRevCode(R_ID, lngX) <> 3& Then
21950                                   blnTmp24 = True
21960                                   lngRevDefCodes = lngRevDefCodes + 1&
21970                                   lngE = lngRevDefCodes - 1&
21980                                   ReDim Preserve arr_varRevDefCode(RD_ELEMS, lngE)
21990                                   arr_varRevDefCode(RD_ID_OLD, lngE) = arr_varRevCode(R_ID, lngX)
22000                                   arr_varRevDefCode(RD_ID_NEW, lngE) = CLng(3)
22010                                   arr_varRevDefCode(RD_DSC, lngE) = "OC Other Charges"
                                        'HERE'S WHERE THEY'LL GET A NEW ID!
22020                                 End If
22030                                 arr_varRevCode(R_NID, lngX) = 3&
22040                               Case "OC Other Credits"
22050                                 If arr_varRevCode(R_ID, lngX) <> 4& Then
22060                                   blnTmp24 = True
22070                                   lngRevDefCodes = lngRevDefCodes + 1&
22080                                   lngE = lngRevDefCodes - 1&
22090                                   ReDim Preserve arr_varRevDefCode(RD_ELEMS, lngE)
22100                                   arr_varRevDefCode(RD_ID_OLD, lngE) = arr_varRevCode(R_ID, lngX)
22110                                   arr_varRevDefCode(RD_ID_NEW, lngE) = CLng(4)
22120                                   arr_varRevDefCode(RD_DSC, lngE) = "OC Other Credits"
                                        'HERE'S WHERE THEY'LL GET A NEW ID!
22130                                 End If
                                      'WHY WAS THIS REMARKED OUT?
22140                                 arr_varRevCode(R_NID, lngX) = 4&
22150                               Case "Ordinary Dividend"
22160                                 If arr_varRevCode(R_ID, lngX) <> 5& Then
22170                                   blnTmp24 = True
22180                                   lngRevDefCodes = lngRevDefCodes + 1&
22190                                   lngE = lngRevDefCodes - 1&
22200                                   ReDim Preserve arr_varRevDefCode(RD_ELEMS, lngE)
22210                                   arr_varRevDefCode(RD_ID_OLD, lngE) = arr_varRevCode(R_ID, lngX)
22220                                   arr_varRevDefCode(RD_ID_NEW, lngE) = CLng(5)
22230                                   arr_varRevDefCode(RD_DSC, lngE) = "Ordinary Dividend"
                                        'HERE'S WHERE THEY'LL GET A NEW ID!
22240                                 End If
22250                                 arr_varRevCode(R_NID, lngX) = 5&
22260                               Case "Interest Income"
22270                                 If arr_varRevCode(R_ID, lngX) <> 6& Then
22280                                   blnTmp24 = True
22290                                   lngRevDefCodes = lngRevDefCodes + 1&
22300                                   lngE = lngRevDefCodes - 1&
22310                                   ReDim Preserve arr_varRevDefCode(RD_ELEMS, lngE)
22320                                   arr_varRevDefCode(RD_ID_OLD, lngE) = arr_varRevCode(R_ID, lngX)
22330                                   arr_varRevDefCode(RD_ID_NEW, lngE) = CLng(6)
22340                                   arr_varRevDefCode(RD_DSC, lngE) = "Interest Income"
                                        'HERE'S WHERE THEY'LL GET A NEW ID!
22350                                 End If
22360                                 arr_varRevCode(R_NID, lngX) = 6&
22370                               End Select
22380                             Else
                                    ' ** Add the record to the new table.
                                    ' ** We've already checked for existing Dividend/Interest entries,
                                    ' ** and we're starting with an empty table with only the defaults.
22390                               rstLoc1.AddNew
22400                               rstLoc1![revcode_DESC] = arr_varRevCode(R_DSC, lngX)       '![revcode_DESC]
22410                               rstLoc1![revcode_TYPE] = arr_varRevCode(R_TYP, lngX)       '![revcode_TYPE]
22420                               If arr_varRevCode(R_NSO, lngX) = -1& Then
22430                                 rstLoc1![revcode_SORTORDER] = arr_varRevCode(R_ORD, lngX)  '![revcode_SORTORDER]
22440                               Else
22450                                 rstLoc1![revcode_SORTORDER] = arr_varRevCode(R_NSO, lngX)  '![revcode_SORTORDER]
22460                               End If
22470                               rstLoc1![revcode_ACTIVE] = arr_varRevCode(R_ACT, lngX)     '![revcode_ACTIVE]
22480 On Error Resume Next
22490                               rstLoc1.Update
22500                               If ERR.Number <> 0 Then
22510                                 blnTmp25 = True
22520                                 If ERR.Number = 3022 Then
                                        ' ** Error 3022: The changes you requested to the table were not successful because they
                                        ' **             would create duplicate values in the index, primary key, or relationship.
22530                                   If gblnDev_NoErrHandle = True Then
22540 On Error GoTo 0
22550                                   Else
22560 On Error GoTo ERRH
22570                                   End If
22580                                   lngDupeNum = lngDupeNum + 1&
22590                                   rstLoc1![revcode_DESC] = Trim(Nz(arr_varRevCode(R_DSC, lngX), vbNullString)) & " #DUPE" & CStr(lngDupeNum)
22600                                   rstLoc1![revcode_SORTORDER] = (100& - ![revcode_SORTORDER])  ' ** An attempt at not causing another conflict.
22610                                   rstLoc1.Update
22620                                   lngDupeUnks = lngDupeUnks + 1&
22630                                   ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
22640                                   arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "DUP"
22650                                   arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
22660                                 Else
22670                                   intRetVal = -6
22680                                   blnContinue = False
22690                                   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
22700                                   MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                          "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                          vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
22710                                   rstLoc1.CancelUpdate
22720                                   If gblnDev_NoErrHandle = True Then
22730 On Error GoTo 0
22740                                   Else
22750 On Error GoTo ERRH
22760                                   End If
22770                                 End If
22780                               Else
22790                                 If gblnDev_NoErrHandle = True Then
22800 On Error GoTo 0
22810                                 Else
22820 On Error GoTo ERRH
22830                                 End If
22840                               End If
22850                               If blnContinue = True Then
22860                                 rstLoc1.Bookmark = rstLoc1.LastModified
                                      ' ** Add the ID cross-reference to tblVersion_Key.
22870                                 rstLoc2.AddNew
22880                                 rstLoc2![tbl_id] = lngCurrTblID
22890                                 rstLoc2![tbl_name] = strCurrTblName  'm_REVCODE
22900                                 rstLoc2![fld_id] = lngCurrKeyFldID
22910                                 rstLoc2![fld_name] = strCurrKeyFldName
22920                                 rstLoc2![key_lng_id1] = arr_varRevCode(R_ID, lngX)  ' ** Original revcode_ID.
                                      'rstLoc2![key_txt_id1] =
22930                                 rstLoc2![key_lng_id2] = rstLoc1![revcode_ID]        ' ** New revcode_ID.
                                      'rstLoc2![key_txt_id2] =
22940                                 arr_varRevCode(R_NID, lngX) = rstLoc1![revcode_ID]
22950                                 rstLoc2.Update
                                      'THIS NOW HAS NEW AND OLD REVCODE IDS!
22960                               Else
22970                                 Exit For
22980                               End If
22990                             End If
23000                           End If
23010                         Next
23020                         rstLoc1.Close
23030                         rstLoc2.Close
23040                       End If  ' ** Records present.
23050                       .Close
23060                     End With  ' ** rstLnk.
23070                   End If  ' ** blnFound.

                        ' ** Update any dupes from m_REVCODE.
23080                   For lngX = 0& To (lngRevCodes - 1&)
23090                     If arr_varRevCode(R_DEL, lngX) = True Then
23100                       arr_varRevCode(R_NID, lngX) = arr_varRevCode(R_NID, arr_varRevCode(R_EIM, lngX))
23110                     End If
23120                   Next

                        ' ** Check if the defaults had to be moved.
                        'THIS SHOULD ALSO BE blnTmp24 = TRUE!
23130                   If lngRevDefCodes > 0& Then
23140                     For lngX = 0& To (lngRevDefCodes - 1&)
23150                       For lngY = 0& To (lngRevCodes - 1&)
23160                         If arr_varRevCode(R_ID, lngY) = arr_varRevDefCode(RD_ID_OLD, lngX) Then
23170                           If arr_varRevCode(R_NID, lngY) = 0& Then
23180                             arr_varRevCode(R_NID, lngY) = arr_varRevDefCode(RD_ID_NEW, lngX)
23190                           End If
23200                         End If
23210                       Next
23220                     Next
23230                   End If

                        ' ** Make sure the defaults have all their info.
                        'For lngX = 0& To (lngRevCodes - 1&)
                        '  ' ** The defaults might not be in the original array.
                        '  ' ** If they were, their new sort order should be in arr_varRevCode(R_NSO, lngX).
                        '  ' ** And if they had to be moved, they should've been found in arr_varRevDefCode().
                        '  ' ** Whether they've been moved or not, their arr_varRevCode(R_NID, lngX) should have the right number.
                        '  ' ** If they weren't moved, they won't be found in tblVersion_Key.
                        '  ' ** And, if they weren't there to begin with, they wouldn't be in either arr_varRevCode() or tblVersion_Key, right?
                        'Next

                        ' ** Cross-check it with tblVersion_Key.
23240                   lngRecs = 0&
23250                   Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
23260                   With rstLoc2
23270                     If .BOF = True And .EOF = True Then
                            ' ** I guess this was an empty one!
23280                     Else
23290                       .MoveLast
23300                       lngRecs = .RecordCount
23310                       .MoveFirst
23320                       For lngX = 1& To lngRecs
23330                         If ![tbl_name] = "m_REVCODE" And ![fld_name] = "revcode_ID" Then
23340                           blnFound = False
23350                           For lngY = 0& To (lngRevCodes - 1&)
23360                             If arr_varRevCode(R_ID, lngY) = ![key_lng_id1] And arr_varRevCode(R_NID, lngY) = ![key_lng_id2] Then
23370                               blnFound = True
23380                               arr_varRevCode(R_FND, lngY) = CBool(True)
23390                               Exit For
23400                             End If
23410                           Next
                                'THIS SHOULD JUST CONFIRM THAT CHANGES IN THE REV CODE ARRAY ARE ALL IN THE CHANGED-KEY TABLE!
23420                         End If
23430                         If lngX < lngRecs Then .MoveNext
23440                       Next
23450                     End If
23460                     .Close
23470                   End With  ' ** rstLoc2.
                        'AT THIS POINT, REV CODE ID'S AND SORT ORDER'S MAY HAVE CHANGED,
                        'AND THOSE CHANGES SHOULD BE WELL DOCUMENTED IN THE REV CODE ARRAY,
                        'AS WELL AS THE KEY TABLE, BUT THEIR USE ELSEWHERE IN THE PROGRAM,
                        'IN OTHER TABLES, IS STILL POINTING TO THEIR OLD ID'S!!!

23480                   lngStats = lngStats + 1&
23490                   lngE = lngStats - 1&
23500                   ReDim Preserve arr_varStat(STAT_ELEMS, lngE)
23510                   arr_varStat(STAT_ORD, lngE) = CInt(7)
23520                   arr_varStat(STAT_NAM, lngE) = "Income/Expense Codes: "
23530                   arr_varStat(STAT_CNT, lngE) = CLng(lngRevCodes)
23540                   arr_varStat(STAT_DSC, lngE) = vbNullString

23550                 End If  ' ** blnContinue.

23560                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: TaxCode.
                        ' ******************************

                        ' ** Not a conversion, but a place to load the TaxCode cross-reference.

                        ' ** qryVersion_Convert_30 (tblTemplate_TaxCode_Old, linked to TaxCode,
                        ' ** by taxcode_description (discription[sic])), with converted 'Unknown'.
23570                   Set qdf = dbsLoc.QueryDefs("qryVersion_Convert_31")
23580                   Set rstLoc3 = qdf.OpenRecordset
23590                   With rstLoc3
23600                     .MoveLast
23610                     lngTaxDefCodes = .RecordCount
23620                     .MoveFirst
23630                     arr_varTaxDefCode = .GetRows(lngTaxDefCodes)
                          ' **************************************************
                          ' ** Array: arr_varTaxDefCode()
                          ' **
                          ' **   Field  Element  Name            Constant
                          ' **   =====  =======  ==============  ===========
                          ' **     1       0     taxcode_old     TD_ID_OLD
                          ' **     3       2     taxcode_new     TD_ID_NEW
                          ' **     2       1     discription     TD_DSC
                          ' **     4       3     taxcode_type    TD_TYP
                          ' **
                          ' **************************************************
23640                     .Close
23650                   End With
23660                   Set rstLoc3 = Nothing

23670                 End If  ' ** blnContinue.

23680                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: Users.
                        ' ******************************

                        ' ** Step 9: Users.
23690                   dblPB_ThisStep = 9#
23700                   Version_Status 3, dblPB_ThisStep, "Users"  ' ** Function: Below.

23710                   strCurrTblName = "Users"
23720                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

23730                   blnFound = False: lngRecs = 0&
23740                   For lngX = 0& To (lngOldTbls - 1&)
23750                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
23760                       blnFound = True
23770                       Exit For
23780                     End If
23790                   Next

23800                   If blnFound = True Then
23810                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
23820                     With rstLnk
23830                       If .BOF = True And .EOF = True Then
                              ' ** This has to have records!
23840                       Else
23850                         strCurrKeyFldName = "Username"  ' ** Though [s_GUID] is an AutoNumber, for all practical purposes it can be ignored.
23860                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
23870                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
23880                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Make sure [Default Group] is populated with 'Users'.
                              ' ** Of 16 example TrustDta.mdb's, all have 7 fields.
                              ' ** Current field count is 8 fields.
                              ' ** Table: Users
                              ' **   ![s_Generation]    : These 3 fields are associated with Data Replication, which we don't use.
                              ' **   ![s_GUID]          : [s_Generation] and [s_GUID] both say AutoNumber, but only [s_GUID] gets filled.
                              ' **   ![s_Lineage]       : Also, with Tbl_Fld_Doc(), [s_GUID] shows a DefaultValue not visible in the GUI: GenGUID().
                              ' **   ![Username]
                              ' **   ![Employee Name]
                              ' **   ![Default Group]
                              ' **   ![Primary Group]
                              ' **   ![Secondary Group] : New field since Trust Import.
                              ' ** No tables reference this directly; old, long-gone users are no problem.
23890                         .MoveLast
23900                         lngRecs = .RecordCount
23910                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
23920                         .MoveFirst
23930                         blnTmp22 = False
23940                         For Each fld In .Fields
23950                           With fld
23960                             If .Name = "Secondary Group" Then
23970                               blnTmp22 = True
23980                               Exit For
23990                             End If
24000                           End With
24010                         Next
24020                         For lngX = 1& To lngRecs
24030                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
24040                           Select Case ![Username]
                                Case "Admin", "Creator", "Engine", "Superuser", "TAAdmin", "TAImport"
                                  ' ** All these are now distributed with a new install.
24050                           Case Else
                                  ' ** Add the record to the new table.
24060                             rstLoc1.AddNew
24070                             rstLoc1![Username] = ![Username]
24080                             Select Case IsNull(![Employee Name])
                                  Case True
24090                               strTmp04 = LCase(![Username])
24100                               strTmp04 = UCase$(Left(strTmp04, 1)) & Mid(strTmp04, 2)
24110                               rstLoc1![Employee Name] = strTmp04
24120                               strTmp04 = vbNullString
24130                             Case False
24140                               rstLoc1![Employee Name] = ![Employee Name]
24150                             End Select
24160                             Select Case IsNull(![Default Group])
                                  Case True
24170                               rstLoc1![Default Group] = "Users"
24180                             Case False
24190                               If Trim(![Default Group]) = vbNullString Then
24200                                 rstLoc1![Default Group] = "Users"
24210                               Else
24220                                 rstLoc1![Default Group] = ![Default Group]
24230                               End If
24240                             End Select
24250                             Select Case IsNull(![Primary Group])
                                  Case True
24260                               rstLoc1![Primary Group] = "ViewOnly"  ' ** Instead of '{no entry}'!
24270                             Case False
24280                               If Trim(![Primary Group]) = vbNullString Then
24290                                 rstLoc1![Primary Group] = "ViewOnly"
24300                               Else
24310                                 rstLoc1![Primary Group] = ![Primary Group]
24320                               End If
24330                             End Select
24340                             Select Case blnTmp22
                                  Case True
24350                               Select Case IsNull(![Secondary Group])  ' ** New to v2.1.69.
                                    Case True
24360                                 Select Case ![Username]
                                      Case "Superuser", "TIAdmin", "TIDemo", "TAAdmin", "TADemo", "TAImport"
24370                                   rstLoc1![Secondary Group] = "Admins"
24380                                 Case Else
                                        ' ** Leave it Null.
24390                                 End Select
24400                               Case False
24410                                 rstLoc1![Secondary Group] = ![Secondary Group]
24420                               End Select
24430                             Case False
24440                               Select Case ![Username]
                                    Case "Superuser", "TIAdmin", "TIDemo", "TAAdmin", "TADemo", "TAImport"
24450                                 rstLoc1![Secondary Group] = "Admins"
24460                               Case Else
                                      ' ** Leave it Null.
24470                               End Select
24480                             End Select
24490 On Error Resume Next
24500                             rstLoc1.Update
24510                             If ERR.Number <> 0 Then
24520                               If ERR.Number = 3022 Then
                                      ' ** Error 3022: The changes you requested to the table were not successful because they
                                      ' **             would create duplicate values in the index, primary key, or relationship.
24530                                 If gblnDev_NoErrHandle = True Then
24540 On Error GoTo 0
24550                                 Else
24560 On Error GoTo ERRH
24570                                 End If
24580                                 lngDupeNum = lngDupeNum + 1&
24590                                 rstLoc1![Username] = Trim(Nz(![Username], vbNullString)) & " #DUPE" & CStr(lngDupeNum)
24600                                 rstLoc1.Update
24610                                 lngDupeUnks = lngDupeUnks + 1&
24620                                 ReDim Preserve arr_varDupeUnk(DU_ELEMS, (lngDupeUnks - 1&))
24630                                 arr_varDupeUnk(DU_TYP, (lngDupeUnks - 1&)) = "DUP"
24640                                 arr_varDupeUnk(DU_TBL, (lngDupeUnks - 1&)) = strCurrTblName
24650                               Else
24660                                 intRetVal = -6
24670                                 blnContinue = False
24680                                 lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
24690                                 MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
                                        "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
                                        vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
24700                                 rstLoc1.CancelUpdate
24710                                 If gblnDev_NoErrHandle = True Then
24720 On Error GoTo 0
24730                                 Else
24740 On Error GoTo ERRH
24750                                 End If
24760                               End If
24770                             Else
24780                               If gblnDev_NoErrHandle = True Then
24790 On Error GoTo 0
24800                               Else
24810 On Error GoTo ERRH
24820                               End If
24830                             End If
24840                             If blnContinue = True Then
                                    ' ** No key field to keep track of.
24850                             Else
24860                               Exit For
24870                             End If
24880                           End Select
24890                           If lngX < lngRecs Then .MoveNext
24900                         Next
24910                         rstLoc1.Close
24920                         rstLoc2.Close
24930                       End If  ' ** Records present.
24940                       .Close
24950                     End With  ' ** rstLnk.
24960                   End If  ' ** blnFound.

24970                 End If  ' ** blnContinue.

24980                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: _~xusr.
                        ' ******************************

                        ' ** Step 10: _~xusr.
24990                   dblPB_ThisStep = 10#
25000                   Version_Status 3, dblPB_ThisStep, "Zeta"  ' ** Function: Below.

25010                   strCurrTblName = "_~xusr"
                        ' ** An error here HAS to be prevented by running Doc's before release!
25020                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

25030                   blnFound = False: lngRecs = 0&
25040                   For lngX = 0& To (lngOldTbls - 1&)
25050                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
25060                       blnFound = True  ' ** Yes, it's in the database to be converted.
25070                       Exit For
25080                     End If
25090                   Next

25100                   If blnFound = True Then
25110                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
25120                     With rstLnk
25130                       If .BOF = True And .EOF = True Then
                              ' ** If it has no records, Security_SyncChk() will take care of it.
25140                       Else
25150                         strCurrKeyFldName = "xusr_id"
25160                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
25170                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
25180                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Table: _~xusr
                              ' **   ![xusr_id]
                              ' **   ![s_GUID]           : Used in order to match the Users table.
                              ' **   ![xusr_extant]      : Current password, encoded.
                              ' **   ![xusr_antecedent]  : Previous password, encoded.
                              ' **   ![xusr_origin]      : Date of last password change, encoded.
                              ' **   ![xusr_user]
                              ' **   ![xusr_datecreated]
                              ' **   ![xusr_datemodified]
25190                         .MoveLast
25200                         lngRecs = .RecordCount
25210                         Version_Status 3, dblPB_ThisStep, strCurrTblName, -1&, lngRecs  ' ** Function: Below.
25220                         rstLoc1.MoveLast
25230                         lngTmp15 = rstLoc1.RecordCount
25240                         rstLoc1.MoveFirst
25250                         .MoveFirst
25260                         For lngX = 1& To lngRecs
25270                           Version_Status 4, dblPB_ThisStep, strCurrTblName, lngX, lngRecs  ' ** Function: Below.
                                ' ** Compare s_GUID's, and only copy those not already here.
                                ' ** The GUID field is designated as a number, with a Type (FieldSize) of 'Replication ID'.
                                ' ** It's really an array of Bytes, but after much searching and experimentation (with lots
                                ' ** of Type Mismatch errors), discovered (via a forum posting somewhere) that all it accepts
                                ' ** is the plain GUID string.
                                ' ** The designation as 'Replication ID' tells Jet how to handle that GUID string.
25280                           blnFound = False: varTmp00 = Empty
25290                           varTmp00 = FilterGUIDString(StringFromGUID(![s_GUID]))  ' ** Module Function: modCodeUtilities.
25300                           For lngY = 1& To lngTmp15
25310                             If FilterGUIDString(StringFromGUID(rstLoc1![s_GUID])) = varTmp00 Then  ' ** Module Function: modCodeUtilities.
                                    ' ** A record distributed with TA, and whose s_GUID and password should remain constant.
25320                               blnFound = True
25330                               Exit For
25340                             End If
25350                             If lngY < lngTmp15 Then rstLoc1.MoveNext
25360                           Next
25370                           If blnFound = False Then
                                  ' ** Add the record to the new table.
25380                             rstLoc1.AddNew
25390                             rstLoc1![s_GUID] = varTmp00                         ' ** Yes, it takes a plain string!
25400                             rstLoc1![xusr_extant] = ![xusr_extant]              ' ** Current password.
25410                             rstLoc1![xusr_antecedent] = ![xusr_antecedent]      ' ** Last password.
25420                             rstLoc1![xusr_origin] = ![xusr_origin]              ' ** Date password last changed (or was created).
25430                             rstLoc1![xusr_user] = ![xusr_user]                  ' ** Computer log-in of the user who created this record.
25440                             rstLoc1![xusr_datecreated] = ![xusr_datecreated]    ' ** Date password record first created.
25450                             rstLoc1![xusr_datemodified] = ![xusr_datemodified]  ' ** Date password record last modified.
25460                             rstLoc1.Update
                                  ' ** If there's an error, I'll have to wait until it happens to figure it out!
25470                           End If  ' ** blnFound.
                                ' ** Reset the local _~xusr.
25480                           rstLoc1.MoveFirst
25490                           If blnContinue = True Then
                                  ' ** No need to keep track of a key field.
25500                           End If
25510                           If lngX < lngRecs Then .MoveNext
25520                         Next  ' ** rstLnk.
25530                       End If  ' ** Records present.
25540                       .Close
25550                     End With  ' ** rstLnk.
25560                   End If  ' ** blnFound.

25570                 End If  ' ** blnContinue.

25580                 If blnContinue = True Then
                        ' ** dbsLoc is still open.

                        ' ******************************
                        ' ** Table: tblSecurity_License.
                        ' ******************************

                        ' ** Step 11: tblSecurity_License.
25590                   dblPB_ThisStep = 11#
25600                   Version_Status 3, dblPB_ThisStep, "License"  ' ** Function: Below.

25610                   strCurrTblName = "tblSecurity_License"
                        ' ** An error here HAS to be prevented by running Doc's before release!
25620                   lngCurrTblID = DLookup("[tbl_ID]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngTrustDbsID) & " And " & _
                          "[tbl_name] = '" & strCurrTblName & "'")

25630                   blnFound = False: lngRecs = 0&
25640                   For lngX = 0& To (lngOldTbls - 1&)
25650                     If arr_varOldTbl(T_TNAM, lngX) = strCurrTblName Then
25660                       blnFound = True  ' ** Yes, it's in the database to be converted.
25670                       Exit For
25680                     End If
25690                   Next

25700                   If blnFound = True Then
25710                     Set rstLnk = .OpenRecordset(strCurrTblName, dbOpenDynaset, dbReadOnly)
25720                     With rstLnk
25730                       If .BOF = True And .EOF = True Then
                              ' ** Shouldn't be the case, but defaults and startup procedures will take/have taken care of this.
25740                       Else
25750                         strCurrKeyFldName = "seclic_id"
                              'Debug.Print "'dbs_id = " & CStr(lngTrustDtaDbsID) & "  tbl_id = " & CStr(lngCurrTblID) & "  fld_name = " & strCurrKeyFldName
                              'dbs_id = 2 (TrustDta.mdb)  tbl_id = 539 (_~xusr)  fld_name = seclic_id
25760                         lngCurrKeyFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", _
                                "[dbs_id] = " & CStr(lngTrustDtaDbsID) & " And " & _
                                "[tbl_id] = " & CStr(lngCurrTblID) & " And [fld_name] = '" & strCurrKeyFldName & "'")
25770                         Set rstLoc1 = dbsLoc.OpenRecordset(strCurrTblName, dbOpenDynaset, dbConsistent)
25780                         Set rstLoc2 = dbsLoc.OpenRecordset(strKeyTbl, dbOpenDynaset, dbConsistent)
                              ' ** Table: tblSecurity_License
                              ' **   ![seclic_id]
                              ' **   ![seclic_licensedto]
                              ' **   ![seclic_clientpath_ta]
                              ' **   ![seclic_datapath_ta]
                              ' **   ![seclic_auxiliarypath]
                              ' **   ![seclic_cycle]
                              ' **   ![seclic_cycle_screen]
                              ' **   ![seclic_cycle_message]
                              ' **   ![seclic_user]
                              ' **   ![Username]
                              ' **   ![seclic_datemodified]
25790                         .MoveFirst  ' ** Only 1 record.
25800                         rstLoc1.MoveFirst
25810                         rstLoc1.Edit
                              ' ** ![seclic_licensedto] : Taken care of elsewhere.
                              ' ** ![seclic_clientpath_ta] : Taken care of elsewhere.
                              ' ** ![seclic_datapath_ta] : Taken care of elsewhere.
                              ' ** ![seclic_auxiliarypath] : Taken care of elsewhere.
25820                         rstLoc1![seclic_cycle] = ![seclic_cycle]  ' ** Encoded string.
25830                         rstLoc1![seclic_cycle_screen] = ![seclic_cycle_screen]  ' ** Encoded string.
25840                         rstLoc1![seclic_cycle_message] = ![seclic_cycle_message]  ' ** Encoded string.
25850                         rstLoc1![seclic_user] = ![seclic_user]
25860                         rstLoc1![Username] = ![Username]
25870                         rstLoc1![seclic_datemodified] = ![seclic_datemodified]
25880                         rstLoc1.Update
25890                         If blnContinue = True Then
                                ' ** Since there's only 1 record, no need to keep track of a key field.
25900                         End If
25910                       End If  ' ** Records present.
25920                       .Close
25930                     End With  ' ** rstLnk.
25940                   End If  ' ** blnFound.

25950                 End If  ' ** blnContinue.

25960               End With  ' ** TrustDta.mdb: dbsLnk.

25970             End If  ' ** dbsLnk opens.

25980           End With  ' ** wrkLnk.

25990         End If  ' ** Workspace opens: blnContinue

26000       End If  ' ** blnConvert_TrustDta.

26010       If blnContinue = False Then
26020         dbsLoc.Close
26030         wrkLoc.Close
26040       End If

26050     Else
            ' ** Conversion already done.
26060       blnContinue = False
26070     End If

26080   Else
          ' ** Not a conversion.
26090     blnContinue = False
26100   End If

26110   If blnContinue = False Then
26120     DoCmd.Hourglass False
26130   End If

EXITP:
26140   Version_Upgrade_03 = intRetVal
26150   Exit Function

ERRH:
26160   intRetVal = -9
26170   DoCmd.Hourglass False
26180   lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
26190   Select Case ERR.Number
        Case Else
26200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26210   End Select
26220   Resume EXITP

End Function

Public Function Version_Status(intInstance As Integer, Optional varStep As Variant, Optional varStatus As Variant, Optional varX As Variant, Optional varRecs As Variant) As Boolean
' ** This handles all progress reports to the conversion form, including the progress bar.
' ** Parameters:
' **   intInstance :  Type of status update:
' **               :    1 : Open frmVersion_Main as acDialog; no Optional parameters required.
' **               :    2 : Open frmVersion_Main as acWindowNormal; no Optional parameters required.
' **               :    3 : Update frmVersion_Main with current step and status; varStep and varStatus required.
' **               :    4 : Update on-screen counter; varX, varRecs required.
' **               :    5 : Finish, with final conversion status; no Optional parameters required.
' **   varStep     :  Current step of the conversion.
' **   varStatus   :  Message to appear with progress bar.
' **   varX        :  Counter iteration.
' **   varRecs     :  Total records for counter iteration.

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "Version_Status"

        Dim frm As Access.Form, rst As DAO.Recordset
        Dim fstx As Scripting.TextStream
        Dim strPath As String
        Dim lngOff1 As Long, lngOff2 As Long
        Dim intPos01 As Integer
        Dim dblZ As Double
        Dim lngStat1Orig_Top As Long, lngStat2Orig_Top As Long
        Dim blnRetVal As Boolean

        Static lngTpp As Long

26310   If gblnDev_NoErrHandle = True Then
26320 On Error GoTo 0
26330   End If

26340   blnRetVal = True

26350   If IsLoaded(FRM_CNV_STATUS, acForm) = True Then  ' ** Module Function: modFileUtilities.
26360     Set frm = Forms(FRM_CNV_STATUS)
26370     lngStat1Orig_Top = frm.Stat1Orig_line.Top  '180&  ' ** Where I first put them, and wrote
26380     lngStat2Orig_Top = frm.Stat2Orig_line.Top  '480&  ' ** code to reflect that spot.
26390   End If

26400   If lngTpp = 0& Then
          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
26410     lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
26420   End If

26430   If intInstance = 3 Then
26440     If IsMissing(varStep) = True Then
            'Stop
26450     Else
26460       If varStep = 1 Then
              'Stop
26470       End If
26480     End If
26490   End If

26500   Select Case intInstance
        Case 1
26510     strTmp04 = arr_varOldFile(F_PTHFIL, lngDtaElem)
26520     strTmp05 = arr_varOldFile(F_PTHFIL, lngArchElem)
26530     strTmp07 = Format(arr_varOldFile(F_APPDATE, lngDtaElem), "mm/dd/yyyy")
26540     Version_DataXFer_PathVars strTmp04, strTmp05, strOldVersion, strTmp07  ' ** Module Procedure: modVersionConvertFuncs2.
26550     blnRetVal = Version_DataXFer("Set", "PathFile")  ' ** Module Function: modVersionConvertFuncs2.
          'Version_DataXFer() IS EXPECTING strOldVersion, WHICH WAS SET IN Version_Upgrade_01(), ABOVE!
26560     If blnRetVal = True Then
26570       DoCmd.OpenForm FRM_CNV_STATUS, , , , , acDialog, THIS_NAME
26580     End If

26590   Case 2
26600     strTmp04 = arr_varOldFile(F_PTHFIL, lngDtaElem)
26610     strTmp05 = arr_varOldFile(F_PTHFIL, lngArchElem)
26620     strTmp07 = Format(arr_varOldFile(F_APPDATE, lngDtaElem), "mm/dd/yyyy")
26630     Version_DataXFer_PathVars strTmp04, strTmp05, vbNullString, strTmp07  ' ** Module Procedure: modVersionConvertFuncs2.
26640     blnRetVal = Version_DataXFer("Set", "PathFile")  ' ** Module Function: modVersionConvertFuncs2.
26650     If blnRetVal = True Then
26660       DoCmd.OpenForm FRM_CNV_STATUS, , , , , acWindowNormal, THIS_NAME
26670     End If

26680   Case 3
26690     If IsMissing(varStep) = False And IsMissing(varStatus) = False Then
26700       Set frm = Forms(FRM_CNV_STATUS)
26710       With frm
26720         Select Case varStep
              Case 1

                ' ** Percentage labels must be left-aligned, with centering done by spaces.
                ' ** This is so that the front label's width (white Forecolor) can expand with the
                ' ** blue bar, revealing the white letters only as the bar approaches half-way.
26730           strSp = Space(74)

                ' ** Put the tblVersion_Conversion ID on the form.
26740           .vercnv_id = lngVerCnvID

                ' ** Change from Intro to Progress.
26750           .FocusHolder.SetFocus
26760           .IntroMsg_lbl_img.Visible = False
26770           .MsgRed_lbl_img.Visible = False
26780           .PathFile_TrustData.Enabled = False
26790           .PathFile_TrustArchive.Enabled = False
26800           .cmdConvert.Enabled = False
26810           .cmdCancel.Enabled = False
26820           .Status1_lbl.Visible = True
26830           .Status2_lbl.Visible = True

                ' ** Move status stuff to a better looking arrangement.
26840           Version_Stat2_Etc 1, lngOff1, lngOff2, lngStat1Orig_Top, lngStat2Orig_Top, arr_dblPB_ThisIncr, dblPB_Width, frm  ' ** Module Procedure: modVersionConvertFuncs2.

                ' ******************************************************************
                ' ** Version_Upgrade_02()
                ' **   Step 0:  Determine whether conversion is necessary.
                ' **
                ' ** Version_Upgrade_03()
                ' **   Step 1:  Beginning. (Connect, get list of tables, get list of accounts.)
                ' **   Step 2:  CompanyInformation
                ' **   Step 3:  adminofficer
                ' **   Step 4:  Location
                ' **   Step 5:  RecurringItems
                ' **   Step 6:  Schedule, ScheduleDetail
                ' **   Step 7:  DDate, PostingDate, Statement Date
                ' **   Step 8:  m_REVCODE
                ' **   Step 9:  Users
                ' **   Step 10: _~xusr
                ' **   Step 11: tblSecurity_License
                ' **
                ' ** Version_Upgrade_04()
                ' **   Step 12: account
                ' **   Step 13: Balance
                ' **   Step 14: masterasset
                ' **   Step 15: ActiveAssets
                ' **   Step 16: ledger
                ' **   Step 17: journal
                ' **   Step 18: LedgerHidden
                ' **
                ' ** Version_Upgrade_06()
                ' **   Step 19: LedgerArchive
                ' **   Step 20: tblPricing_MasterAsset_History (tblAssetPricing)
                ' **   Step 21: tblJournal_Memo
                ' **
                ' ** Version_Upgrade_07()
                ' **   Step 22: tblCurrency
                ' **   Step 23: tblCurrency_History
                ' **   Step 24: tblLedgerHidden
                ' **
                ' ** Version_Upgrade_01()
                ' **   Step 25: Rename files in \Convert_New\
                ' ******************************************************************

                ' ** Initialize the progress bar.
26850           dblPB_Steps = 34#
26860           ReDim arr_dblPB_ThisIncr(dblPB_Steps)  ' ** Since arrays are zero-based, this one will only use 1-25, and not 0.
26870           dblPB_Width = .ProgBar_box.Width

                ' ** Weight the steps.
26880           Version_Stat2_Etc 2, lngOff1, lngOff2, lngStat1Orig_Top, lngStat2Orig_Top, arr_dblPB_ThisIncr, dblPB_Width, frm  ' ** Module Procedure: modVersionConvertFuncs2.

                ' ***************************************************************
26890           dblPB_ThisWidth = 0#  ' ** dblPB_ThisStep set above.
26900           For dblZ = 1# To (dblPB_ThisStep - 1#)
                  ' ** Assemble the weighted widths up to, but not including, this width.
26910             dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
26920           Next
26930           dblPB_StepSubs = 0#  ' ** No subs in this step.
26940           dblPB_ThisIncrSub = 0#
26950           dblPB_ThisStepSub = 0#
26960           ProgBar_Width_Conv frm, dblPB_ThisWidth, 2  ' ** Module Procedure: modVersionConvertFuncs2.
                '.ProgBar_bar.Width = dblPB_ThisWidth
26970           .ProgBar_lbl2.Width = dblPB_ThisWidth + (2& * lngTpp)  ' ** Because of the label's right margin.
26980           .Status1_lbl.Caption = "Beginning Conversion"
26990           .Status2_lbl.Caption = "Establishing link to old file"
27000           strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
27010           .ProgBar_lbl1.Caption = strSp & strPB_ThisPct
27020           .ProgBar_lbl2.Caption = strSp & strPB_ThisPct
                ' ***************************************************************

27030           .ProgBar_box.Visible = True
27040           .ProgBar_box2.Visible = True
27050           ProgBar_Width_Conv frm, True, 1  ' ** Module Procedure: modVersionConvertFuncs2.
                '.ProgBar_bar.Visible = True
27060           .ProgBar_lbl1.Visible = True
27070           .ProgBar_lbl2.Visible = True

27080           .Status1Cnt_lbl.Top = .Status2_lbl.Top + (2& * lngTpp)
27090           .Status1Of_lbl.Top = .Status2_lbl.Top + (2& * lngTpp)
27100           .Status1Tot_lbl.Top = .Status2_lbl.Top + (2& * lngTpp)

27110           DoEvents

27120         Case Else

27130           If varStatus <> "End" Then

                  ' ***************************************************************
27140             dblPB_ThisWidth = 0#  ' ** dblPB_ThisStep set above.
27150             For dblZ = 1# To (dblPB_ThisStep - 1#)
                    ' ** Assemble the weighted widths up to, but not including, this width.
27160               dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
27170             Next
                  ' ** First we come through with the step and name, but don't know lngRecs;
                  ' ** varRecs is missing.
27180             If IsMissing(varRecs) = False Then
                    ' ** We come back again with lngRecs, but before starting the loop;
                    ' ** varX = -1, varRecs = lngRecs.
27190               dblPB_StepSubs = varRecs
27200               dblPB_ThisIncrSub = (arr_dblPB_ThisIncr(dblPB_ThisStep) / varRecs)  ' ** The total width for just this step, divided by the sub steps.
27210               dblPB_ThisStepSub = 0#
27220               .Status1Tot_lbl.Caption = CStr(varRecs)
27230             Else
27240               ProgBar_Width_Conv frm, dblPB_ThisWidth, 2  ' ** Module Procedure: modVersionConvertFuncs2.
                    '.ProgBar_bar.Width = dblPB_ThisWidth
27250               .ProgBar_lbl2.Width = dblPB_ThisWidth + (2& * lngTpp)  ' ** Because of the label's right margin.
27260               .Status1_lbl.Caption = "Converting"
27270               If varStatus = "tblPricing_MasterAsset_History" Then
27280                 .Status2_lbl.Caption = "MasterAsset Pricing History"
27290               Else
27300                 .Status2_lbl.Caption = varStatus
27310               End If
27320               If dblPB_Width = 0# Then
27330                 strPB_ThisPct = Format(dblPB_Width, "##0%")
27340               Else
27350                 strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
27360               End If
27370               .ProgBar_lbl1.Caption = strSp & strPB_ThisPct
27380               .ProgBar_lbl2.Caption = strSp & strPB_ThisPct
                    ' ** This sets up the counter.
27390               .Status1Cnt_lbl.Visible = True
27400               .Status1Of_lbl.Visible = True
27410               .Status1Tot_lbl.Visible = True
27420               .Status1Cnt_lbl.Caption = CStr(0)
27430               .Status1Tot_lbl.Caption = CStr(0)
27440             End If
                  ' ***************************************************************

27450           Else

                  ' ***************************************************************
27460             dblPB_ThisWidth = dblPB_Width
27470             strPB_ThisPct = Format(1#, "##0%")
27480             .Status1_lbl.Caption = "Finished"
27490             .Status1_lbl.FontBold = True
27500             .Status2_lbl.Caption = vbNullString
27510             .ProgBar_lbl1.Caption = strSp & strPB_ThisPct
27520             .ProgBar_lbl2.Caption = strSp & strPB_ThisPct
27530             ProgBar_Width_Conv frm, dblPB_ThisWidth, 2  ' ** Module Procedure: modVersionConvertFuncs2.
                  '.ProgBar_bar.Width = dblPB_ThisWidth
27540             .ProgBar_lbl2.Width = dblPB_ThisWidth
                  ' ***************************************************************

27550             If IsNull(.PathFile_TrustData) = False And IsNull(.PathFile_TrustArchive) = False Then
27560               If Trim(.PathFile_TrustData) <> vbNullString And Trim(.PathFile_TrustArchive) <> vbNullString Then
27570                 .PathFile_TrustData = Left(.PathFile_TrustData, (Len(.PathFile_TrustData) - 3)) & "BAK"
27580                 .PathFile_TrustArchive = Left(.PathFile_TrustArchive, (Len(.PathFile_TrustArchive) - 3)) & "BAK"
27590               Else
27600                 DoCmd.Hourglass False
27610                 Beep
27620                 MsgBox "Paths not present!", vbCritical + vbOKOnly, "Path Not Found"
27630               End If
27640             Else
27650               DoCmd.Hourglass False
27660               Beep
27670               MsgBox "Paths not present!", vbCritical + vbOKOnly, "Path Not Found"
27680             End If

27690             .Status1Cnt_lbl.Visible = False
27700             .Status1Of_lbl.Visible = False
27710             .Status1Tot_lbl.Visible = False

                  ' ** Wait till summary is on the screen.
                  '.cmdCancel.Enabled = True
27720             .cmdCancel.Caption = "&Continue"
27730             DoEvents
                  '.cmdCancel.SetFocus
27740             .FocusHolder.SetFocus

27750           End If

27760           DoEvents
27770         End Select
27780       End With  ' ** frm.

            ' ** Update tblVersion_Conversion, by specified [vercid], [verstp].
27790       Set qdf = dbsLoc.QueryDefs("qryVersion_Convert_04")
27800       With qdf.Parameters
27810         ![vercid] = lngVerCnvID
27820         ![verstp] = dblPB_ThisStep
27830       End With  ' ** Parameters.
27840       qdf.Execute

27850     End If
27860   Case 4
          ' ** Counter

27870     Set frm = Forms(FRM_CNV_STATUS)
27880     With frm
            ' ********************************************
27890       dblPB_ThisStepSub = varX
27900       dblPB_ThisWidth = 0#
27910       For dblZ = 1# To (dblPB_ThisStep - 1#)
              ' ** Assemble the weighted widths up to, but not including, this width.
27920         dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
27930       Next
27940       dblPB_ThisWidth = dblPB_ThisWidth + (dblPB_ThisStepSub * dblPB_ThisIncrSub)
27950       ProgBar_Width_Conv frm, dblPB_ThisWidth, 2  ' ** Module Procedure: modVersionConvertFuncs2.
            '.ProgBar_bar.Width = dblPB_ThisWidth
27960       .ProgBar_lbl2.Width = dblPB_ThisWidth + (2& * lngTpp)
27970       strPB_ThisPct = Format((dblPB_ThisWidth / dblPB_Width), "##0%")
27980       .ProgBar_lbl1.Caption = strSp & strPB_ThisPct
27990       .ProgBar_lbl2.Caption = strSp & strPB_ThisPct
28000       .Status1Cnt_lbl.Caption = CStr(dblPB_ThisStepSub)
28010       .Status1Tot_lbl.Caption = CStr(varRecs)
            ' ********************************************
28020     End With  ' ** frm

28030     DoEvents

28040   Case 5
          ' ** Finish, with final conversion status.

28050     Set frm = Forms(FRM_CNV_STATUS)
28060     With frm

28070       DoCmd.SelectObject acForm, FRM_CNV_STATUS, False

28080       If gintConvertResponse = 0 Then

              ' ** Binary Sort arr_varStat() array.
28090         For lngX = UBound(arr_varStat, 2) To 1& Step -1
28100           For lngY = 0 To (lngX - 1)
28110             If arr_varStat(STAT_ORD, lngY) > arr_varStat(STAT_ORD, (lngY + 1)) Then
28120               varTmp00 = Empty
28130               For lngZ = 0& To STAT_ELEMS
28140                 varTmp00 = arr_varStat(lngZ, lngY)
28150                 arr_varStat(lngZ, lngY) = arr_varStat(lngZ, (lngY + 1))
28160                 arr_varStat(lngZ, (lngY + 1)) = varTmp00
28170                 varTmp00 = Empty
28180               Next
28190             End If
28200           Next
28210         Next

28220         lngOff1 = (9& * lngTpp)  '135
28230         If (.TAVer_Old_lbl.Top - lngOff1) < 0& Then lngOff1 = .TAVer_Old_lbl.Top
28240         .TAVer_Old_lbl.Top = .TAVer_Old_lbl.Top - lngOff1
28250         .TAVer_Old.Top = .TAVer_Old.Top - lngOff1
28260         .TAVer_Old_RelDate.Visible = False
28270         .TAVer_New_lbl.Top = .TAVer_New_lbl.Top - lngOff1
28280         .TAVer_New.Top = .TAVer_New.Top - lngOff1
28290         .TAVer_New_RelDate.Visible = False
28300         .TAVer_Arrow.Top = .TAVer_Arrow.Top - lngOff1
28310         .TAVer_box.Visible = False

              ' ** PathFile_TrustData and PathFile_TrustArchive heights: 1-line = 285&, 2-line = 510&.
28320         If .PathFile_TrustData.Height <> 285& Then .PathFile_TrustData.Height = 285&
28330         If .PathFile_TrustArchive.Height <> 285& Then .PathFile_TrustArchive.Height = 285&

              ' ** Arrange the paths.
28340         Version_Stat2_Etc 3, lngOff1, lngOff2, lngStat1Orig_Top, lngStat2Orig_Top, arr_dblPB_ThisIncr, dblPB_Width, frm  ' ** Module Procedure: modVersionConvertFuncs2.

              ' *********************************************
              ' ** Array: arr_varStat()
              ' **
              ' **   Element  Name             Constant
              ' **   =======  ===============  ============
              ' **      0     Display Order    STAT_ORD
              ' **      1     Table Name       STAT_NAM
              ' **      2     Record Count     STAT_CNT
              ' **      3     Description      STAT_DSC
              ' **
              ' *********************************************
              ' ** 1. Company
              ' ** 2. Accounts
              ' ** 3. Ledger Entries
              ' ** 4. Archived Entries
              ' ** 5. Tax Lots
              ' ** 6. Master Assets
              ' ** 7. RevCodes
              ' ** 8. 1st Statement Date
              ' ** 9. Latest Statement Date.
28350         strTmp04 = vbNullString
28360         For lngX = 0& To (lngStats - 1&)
28370           If lngX > 0& Then strTmp04 = strTmp04 & vbCrLf
28380           Select Case arr_varStat(STAT_ORD, lngX)
                Case 1
28390             strTmp04 = strTmp04 & arr_varStat(STAT_NAM, lngX) & arr_varStat(STAT_DSC, lngX)
28400           Case 8
28410             strTmp04 = strTmp04 & arr_varStat(STAT_NAM, lngX) & vbCrLf
28420             If arr_varStat(STAT_CNT, lngX) > 0 Then
28430               strTmp04 = strTmp04 & "  " & arr_varStat(STAT_DSC, lngX) & Format(CDate(arr_varStat(STAT_CNT, lngX)), "mm/dd/yyyy")
28440             Else
28450               strTmp04 = strTmp04 & "  " & arr_varStat(STAT_DSC, lngX)
28460             End If
28470           Case 9
28480             If arr_varStat(STAT_CNT, lngX) > 0 Then
28490               strTmp04 = strTmp04 & "  " & arr_varStat(STAT_DSC, lngX) & "  " & Format(CDate(arr_varStat(STAT_CNT, lngX)), "mm/dd/yyyy")
28500             Else
28510               strTmp04 = strTmp04 & "  " & arr_varStat(STAT_DSC, lngX)
28520             End If
28530           Case Else
28540             strTmp04 = strTmp04 & arr_varStat(STAT_NAM, lngX) & CStr(arr_varStat(STAT_CNT, lngX))
28550           End Select
28560         Next

28570         strTmp05 = vbNullString
28580         strTmp05 = strTmp05 & "Trust Accountant Conversion" & vbCrLf
28590         strTmp05 = strTmp05 & "Version " & .TAVer_Old & " to " & .TAVer_New & vbCrLf
28600         strTmp05 = strTmp05 & Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") & vbCrLf
28610         strTmp05 = strTmp05 & String(40, "=") & vbCrLf

28620         strSummaryMsg = strTmp05 & strTmp04

              ' ** Empty Ledger/LedgerArchive records.
28630         If lngLedgerEmptyDels > 0& Or lngLedgerArchEmptyDels = 0& Then
28640           strTmp04 = vbNullString
28650           If lngLedgerEmptyDels > 0& Then
28660             strTmp04 = vbCrLf & vbCrLf & CStr(lngLedgerEmptyDels) & " completely empty Ledger records were not converted."
28670           End If
28680           If lngLedgerArchEmptyDels > 0& Then
28690             If strTmp04 = vbNullString Then strTmp04 = vbCrLf & vbCrLf
28700             strTmp04 = strTmp04 & vbCrLf & CStr(lngLedgerArchEmptyDels) & " completely empty LedgerArchive records were not converted."
28710           End If
28720           strSummaryMsg = strSummaryMsg & strTmp04
28730         End If

              ' ** Add summary of #DUPE and UNKNOWN.
              ' *********************************************
              ' ** Array: arr_varDupeUnk()
              ' **
              ' **   Element  Name               Constant
              ' **   =======  =================  ==========
              ' **      0     Dupe or Unknown    DU_TYP
              ' **      1     tbl_name           DU_TBL
              ' **
              ' *********************************************

28740         If lngDupeUnks > 0& Then
28750           lngTmp13 = 0&: lngTmp14 = 0&
28760           strTmp05 = vbNullString: strTmp06 = vbNullString
28770           For lngX = 0& To (lngDupeUnks - 1&)
28780             Select Case arr_varDupeUnk(DU_TYP, lngX)
                  Case "DUP"
28790               lngTmp13 = lngTmp13 + 1&
28800             Case "UNK"
28810               lngTmp14 = lngTmp14 + 1&
28820             End Select
28830           Next

                ' ** A little diversion...
28840           For lngX = 0& To (lngDupeUnks - 1&)
28850             If arr_varDupeUnk(DU_TBL, lngX) = "Location" And arr_varDupeUnk(DU_TYP, lngX) = "DUP" Then
28860               Set rstLoc3 = CurrentDb.OpenRecordset("Location", dbOpenDynaset, dbConsistent)
28870               With rstLoc3
28880                 .MoveLast
28890                 lngRecs = .RecordCount
28900                 .MoveFirst
28910                 strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString: lngTmp20 = 0&: lngTmp21 = 0&
28920                 For lngY = 1& To lngRecs
28930                   If InStr(![Location_Name], "#DUPE") > 0 Then
28940                     lngTmp20 = ![Location_ID]
28950                     strTmp08 = ![Location_Name]  ' ** With '#DUPE'.
28960                     strTmp09 = Trim(Left(strTmp08, (InStr(strTmp08, "#DUPE") - 1)))  ' ** Without '#DUPE'.
28970                     If IsNull(![Location_Address1]) = True Then
28980                       strTmp10 = "#NULL"
28990                     Else
29000                       strTmp10 = ![Location_Address1]
29010                     End If
29020                     Exit For
29030                   End If
29040                   If lngX < lngRecs Then .MoveNext
29050                 Next
29060                 If strTmp08 <> vbNullString Then
29070                   .MoveFirst
29080                   For lngY = 1& To lngRecs
29090                     If ![Location_ID] <> lngTmp20 Then
                            ' ** Skip the marked one.
29100                       If Left(![Location_Name], Len(strTmp09)) = strTmp09 And InStr(![Location_Name], "#DUPE") = 0 Then
                              ' ** This must be the other one. (Or at least the only one that isn't a dupe!)
29110                         lngTmp21 = ![Location_ID]
29120                         If IsNull(![Location_Address1]) = True Then
29130                           If strTmp10 = "#NULL" Then
                                  ' ** Both Null, and I'm not going to check any of the other fields!
29140                             strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
29150                             Exit For
29160                           Else
                                  ' ** Darn! The marked one is the good one.
29170                             strTmp09 = ![Location_Name]  ' ** This now has a full, good name.
29180                             Exit For
29190                           End If
29200                         Else
29210                           If strTmp10 = "#NULL" Then
                                  ' ** Fabulous! The unmarked one is the good one.
29220                             strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
29230                             Exit For
29240                           Else
                                  ' ** So, neither is Null! Forget the whole idea.
29250                             strTmp08 = vbNullString: strTmp09 = vbNullString: strTmp10 = vbNullString
29260                             Exit For
29270                           End If
29280                         End If
29290                       End If
29300                     End If
29310                     If lngX < lngRecs Then .MoveNext
29320                   Next
29330                 End If
29340                 If strTmp08 <> vbNullString Then
29350                   .FindFirst "[Location_ID] = " & CStr(lngTmp20)  ' ** The marked one.
29360                   If .NoMatch = False Then
29370                     .Edit
29380                     ![Location_Name] = Left("XX" & ![Location_Name], 50)  ' ** Alter the Name temporarily.
29390                     .Update
29400                     .MoveFirst
29410                     .FindFirst "[Location_ID] = " & CStr(lngTmp21)  ' ** The bad, unmarked one.
29420                     If .NoMatch = False Then
29430                       .Edit
29440                       ![Location_Name] = strTmp08  ' ** Put in the marked name.
29450                       .Update
29460                       .MoveFirst
29470                       .FindFirst "[Location_ID] = " & CStr(lngTmp20)  ' ** Back to the marked one.
29480                       If .NoMatch = False Then
29490                         .Edit
29500                         ![Location_Name] = strTmp09  ' ** The full, good name.
29510                         .Update
29520                       End If
29530                     End If
29540                   End If
29550                 End If
29560                 .Close
29570               End With
29580             End If
29590           Next
                ' ** DO I WANT TO NOW CHANGE THE RECORDS REFERRING TO THEM?

29600           For lngX = 0& To (lngDupeUnks - 1&)
29610             Select Case arr_varDupeUnk(DU_TBL, lngX)
                  Case "account"
29620               arr_varDupeUnk(DU_TBL, lngX) = "Accounts        : View / Add Account"
29630             Case "adminofficer"
29640               arr_varDupeUnk(DU_TBL, lngX) = "Admin Officers  : Utility Menu -> Administrative Officer"
29650             Case "Location"
29660               arr_varDupeUnk(DU_TBL, lngX) = "Locations       : Utility Menu -> Locations"
29670             Case "m_REVCODE"
29680               arr_varDupeUnk(DU_TBL, lngX) = "Inc/Exp Codes   : Utility Menu -> Income / Expense Codes"
29690             Case "masterasset"
29700               arr_varDupeUnk(DU_TBL, lngX) = "Master Assets   : Asset Menu -> Add / Edit Assets"
29710             Case "RecurringItems"
29720               arr_varDupeUnk(DU_TBL, lngX) = "Recurring Items : Utility Menu -> Recurring Items"
29730             Case "Schedule"
29740               arr_varDupeUnk(DU_TBL, lngX) = "Schedules       : Utility Menu -> Fee Schedules"
29750             Case "Users"
29760               arr_varDupeUnk(DU_TBL, lngX) = "Users           : Utility Menu -> User Maintenance"
29770             End Select
29780           Next

29790           If lngTmp13 > 0& Then
                  ' ** The #DUPE identifier may show up in these tables:
                  ' **   adminofficer    : Utility Menu -> Administrative Officer
                  ' **   Location        : Asset Menu -> Locations
                  ' **   RecurringItems  : Utility Menu -> Recurring Items
                  ' **   Schedule        : Utility Menu -> Fee Schedules
                  ' **   m_REVCODE       : Utility Menu -> Income / Expense Codes
                  ' **   Users           : Utility Menu -> User Maintenance
                  ' **   account         : View / Add Account
29800             strTmp04 = "Duplicate entries were found in the following tables." & vbCrLf
29810             strTmp04 = strTmp04 & "Please review the entries for accuracy."
29820             For lngX = 0& To (lngDupeUnks - 1&)
29830               If arr_varDupeUnk(DU_TYP, lngX) = "DUP" Then
29840                 If InStr(strTmp04, (vbCrLf & arr_varDupeUnk(DU_TBL, lngX))) = 0 Then
29850                   strTmp04 = strTmp04 & vbCrLf & arr_varDupeUnk(DU_TBL, lngX)
29860                 End If
29870               End If
29880             Next
29890             strSummaryMsg = strSummaryMsg & vbCrLf & vbCrLf & strTmp04
29900           End If

29910           If lngTmp14 > 0& Then
                  ' ** The UNKNOWN identifier may show up in these tables:
                  ' **   adminofficer    : Utility Menu -> Administrative Officer
                  ' **   Schedule        : Utility Menu -> Fee Schedules
                  ' **   masterasset     : Asset Menu -> Add / Edit Assets
29920             strTmp05 = "Unknown missing or miss-named entries were found in the following tables." & vbCrLf
29930             strTmp05 = strTmp05 & "Please review the entries for accuracy."
29940             For lngX = 0& To (lngDupeUnks - 1&)
29950               If arr_varDupeUnk(DU_TYP, lngX) = "UNK" Then
29960                 If InStr(strTmp05, (vbCrLf & arr_varDupeUnk(DU_TBL, lngX))) = 0 Then
29970                   strTmp05 = strTmp05 & vbCrLf & arr_varDupeUnk(DU_TBL, lngX)
29980                 End If
29990               End If
30000             Next
30010             strSummaryMsg = strSummaryMsg & vbCrLf & vbCrLf & strTmp05
30020           End If

30030         End If

30040         If strTruncatedFields <> vbNullString Then
30050           strSummaryMsg = strSummaryMsg & vbCrLf & vbCrLf & strTruncatedFields
30060         End If

30070         Set dbsLoc = CurrentDb
30080         With dbsLoc
30090           Set rst = .OpenRecordset("tblVersion_Conversion", dbOpenDynaset, dbConsistent)
30100           With rst
30110             .FindFirst "[vercnv_id] = " & CStr(lngVerCnvID)
30120             If .NoMatch = False Then
30130               .Edit
30140               ![vercnv_note] = strSummaryMsg
30150               ![vercnv_datemodified] = Now()
30160               .Update
30170             End If
30180             .Close
30190           End With
30200           .Close
30210         End With

30220         .Status3 = strSummaryMsg
30230         .Status3.Visible = True
30240         DoEvents
30250         .cmdCancel.Enabled = True  ' ** Which now says 'Continue'. (See above)
30260         .cmdCancel.SetFocus
30270         DoEvents

30280         strTmp04 = CurrentAppPath  ' ** Module Function: modFileUtilities.
30290         If Left(strTmp04, 37) = Left(gstrDir_Dev, 37) Then        ' ** It's one of my test directories.
30300           If Left(strTmp04, 47) = gstrDir_Dev Then                 ' ** If it's in my main test directory...
                  ' ** "C:\VictorGCS_Clients\TrustAccountant\NewWorking"  '## OK
30310             strPath = (CurrentAppPath & LNK_SEP & gstrDir_Convert)  ' ** Module Function: modFileUtilities.
30320           ElseIf Parse_File(Left(strTmp04, 44)) = "NewDemo" Then  ' ** Module Function: modFileUtilities.
30330             strPath = (CurrentAppPath & LNK_SEP & gstrDir_Convert)  ' ** Module Function: modFileUtilities.
30340           ElseIf Parse_File(Left(strTmp04, 44)) = "Clients" And _
                    Right(strTmp04, 27) = "Delta Data\Trust Accountant" Then
30350             strPath = (CurrentAppPath & LNK_SEP & "Database" & LNK_SEP & gstrDir_Convert)
30360           Else
30370             Beep
30380             MsgBox "Where am I?", vbCritical + vbOKOnly, "Where Am I?"
30390           End If
30400         Else                                                     ' ** Otherwise...
30410           strPath = (gstrTrustDataLocation & gstrDir_Convert)     ' ** gstrTrustDataLocation INCLUDES FINAL BACKSLASH!
30420         End If
30430         Set fso = CreateObject("Scripting.FileSystemObject")
30440         Set fstx = fso.CreateTextFile(strPath & LNK_SEP & "ConvLog.txt", True)  ' ** {filename} {overwrite} {unicode}
30450         fstx.Write strSummaryMsg
30460         fstx.Close
30470         Set fsfl = Nothing
30480         Set fso = Nothing

30490       ElseIf gintConvertResponse < 0 Then
              ' ** Other controls have not been moved from their Progress position!
30500         If .PathFile_TrustData.Height <> 285& Then .PathFile_TrustData.Height = 285&
30510         If .PathFile_TrustArchive.Height <> 285& Then .PathFile_TrustArchive.Height = 285&
30520         .PathFile_TrustData_box.Height = (.PathFile_TrustData_box.Height - (15& * lngTpp))
30530         .PathFile_TrustArchive_box.Height = (.PathFile_TrustArchive_box.Height - (15& * lngTpp))
30540         .PathFile_TrustArchive_box.Top = (.PathFile_TrustArchive_box.Top - (15& * lngTpp))
30550         .PathFile_TrustArchive_box2.Top = (.PathFile_TrustArchive_box2.Top - (15& * lngTpp))
30560         .PathFile_TrustArchive.Top = (.PathFile_TrustArchive.Top - (15& * lngTpp))
30570         .PathFile_TrustArchive_lbl.Top = (.PathFile_TrustArchive_lbl.Top - (15& * lngTpp))
30580         .PathFile_TrustArchive_hline01.Top = (.PathFile_TrustArchive_hline01.Top - (15& * lngTpp))
30590         .PathFile_TrustArchive_hline02.Top = (.PathFile_TrustArchive_hline02.Top - (15& * lngTpp))
30600         .PathFile_TrustArchive_hline03.Top = (.PathFile_TrustArchive_hline03.Top - (15& * lngTpp))
30610         .PathFile_TrustArchive_vline01.Top = (.PathFile_TrustArchive_vline01.Top - (15& * lngTpp))
30620         .PathFile_TrustArchive_vline02.Top = (.PathFile_TrustArchive_vline02.Top - (15& * lngTpp))
30630         .PathFile_TrustArchive_vline03.Top = (.PathFile_TrustArchive_vline03.Top - (15& * lngTpp))
30640         .PathFile_TrustArchive_vline04.Top = (.PathFile_TrustArchive_vline04.Top - (15& * lngTpp))
30650         .Status1Cnt_lbl.Visible = False
30660         .Status1Of_lbl.Visible = False
30670         .Status1Tot_lbl.Visible = False
30680         lngOff2 = .Status3.Left - .Status3_lbl.Left
30690         .Status3.Top = ((.PathFile_TrustArchive.Top + .PathFile_TrustArchive.Height) + 180&)
30700         .Status3_lbl.Top = .Status3.Top
30710         .Status3.Left = (119& * lngTpp)
30720 On Error Resume Next
30730         .Status3_lbl.Left = .Status3.Left - lngOff2
30740         If ERR.Number <> 0 Then
30750 On Error GoTo ERRH
                ' ** 2100  ' ** The control or subform control is too large for this location.
30760           .Status3_lbl.Left = 0&
30770         Else
30780 On Error GoTo ERRH
30790         End If
30800         .Status3.Width = 7200&  ' ** 5"
30810         .Status3.Height = (.ProgBar_box.Top - .Status3.Top) - 120&
30820         .Status1_lbl.Visible = False
30830         .Status2_lbl.Visible = False
30840         intPos01 = InStr(strFailMsg, vbCrLf & vbCrLf)
30850         If intPos01 > 0 Then
30860           Do While intPos01 > 0
30870             strFailMsg = Left(strFailMsg, (intPos01 + 1)) & Mid(strFailMsg, (intPos01 + 4))
30880             intPos01 = InStr(strFailMsg, vbCrLf & vbCrLf)
30890           Loop
30900         End If
30910         .Status3 = strFailMsg
30920         .Status3.Visible = True
30930       End If
30940     End With

30950     DoCmd.SelectObject acForm, FRM_CNV_STATUS, False

30960   End Select

EXITP:
30970   Set frm = Nothing
30980   Set rst = Nothing
30990   Set fstx = Nothing
31000   Version_Status = blnRetVal
31010   Exit Function

ERRH:
31020   DoCmd.Hourglass False
31030   blnRetVal = False
31040   Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
31050   Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
31060   Case Else
31070     lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
31080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31090   End Select
31100   Resume EXITP

End Function

Public Sub ErrInfo_Set(lngENum As Long, lngELine As Long, strEDesc As String)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "ErrInfo_Set"

31210   lngErrNum = lngENum
31220   lngErrLine = lngELine
31230   strErrDesc = strEDesc

EXITP:
31240   Exit Sub

ERRH:
31250   DoCmd.Hourglass False
31260   Select Case ERR.Number
        Case Else
31270     lngErrNum = ERR.Number: lngErrLine = Erl: strErrDesc = ERR.description
31280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31290   End Select
31300   Resume EXITP

End Sub
