Attribute VB_Name = "modGlobConst"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modGlobConst"

'VGC 10/27/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const IsDemo = 0  ' ** 0 = new/upgrade; -1 = demo.
' ** Also in:
' **   modSecurityFunctions
' **   zz_mod_DatabaseDocFuncs
' **   zz_mod_MDEPrepFuncs

Public gstrTrustDataLocation As String  ' ** INCLUDES FINAL BACKSLASH!
Public gstrTrustAuxLocation As String

' ################################################################################
' ## Global Switches:
' ## ================
' ##   These global switches may be changed to variables, or vice versa, so that
' ##   they can be turned on and off at various places as the program runs.
' ##
' ## App Background Switch.
Public gblnDev_NoAppBackground As Boolean  ' ** Defaults to False.
' ## True: Don't open frmMenu_Background.
' ## False: Load frmMenu_Background at startup.
' ## We'll see whether this works or not!
' ##
' ## Error Switch.
Public Const gblnDev_NoErrHandle As Boolean = False
' ## True: Stop on the error line.
' ## False: Use ERRH error handling.
' ## Whether to stop at error, or use standard error handler during conversion.
' ##
' ## Developer Switch.
Public Const gblnDev_NoDispName As Boolean = True  ' ** Turn them off for now.
' ## True: Don't show form name in form's caption.
' ## False: Include form name in form's caption, within parentheses.
' ## So you can tell which form is showing.
' ##
' ## Debug Switch.
Public Const gblnDev_Debug As Boolean = False
' ## True: Redirect certain commands; *TEST*.
' ## False: Run as end-user.
' ## Mostly used to redirect report print commands.
' ##
' ################################################################################

Public Const LNK_IDENT As String = "DATABASE="
Public Const LNK_SEP As String = "\"

' ** The user name of the developer currently working on Trust.mdb.
' ** If used in a field, replace with "VictorC"
Public Const gstrDevUserName      As String = "VictorC"
#If IsDemo Then
  Public Const gstrDir_Dev        As String = "C:\VictorGCS_Clients\TrustAccountant\NewDemo"
#Else
  Public Const gstrDir_Dev        As String = "C:\VictorGCS_Clients\TrustAccountant\NewWorking"
#End If
Public Const gstrDir_DevClient    As String = "C:\VictorGCS_Clients\TrustAccountant\Clients"
Public Const gstrDir_Def          As String = "C:\Program Files\Delta Data\Trust Accountant"
Public Const gstrDir_Def64        As String = "C:\Program Files (x86)\Delta Data\Trust Accountant"
Public Const gstrDir_PricingOld   As String = "C:\My Documents\EVP Systems\EstateVal\"
Public Const gstrDir_Pricing      As String = "Pricing"
Public Const gstrDir_Convert      As String = "Convert_New"
Public Const gstrDir_ConvertEmpty As String = "Convert_Empty"
Public Const gstrDir_DevEmpty     As String = "EmptyDatabase"
Public Const gstrDir_DevDemo      As String = "DemoDatabase"
Public Const gstrDir_DevVer       As String = "PreviousVersionDBs"
Public gblnIsAccess2007 As Boolean
Public gblnIsAccess2010 As Boolean

' ** Security variables.
'Public gblnFirstOpen As Boolean
Public gblnBadSec As Boolean
Public gblnBadLink As Boolean
Public gblnCompact As Boolean
Public gblnAdmin As Boolean
Public gblnPricingAllowed  As Boolean  ' ** Is the pricing module allowed to be used.
Public gblnBeenToBackup As Boolean
Public gdbsDBLock As DAO.Database
Public gblnSetFocus As Boolean         ' ** Used to grab focus to Form A when Form B is closed.
Public gblnDemo As Boolean
Public gblnChangeDB As Boolean         ' ** To figure out if we want to change the backend database.
Public gblnSwitchTo As Boolean         ' ** Used to switch between Classic and Columnar Journal views.
Public gblnReportClose As Boolean
Public glngInstance As Long            ' ** Miscellaneous.

Public gstrActNo As String             ' ** For hiding transactions.
Public gstrReturningForm As String     ' ** For setting focus to button just used.
Public gblnClosing As Boolean          ' ** To let required-field checks know not to hang things up (e.g., Cancel = -1).
Public gblnDeleting As Boolean         ' ** Ditto!
Public glngPostingDateID As Long       ' ** Used with PostingDate and tblCalendar_Staging.

' ** Progress bar variables.
Public gdblPBar_MaxWidth As Double
Public gdblPBar_Steps As Double
Public gdblPBar_Increment As Double
Public gdblPBar_CurWidth As Double
Public gdblPBar_ThisStep As Double
'Public gctlPBar_Bar As Access.Rectangle
Public gctlPBar_Box1 As Access.Rectangle
Public gctlPBar_Box2 As Access.Rectangle
Public gctlPBar_Lbl As Access.Label
'gctlPBar  -> gctlPBar_Bar
'gctlPBox  -> gctlPBar_Box1
'gctlPBox2 -> gctlPBar_Box2
'gctlPLbl  -> gctlPBar_Lbl

'Public gctlPBar_Bar_2 As Access.Rectangle
Public gctlPBar_Box1_2 As Access.Rectangle
Public gctlPBar_Box2_2 As Access.Rectangle
Public gctlPBar_Lbl_2 As Access.Label
'gctlPBar_2  -> gctlPBar_Bar_2
'gctlPBox_2  -> gctlPBar_Box1_2
'gctlPBox2_2 -> gctlPBar_Box2_2
'gctlPLbl_2  -> gctlPBar_Lbl_2

' ** Used for the GoToReport feature on frmReportList.
Public gblnGoToReport As Boolean
Public gblnGoToReportMsg As Boolean
Public Const GTR_WAIT As Long = 1500&  ' ** Wait time on menus (1000 = 1 sec.).
Public garr_varGoToReport() As Variant
Public Const GTR_ELEMS As Integer = 28  ' ** Array's first-element UBound().
Public Const GTR_RLRID As Integer = 0
Public Const GTR_RID   As Integer = 1
Public Const GTR_RNAM  As Integer = 2
Public Const GTR_RCAP  As Integer = 3
Public Const GTR_ORD   As Integer = 4
Public Const GTR_FRM1  As Integer = 5
Public Const GTR_CTL1  As Integer = 6
Public Const GTR_INV1  As Integer = 7
Public Const GTR_FOC1  As Integer = 8
Public Const GTR_FRM2  As Integer = 9
Public Const GTR_CTL2  As Integer = 10
Public Const GTR_INV2  As Integer = 11
Public Const GTR_FOC2  As Integer = 12
Public Const GTR_FRM3  As Integer = 13
Public Const GTR_CTL3  As Integer = 14
Public Const GTR_INV3  As Integer = 15
Public Const GTR_FOC3  As Integer = 16
Public Const GTR_FRM4  As Integer = 17
Public Const GTR_CTL4  As Integer = 18
Public Const GTR_INV4  As Integer = 19
Public Const GTR_FOC4  As Integer = 20
Public Const GTR_FRM5  As Integer = 21
Public Const GTR_CTL5  As Integer = 22
Public Const GTR_INV5  As Integer = 23
Public Const GTR_FOC5  As Integer = 24
Public Const GTR_FRM6  As Integer = 25
Public Const GTR_CTL6  As Integer = 26
Public Const GTR_INV6  As Integer = 27
Public Const GTR_FOC6  As Integer = 28

' ** Constants used with SizeChecks() function, modStartupFuncs.
Public Const SZ_OK   As Integer = 0  ' ** All's well.
Public Const SZ_COMP As Integer = 1  ' ** TrustDta.mdb too big.
Public Const SZ_RECS As Integer = 2  ' ** A table has too many records.
Public Const SZ_ERR  As Integer = 3  ' ** Unexpected error.

' ** Warning constants.
Public Const glngWarnSize As Long = 104857600  ' ** Warn at 100MB data file size (1024 * 1024 * 100).
Public Const glngWarnRecs As Long = 300000     ' ** Warn at 300,000 records in any one table.

Public Const gintShareFaceDecimals As Integer = 4
Public Const gstrRegKeyName        As String = "Trust Accountant"

Public Const gstrExt_AppDev            As String = "mdb"
Public Const gstrExt_AppRun            As String = "mde"
Public Const gstrExt_AppSec            As String = "mdw"
Public Const gstrFile_App              As String = "Trust"
Public Const gstrFile_DataName         As String = "TrustDta.mdb"
Public Const gstrFile_DataLockfile     As String = "TrustDta.ldb"
Public Const gstrFile_ArchDataName     As String = "TrstArch.mdb"
Public Const gstrFile_ArchDataLockfile As String = "TrstArch.ldb"
Public Const gstrFile_AuxDataName      As String = "TrustAux.mdb"
Public Const gstrFile_AuxDataLockfile  As String = "TrustAux.ldb"
Public Const gstrFile_SecurityName     As String = "TrustSec.mdw"
Public Const gstrFile_INI              As String = "DDTrust.ini"
Public Const gstrFile_LIC              As String = "TA.lic"
Public Const gstrFile_Icon             As String = "Trust.ico"
Public Const gstrFile_Manual           As String = "TA Manual.pdf"
Public Const gstrFile_RestoreDataName  As String = "TrstRest.mdb"
Public Const gstrFile_RePostDataName   As String = "TAJrnTmp.mdb"
'Public Const gstrFile_SqlCatalogCSV    As String = "SQL_Catalog.csv"
'Public Const gstrFile_TblCatalogCSV    As String = "CSV_Catalog.csv"
Public Const gstrFile_DevNorthForkDta  As String = "TrustDta_WmBJohnson.mdb"
Public Const gstrFile_DevNorthForkArch As String = "TrstArch_WmBJohnson.mdb"
Public Const gstrFile_DevHintonDta     As String = "TrustDta_Hinton.mdb"
Public Const gstrFile_DevHintonArch    As String = "TrstArch_Hinton.mdb"
'Public Const gstrFile_PricingFileOld   As String = "TAPricing.txt"
Public Const gstrFile_ArchiveLog       As String = "TALog.txt"
Public Const gstrFile_ConvertLog       As String = "ConvLog.txt"
Public Const gstrTable_LedgerArchive   As String = "LedgerArchive_Backup"

' ** Public Error Messages.
Public Const gstrNoPermission As String = "You do not have permission."

Public Const RECUR_I_TO_P As String = "Transferred Income Cash to Principal Cash"
Public Const RECUR_I_TO_P_ID As Long = 1&
Public Const RECUR_P_TO_I As String = "Transferred Principal Cash to Income Cash"
Public Const RECUR_P_TO_I_ID As Long = 2&

Public Const WIDEN_MULT As Long = 16&    ' ** Standard widen/shorten multiplier of the screen's Twips Per Pixel.
Public Const WIDEN_MAX As Long = 21600&  ' ** Cap the forms width at 15 in.

' ** Number of rows to show on continuous forms when sizing window.
Public Const RZ_WINROWS As Long = 10&
Public Const RZ_TWIPSPERINCH As Long = 1440&

' ** Check Register constants.
Public Const CHKREG_NUM  As Long = 1&
Public Const CHKREG_PAID As Long = 2&

' ** Return code.
Public Const RET_ERR As String = "#ERROR"

' ** Company Information variables (CoInfo, CompanyInformation).
Public gstrCo_Name As String
Public gstrCo_Address1 As String
Public gstrCo_Address2 As String
Public gstrCo_City As String
Public gstrCo_State As String
Public gstrCo_Zip As String
Public gstrCo_Country As String
Public gstrCo_PostalCode As String
Public gstrCo_Phone As String
Public gstrCo_InfoBlock As String                ' ** Company information formated to use in report headings.

Public grstPostingDate As DAO.Recordset
Public gblnMessage As Boolean
Public gblnSignal As Boolean
Public gblnTimer As Boolean
Public gdatStartDate As Date                  ' ** Used to pass data from menus to reports.
Public glngStartDateLong As Long
Public gdatEndDate As Date                    ' ** Used to pass data from menus to reports.
Public glngEndDateLong As Long
Public gstrAccountNo As String                ' ** Used to pass data from menus to reports.
Public gstrAccountName As String              ' ** Used to pass data from menus to reports.
Public gblnLegalName As Boolean               ' ** Used to pass data from menus to reports.
Public gblnForeignCurrencies As Boolean
Public gblnHasForEx As Boolean                ' ** At least one foreign currency Ledger record or asset in Trust Accountant.
Public gblnHasForExThis As Boolean            ' ** This account has foreign currency.
Public glngCurrID As Long
Public glngMonthID As Long
Public gdatMarketDate As Date
Public glngAssetNo As Long

Public gblnUseReveuneExpenseCodes As Boolean  ' ** Used to pass data from menus to reports.
Public gblnCombineAssets  As Boolean          ' ** Used to pass from statement form to asset list report.

' ** Startup options.
Public gintDefaultOpenMode As Integer  ' ** These are for users' original default settings.
Public gintDefaultRecordLocking As Integer
Public gblnConfirmRecordChanges As Boolean
Public gblnConfirmActionQueries As Boolean
Public gblnConfirmDocumentDeletions As Boolean
Public gblnAutoCompact As Boolean
Public gblnShowWindowsInTaskbar As Boolean
Public gblnShowHiddenObjects As Boolean
Public gstrDefaultDatabaseDirectory As String
'Public gblnPerformNameAutoCorrect As Boolean
'Public gblnTrackNameAutoCorrectInfo As Boolean
'Public gblnLogNameAutoCorrectChanges As Boolean

Public gblnCBarVis As Boolean

Public gblnIncomeTaxCoding As Boolean
Public gblnRevenueExpenseTracking As Boolean
Public gblnAccountNoWithType As Boolean
Public gblnSeparateCheckingAccounts As Boolean
Public gblnTabCopyAccount As Boolean
Public gblnLinkRevTaxCodes As Boolean
Public gblnSpecialCapGainLoss As Boolean
Public gintSpecialCapGainLossOpt As Integer
Public gblnLocSuggest As Boolean

Public gstrJournalUser As String
Public gblnSingleUser As Boolean              ' ** Used for Pre-Post Journal report.
Public gblnLocalData As Boolean               ' ** Used with frmJournal_Columns.
Public glngUserCntLedger As Long              ' ** Used with above to cut querying time.
Public gblnPrintAll As Boolean
'Public gblnMaxDateAlreadyPresent As Boolean
'Public gblnInterimReport As Boolean

Public glngJournalForm As Long

' ** TransactionForm family variables.
Public JRN_DIV As Long
Public JRN_INT As Long
Public JRN_PUR As Long
Public JRN_SAL As Long
Public JRN_MIS As Long

Public gblnDividendChanged As Boolean
Public gblnDividendValidated As Boolean
Public gstrDividendType As String
Public gstrDividendAsset As String
Public gstrDividendShareFace As String
Public gstrDividendAccountNumber As String
Public gstrDividendICash As String
Public gstrDividendPerShare As String

Public gblnInterestChanged As Boolean
Public gblnInterestValidated As Boolean
Public gstrInterestType As String
Public gstrInterestAsset As String
Public gstrInterestShareFace As String
Public gstrInterestAccountNumber As String
Public gstrInterestICash As String

Public gblnMiscChanged As Boolean
Public gblnMiscValidated As Boolean
Public gstrMiscType As String
Public gstrMiscAccountNumber As String
Public gstrMiscICash As String
Public gstrMiscPCash As String
'Public gstrMiscCost As String

Public gblnPurchaseChanged As Boolean
Public gblnPurchaseValidated As Boolean
Public gstrPurchaseType As String
Public gstrPurchaseAsset As String
Public gstrPurchaseShareFace As String
Public gstrPurchaseAccountNumber As String
Public gstrPurchaseICash As String
Public gstrPurchasePCash As String
Public gstrPurchaseCost As String
Public gblnIsLiability As Boolean

Public gblnSaleChanged As Boolean
Public gblnSaleValidated As Boolean
Public gstrSaleType As String
Public gstrSaleAsset As String
Public gstrSaleShareFace As String
Public gstrSaleAccountNumber As String
Public gstrSaleICash As String
Public gstrSalePCash As String
Public gstrSaleCost As String

'Public gblnCashChanged As Boolean
'Public gblnCashValidated As Boolean
'Public gstrCashType As String
'Public gstrCashAsset As String
'Public gstrCashShareFace As String
'Public gstrCashAccountNumber As String
'Public gstrCashPCash As String
'Public gstrCashCost As String

' ** Used for modQueryFunctions1.
Public gstrReportQuerySpec As String
Public gstrReportCallingForm As String
Public gstrFormQuerySpec As String

' ** Used for Court Reports.
Public gdblCrtRpt_CA_COHBeg As Double
Public gdblCrtRpt_CA_COHEnd As Double
Public gdblCrtRpt_CA_InvestChange As Double
Public gdblCrtRpt_CA_InvestInfo As Double
Public gdblCrtRpt_CA_POHBeg As Double
Public gdblCrtRpt_CA_POHEnd As Double
Public gvarCrtRpt_FL_SpecData As Variant
Public gstrCrtRpt_NY_AccountNo As String
Public gdatCrtRpt_NY_DateEnd As Date
Public gdatCrtRpt_NY_DateStart As Date
Public gcurCrtRpt_NY_ICash As Currency
Public gcurCrtRpt_NY_IncomeBeg As Currency
Public gcurCrtRpt_NY_InputAmtForm As Currency
Public gcurCrtRpt_NY_InputNew As Currency
Public gstrCrtRpt_NY_InputTitle As String
Public gblnCrtRpt_NY_InvIncChange As Boolean
Public gstrCrtRpt_Account As String
Public gstrCrtRpt_CashAssets_Beg As String
Public gstrCrtRpt_CashAssets_End As String
Public gdblCrtRpt_CostTot As Double
Public gdblCrtRpt_IncTot As Double
Public gstrCrtRpt_NetIncome As String
Public gstrCrtRpt_NetLoss As String
Public gstrCrtRpt_Ordinal As String
Public gstrCrtRpt_Period As String
Public gdblCrtRpt_PrinTot As Double
Public gstrCrtRpt_Version As String
Public gblnCrtRpt_Zero As Boolean
Public gblnCrtRpt_ZeroDialog As Boolean
'"gblnInvestedIncomeChangeNY" -> "gblnCrtRpt_NY_InvIncChange"
'"gcurAmountInputFormNY" -> "gcurCrtRpt_NY_InputAmtForm"
'"gcurICashNY" -> "gcurCrtRpt_NY_ICash"
'"gcurIncomeAtBeginNY" -> "gcurCrtRpt_NY_IncomeBeg"
'"gcurNewInputNY" -> "gcurCrtRpt_NY_InputNew"
'"gdatDateEndNY" -> "gdatCrtRpt_NY_DateEnd"
'"gdatDateStartNY" -> "gdatCrtRpt_NY_DateStart"
'"gstrAccountNoNY" -> "gstrCrtRpt_NY_AccountNo"
'"gstrInputTitleNY" -> "gstrCrtRpt_NY_InputTitle"
'"gdblCA_Beg" -> "gdblCrtRpt_CA_COHBeg"
'"gdblCA_End" -> "gdblCrtRpt_CA_COHEnd"
'"gdblInvestChange" -> "gdblCrtRpt_CA_InvestChange"
'"gdblInvestInfo" -> "gdblCrtRpt_CA_InvestInfo"
'"gdblPOH_Beg" -> "gdblCrtRpt_CA_POHBeg"
'"gdblPOH_End" -> "gdblCrtRpt_CA_POHEnd"
'"gvarSpecialFloridaData" -> "gvarCrtRpt_FL_SpecData"

Public glngTaxCode_Distribution As Long

' ** Array: arr_varPrintRpt().
Public glngPrintRpts As Long, garr_varPrintRpt() As Variant
Public Const PR_ELEMS As Integer = 4  ' ** Array's first-element UBound().
Public Const PR_ACTNO As Integer = 0
Public Const PR_ALIST As Integer = 1
Public Const PR_TRANS As Integer = 2
Public Const PR_SUMRY As Integer = 3

' *******************************************************************************
' ** SRSEncrypt.dll was written and compiled by Duane Johnson, Minneapolis, MN
' *******************************************************************************
Public Declare Function encrypt Lib "SRSencrypt.dll" (ByVal text As String) As String

' ** VbDataType enumeration:  (my own)
'Public Const vbArrayByte         As Long = 8209&

' ** DbDataType enumeration:  (my own)
Public Const dbUnknown           As Long = 0&

' ** AcControlType enumeration:  (my own)
'Public Const acNone              As Long = 99&
'Public Const acDatasheetColumn   As Long = 115&
'Public Const acEmptyCell         As Long = 127&
'Public Const acWebBrowser        As Long = 128&
'Public Const acNavigationControl As Long = 129&
'Public Const acNavigationButton  As Long = 130&

' ** AcColumnWidth enumeration: (my own)  ' ** Not currently used!
'Public Const acColumnWidthHide As Integer = 0      ' ** Hides the column.
'Public Const acColumnWidthDefault As Integer = -1  ' ** Sizes the column to the default width. (Default)
'Public Const acColumnWidthToFit As Integer = -2    ' ** Sizes the column to fit the size of the visible text.

' ** AcObjectType enumeration:  (my own)
Public Const acSQL     As Long = -2&
Public Const acNothing As Long = -3&

' ** AcCurrentView enumeration:  (Access 2007)
'Public Const acCurViewNone         As Integer = -1  ' ** The object is not loaded.
Public Const acCurViewDesign       As Integer = 0   ' ** The object is in Design view.
Public Const acCurViewFormBrowse   As Integer = 1   ' ** The object is in Form view.
Public Const acCurViewDatasheet    As Integer = 2   ' ** The object is in Datasheet view.
'Public Const acCurViewPivotTable   As Integer = 3   ' ** The object is in PivotTable view.
'Public Const acCurViewPivotChart   As Integer = 4   ' ** The object is in PivotChart view.
'Public Const acCurViewPreview      As Integer = 5   ' ** The object is in Print Preview.
'Public Const acCurViewReportBrowse As Integer = 6   ' ** The object is in Report view.
'Public Const acCurViewLayout       As Integer = 7   ' ** The object is in Layout view.

' ** AcBackStyle enumeration:
Public Const acBackStyleTransparent As Integer = 0  ' ** The control has its interior color set by the BackColor property. (Default for all controls except option group)
Public Const acBackStyleNormal      As Integer = 1  ' ** The control is transparent. The color of the form or report behind the control is visible. (Default for option group)
 

' ** AcTextAlign enumeration:
'Public Const acTextAlignGeneral As Integer = 0  ' ** The text aligns to the left; numbers and dates align to the right. (Default)
Public Const acTextAlignLeft    As Integer = 1  ' ** The text, numbers, and dates align to the left.
Public Const acTextAlignCenter  As Integer = 2  ' ** The text, numbers, and dates are centered.
Public Const acTextAlignRight   As Integer = 3  ' ** The text, numbers, and dates align to the right.

' ** AcObjState enumeration:
' **   0  acObjStateClosed  Not open. (my own)
' **   1  acObjStateOpen    Open.
' **   2  acObjStateDirty   Design changed but not saved.
' **   3                    Design changed but not saved!  ' ** Additive: open and changed but not saved.
' **   4  acObjStateNew     New.
' **   5                    New!                           ' ** Additive: open and new.
Public Const acObjStateClosed As Integer = 0  ' ** (my own)

' ** VbDataType enumeration:
'Public Const vbUndeclared As Integer = -3  ' ** (my own)

' ** AcSpecialEffect enumeration:
Public Const acSpecialEffectFlat     As Long = 0
'Public Const acSpecialEffectRaised   As Long = 1
Public Const acSpecialEffectSunken   As Long = 2
'Public Const acSpecialEffectEtched   As Long = 3
'Public Const acSpecialEffectShadowed As Long = 4
'Public Const acSpecialEffectChiseled As Long = 5

' ** AcBorderStyle enumeration:  (my own)
Public Const acBorderStyleTransparent As Long = 0
Public Const acBorderStyleSolid       As Long = 1
'Public Const acBorderStyleDash        As Long = 2
'Public Const acBorderStyleShortDash   As Long = 3
'Public Const acBorderStyleDot         As Long = 4
'Public Const acBorderStyleSparseDot   As Long = 5
'Public Const acBorderStyleDashDot     As Long = 6
'Public Const acBorderStyleDashDotDot  As Long = 7
'Public Const acBorderStyleDoubleSolid As Long = 8

' ** AcDatabaseType enumeration:  (my own)  ' ** Not currently used!
'Public Const acDbTypeAccess As String = "Microsoft Access"
'Public Const acDbTypeDB5    As String = "dBase 5#"
'Public Const acDbTypeDB3    As String = "dBase III"
'Public Const acDbTypeDB4    As String = "dBase IV"
'Public Const acDbTypeJet2   As String = "Jet 2.x"
'Public Const acDbTypeJet3   As String = "Jet 3.x"
'Public Const acDbTypePDX3   As String = "Paradox 3.x"
'Public Const acDbTypePDX4   As String = "Paradox 4.x"
'Public Const acDbTypePDX5   As String = "Paradox 5.x"
'Public Const acDbTypePDX7   As String = "Paradox 7.x"
'Public Const acDbTypeWSS    As String = "WSS"
'Public Const acDbTypeODBC   As String = "ODBC Database"

' ** dbQueryDefType enumeration:  (my own)  ' ** Not currently used!
'Public Const dbQNothing    As Integer = -3  ' ** Not a query.

' ** DbConnect enumeration:  (my own)
Public Const dbCNothing    As Integer = -3  ' ** Not a table.
Public Const dbCAccess     As Integer = 0   ' ** Microsoft Access Table (not linked)
Public Const dbCJet        As Integer = 1   ' ** Microsoft Jet Database
'Public Const dbCdBASEIII   As Integer = 2   ' ** dBASE III
'Public Const dbCdBASEIV    As Integer = 3   ' ** dBASE IV
'Public Const dbCdBASE5     As Integer = 4   ' ** dBASE 5
'Public Const dbCParadox3   As Integer = 5   ' ** Paradox 3.x
'Public Const dbCParadox4   As Integer = 6   ' ** Paradox 4.x
'Public Const dbCParadox5   As Integer = 7   ' ** Paradox 5.x
'Public Const dbCExcel3     As Integer = 8   ' ** Microsoft Excel 3.0
'Public Const dbCExcel4     As Integer = 9   ' ** Microsoft Excel 4.0
'Public Const dbCExcel5     As Integer = 10  ' ** Microsoft Excel 5.0, Microsoft Excel 95
'Public Const dbCExcel8     As Integer = 11  ' ** Microsoft Excel 97
'Public Const dbCWK1        As Integer = 12  ' ** Lotus 1-2-3 WKS, WK1
'Public Const dbCWK3        As Integer = 13  ' ** Lotus 1-2-3 WK3
'Public Const dbCWK4        As Integer = 14  ' ** Lotus 1-2-3 WK4
'Public Const dbCHTMLImport As Integer = 15  ' ** HTML Import
'Public Const dbCHTMLExport As Integer = 16  ' ** HTML Export
'Public Const dbCCSV        As Integer = 17  ' ** Text
'Public Const dbCODBC       As Integer = 18  ' ** ODBC
'Public Const dbCExchange4  As Integer = 19  ' ** Microsoft Exchange
'Public Const dbCOutlook    As Integer = 20  ' ** Microsoft Outlook

' ** AcRibbonState enumeration:  (my own)
Public Const acRibbonStateNormalAbove  As Integer = 0  ' ** QAT Normal, above ribbon.
'Public Const acRibbonStateNormalBelow  As Integer = 1  ' ** QAT Normal, below ribbon.
Public Const acRibbonStateAutohide     As Integer = 4  ' ** QAT Autohide.

' ** AcForceNewPage enumeration: (my own)
Public Const acForceNewPageNone   As Integer = 0
'Public Const acForceNewPageBefore As Integer = 1
Public Const acForceNewPageAfter  As Integer = 2
'Public Const acForceNewPageBoth   As Integer = 3

' ** AcNewRowOrCol enumeration: (my own)  ' ** Not currently used!
'Public Const acNewRowOrColNone   As Integer = 0
'Public Const acNewRowOrColBefore As Integer = 1
'Public Const acNewRowOrColAfter  As Integer = 2
'Public Const acNewRowOrColBoth   As Integer = 3

' ** AcMultiSelect enumeration: (my own)  ' ** Not currently used!
'Public Const acMultiSelectNone     As Integer = 0
'Public Const acMultiSelectSimple   As Integer = 1
'Public Const acMultiSelectExtended As Integer = 2

' ** VbSystemColors enumeration:  ' ** Not currently used!
'Public Const vbActiveTitleBarText       As Long = -2147483639
'Public Const vbInactiveTitleBarText     As Long = -2147483629
'Public Const vbStaticBackground         As Long = -2147483623
'Public Const vbStaticText               As Long = -2147483622
'Public Const vbActiveTitleBarGradient   As Long = -2147483621
'Public Const vbInactiveTitleBarGradient As Long = -2147483620
'Public Const vbMenuHighlightFlat        As Long = -2147483619
'Public Const vbMenuBackgroundFlat       As Long = -2147483618

' ** AcDisplay enumeration:  (my own)  ' ** Not currently used!
'Public Const acDisplayAlways As Integer = 0
'Public Const acDisplayPrint  As Integer = 1
'Public Const acDisplayScreen As Integer = 2

Public CLR_HILITE As Long

' ** Color constants.
Public Const MY_CLR_BGE    As Long = 14215660  ' ** My standard beige, form BackColor: 236/233/216.
Public Const MY_CLR_MDBGE  As Long = 14544110  ' ** My standard medium beige: 238/236/221.
Public Const MY_CLR_LTBGE  As Long = 14872561  ' ** My standard light beige, 2nd color: 241/239/226.
Public Const MY_CLR_VLTBGE As Long = 15923449  ' ** My standard very light beige: 249/248/242.

Public Const CLR_POST_BG  As Long = 11442576  ' ** The gray background for frmMenu_Background.
Public Const CLR_NEW_BLUE As Long = 16314086  ' ** The light blue background for 'NEW BLUE' forms before PictureData.

Public Const CLR_AC07     As Long = 13603685  ' ** Medium blue, standard Access 2007 border color.
Public Const CLR_BLK      As Long = 0&        ' ** Black, just to be consistent with constants.
Public Const CLR_BLU      As Long = 16711680  ' ** Blue.
Public Const CLR_BLUGRY   As Long = 8404992   ' ** Slightly grayish blue, frmJournal_Columns special purpose button labels.
Public Const CLR_BRN      As Long = 7223      ' ** Brown
Public Const CLR_DKBLU    As Long = 8388608   ' ** Dark blue, for label BackColor.
Public Const CLR_DKBLU2   As Long = 10485760  ' ** Not as dark as above.
Public Const CLR_DKGRN    As Long = 16384     ' ** Dark, deep green.
Public Const CLR_DKGRY    As Long = 6052956   ' ** Dark gray, for ForeColor. (92/92/92)
Public Const CLR_DKGRY2   As Long = 4868682   ' ** Darker gray, for ForeColor.  (74/74/74)
Public Const CLR_DKGRY3   As Long = 7105644   ' ** Slightly lighter than CLR_DKGRY.  (108/108/108)
Public Const CLR_DKGRY4   As Long = 4079166   ' **   (62/62/62)
Public Const CLR_DKRED    As Long = 128&      ' ** Dark red.
Public Const CLR_GOLD1    As Long = 8442361   ' ** Delta Data gold.
Public Const CLR_GRN      As Long = 26112     ' ** Green.
Public Const CLR_GRY      As Long = 8421504   ' ** Gray, for BorderColor, BackColor on frmTransaction_Audit for printing. (128/128/128)
Public Const CLR_GRY2     As Long = 14410211  ' ** Gray, for BackColor, with black ForeColor. (227/225/219)
Public Const CLR_GRY3     As Long = 12632256  ' ** Gray,
Public Const CLR_GRY4     As Long = 9211020   ' ** Gray, very medium, used on frmAccountHideTrans2_Sub_Pick. (140/140/140)
Public Const CLR_GRY5     As Long = 8553090   ' ** Gray, label line for CLR_GRY4. (130/130/130)
Public Const CLR_HID      As Long = 8453888   ' ** Green, the .BackColor of form text boxes that aren't visible.
Public Const CLR_LTBLU    As Long = 16774128  ' ** Light blue, frmAccountHideTrans_Match highlight color.
Public Const CLR_LTBLU2   As Long = 12164479  ' ** New Access 2003 text box border color.
Public Const CLR_LTCYAN   As Long = 16777088  ' ** Light blue, for ForeColor on frmTransaction_Audit for printing.
Public Const CLR_LTGRN    As Long = 15138802  ' ** Light green, frmAccountHideTrans_Match highlight color.
Public Const CLR_LTGRY    As Long = 14737632  ' ** Light gray, for BackColor, with dark gray ForeColor. (224/224/224)
Public Const CLR_LTGRY2   As Long = 13158600  ' ** Very light gray, for some disabled ForeColor.  (200/200/200)
Public Const CLR_LTORNG   As Long = 7785976   ' ** Light orange.
Public Const CLR_LTPRP    As Long = 16771829  ' ** Light purple.
Public Const CLR_LTRED    As Long = 14013951  ' ** Light red, for BackColor when 'error'.
Public Const CLR_LTRED2   As Long = 15921919  ' ** Very light red.
Public Const CLR_LTTEAL   As Long = 16775920  ' ** Light teal, trying this out for locked records.
Public Const CLR_LTTEAL2  As Long = 16438855  ' ** Site Map label box borders.
Public Const CLR_LTTEAL3  As Long = 16774365  ' ** Background color for closed accounts
Public Const CLR_LTYEL    As Long = 8454143   ' ** Yellow, frmAccountHideTrans_Match highlight color.
Public Const CLR_OFFWHT   As Long = 16710908  ' ** Used with Tab Controls.
Public Const CLR_ORNG     As Long = 3389428   ' ** Orange, like the gold medallion.
Public Const CLR_RED      As Long = 255&      ' ** Red.
Public Const CLR_ROSE     As Long = 4210816   ' ** Rose, dark, dusty red.
Public Const CLR_TEAL     As Long = 8421440   ' ** The color used for all the notes and table legends.
Public Const CLR_VDKGRY   As Long = 3026478   ' ** Very dark gray, for menu buttons. (46/46/46)
Public Const CLR_VDKGRY2  As Long = 2631720   ' ** Very darker gray, for 7 Pt. labels. (40/40/40)
Public Const CLR_VLTBLU   As Long = 16776441  ' ** Very light blue.
Public Const CLR_VLTBLU2  As Long = 15129808  ' ** CLR_LTBLU2, but lighter.
Public Const CLR_VLTBLU3  As Long = 15986665  ' ** Very, very light blue border.
Public Const CLR_VLTGRN   As Long = 16056314  ' ** Very light green, foreign currency backcolor.
Public Const CLR_VLTGRY   As Long = 16119285  ' ** Very light gray, for BackColor, with dark gray ForeColor.
Public Const CLR_VLTPRP   As Long = 16774907  ' ** Very light purple.
Public Const CLR_VLTRED   As Long = 15921919  ' ** Very light red.
Public Const CLR_VLTROSE  As Long = 9342639   ' ** Very light rose foreground.
Public Const CLR_VLTTEAL  As Long = 16776697  ' ** Very Very light teal
Public Const CLR_VLTYEL   As Long = 12648446  ' ** Very light yellow.
Public Const CLR_WHT      As Long = 16777215  ' ** White

Public Const CLR_IEI_BLU     As Long = 16734553  ' ** Blue
Public Const CLR_IEI_BLU_DIS As Long = 14466750  ' ** Disabled Blue
Public Const CLR_IEE_RED     As Long = 6974207   ' ** Red
Public Const CLR_IEE_RED_DIS As Long = 12500690  ' ** Disable Red

Public Const WIN_CLR_DISF As Long = 10526880  '-2147483632  ' ** Dimmed (disabled) text; indicating a disabled control or form.
'  VGC 11/28/2011: current disabled forecolor 10526880, 160/160/160.
Public Const WIN_CLR_DISB As Long = -2147483633  ' ** 3-D face; Dimmed (disabled) BackColor.
'  VGC 11/28/2011: current disabled text box backcolor: 240/240/240
'  VGC 11/28/2011: (current disabled command button backcolor: -2147483637 244/244/244)
Public Const WIN_CLR_DISR As Long = 11775403     ' ** Disabled border color, and Trust Import dimmed-form BackColor.
'  VGC 11/28/2011: current disabled text box bordercolor: 171/173/179
'  VGC 11/28/2011: (current disabled command button border: 11907757 173/178/181)
Public Const WIN_CLR_DIM2 As Long = -2147483628  ' ** Dimmed (disabled) text highlight; 2nd, offset highlight color.
'  VGC 11/28/2011: current, and always, disabled shadow: 255/255/255
Public Const WIN_CLR_DIM3 As Long = 13359838     'NOT A WINDOWS COLOR!
'  old etched 1st color
Public Const WIN_CLR_3DDK As Long = -2147483627  ' ** Color of the dark shadow for three-dimensional display elements.
'  3rd color in box borders
Public Const WIN_CLR_3DLT As Long = -2147483626  ' ** Highlight color of three-dimensional display elements for edges that face the light source.
'  4th color in box borders
Public Const WIN_CLR_NORM As Long = 14215660
'  VGC 11/28/2011: current standard beige, form backcolor; see MY_CLR_BGE.

'Public Const WIN_CLR_NORM As Long = 14215660
'Public Const WIN_CLR_DIM  As Long = 10070188
'Public Const WIN_CLR_DIM2 As Long = 16777215
'Public Const WIN_CLR_3DLT As Long = 14872561
'Public Const WIN_CLR_3DDK As Long = 6582129

'Forms(0).Nav_hline03.BorderColor = CLR_AC07
'Forms(0).Nav_hline03.Top = (Forms(0).Controls(Subform_Get()).Top + Forms(0).Controls(Subform_Get()).Height)
'Forms(0).Nav_hline03.Left = Forms(0).Controls(Subform_Get()).Left
'Forms(0).Nav_hline03.Width = Forms(0).Controls(Subform_Get()).Width

'3026478 : 46/46/46    : CLR_VDKGRY   TOO CLOSE! ONE OR THE OTHER!
'2631720 : 40/40/40    : CLR_VDKGRY2  RETIRE THIS ONE!
'4079166 : 62/62/62    : CLR_DKGRY4
'4868682 : 74/74/74    : CLR_DKGRY2   ALL 'Remembers'
'7105644 : 108/108/108 : CLR_DKGRY3   ALL CURRENCY ALT GRPS!
